"""
MySQL persistence layer for SkillGauge AI.

The database is the primary store for users, skills, assessments and
reports. Excel is generated on demand FROM this data (see
services/excel_export.py) - it is a backup/export format, not the
source of truth.

Every other module should go through the functions in this file rather
than writing raw SQL elsewhere, so the schema/queries stay in one place.
"""
import json
from datetime import datetime

import mysql.connector
from mysql.connector import pooling

import config

_pool = None


def _get_pool():
    global _pool
    if _pool is None:
        _pool = pooling.MySQLConnectionPool(
            pool_name="skillup_pool",
            pool_size=5,
            host=config.MYSQL_HOST,
            port=config.MYSQL_PORT,
            database=config.MYSQL_DATABASE,
            user=config.MYSQL_USER,
            password=config.MYSQL_PASSWORD,
            autocommit=True,
        )
    return _pool


def get_connection():
    """Get a pooled connection. Caller is responsible for closing it."""
    return _get_pool().get_connection()


# =====================================================
# SCHEMA
# =====================================================
_SCHEMA = """
CREATE TABLE IF NOT EXISTS users (
    id INT AUTO_INCREMENT PRIMARY KEY,
    name VARCHAR(255) NULL,
    email VARCHAR(255) NOT NULL UNIQUE,
    role ENUM('candidate', 'admin', 'recruiter') NOT NULL DEFAULT 'candidate',
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
) ENGINE=InnoDB;

CREATE TABLE IF NOT EXISTS skills (
    id INT AUTO_INCREMENT PRIMARY KEY,
    user_id INT NOT NULL,
    skill_name VARCHAR(255) NOT NULL,
    category VARCHAR(100) NOT NULL DEFAULT 'Uncategorized',
    proficiency_level VARCHAR(50) NULL,
    source ENUM('manual', 'resume') NOT NULL DEFAULT 'manual',
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY (user_id) REFERENCES users(id) ON DELETE CASCADE
) ENGINE=InnoDB;

CREATE TABLE IF NOT EXISTS assessments (
    id INT AUTO_INCREMENT PRIMARY KEY,
    user_id INT NOT NULL,
    skills_snapshot TEXT NULL,
    questions_json LONGTEXT NULL,
    responses_json LONGTEXT NULL,
    score INT NOT NULL DEFAULT 0,
    max_score INT NOT NULL DEFAULT 25,
    accuracy DECIMAL(5,2) NOT NULL DEFAULT 0.00,
    ai_generated BOOLEAN NOT NULL DEFAULT TRUE,
    completion_status ENUM('completed', 'in_progress', 'abandoned') NOT NULL DEFAULT 'completed',
    started_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    completed_at TIMESTAMP NULL,
    FOREIGN KEY (user_id) REFERENCES users(id) ON DELETE CASCADE
) ENGINE=InnoDB;

CREATE TABLE IF NOT EXISTS reports (
    id INT AUTO_INCREMENT PRIMARY KEY,
    assessment_id INT NOT NULL,
    pdf_path VARCHAR(500) NULL,
    email_status ENUM('sent', 'failed', 'not_attempted') NOT NULL DEFAULT 'not_attempted',
    generated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY (assessment_id) REFERENCES assessments(id) ON DELETE CASCADE
) ENGINE=InnoDB;

CREATE TABLE IF NOT EXISTS jobs (
    id INT AUTO_INCREMENT PRIMARY KEY,
    title VARCHAR(255) NOT NULL,
    description TEXT NULL,
    required_skills TEXT NULL,
    department VARCHAR(100) NULL,
    embedding_json LONGTEXT NULL,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
) ENGINE=InnoDB;

CREATE TABLE IF NOT EXISTS mentors (
    id INT AUTO_INCREMENT PRIMARY KEY,
    name VARCHAR(255) NOT NULL,
    expertise TEXT NULL,
    bio TEXT NULL,
    embedding_json LONGTEXT NULL,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
) ENGINE=InnoDB;

CREATE TABLE IF NOT EXISTS learning_resources (
    id INT AUTO_INCREMENT PRIMARY KEY,
    title VARCHAR(255) NOT NULL,
    skill_tag VARCHAR(255) NULL,
    url VARCHAR(500) NULL,
    description TEXT NULL,
    embedding_json LONGTEXT NULL,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
) ENGINE=InnoDB;
"""


def init_db():
    """Create all tables if they don't already exist. Safe to call every startup."""
    conn = get_connection()
    try:
        cursor = conn.cursor()
        for statement in _SCHEMA.strip().split(";"):
            statement = statement.strip()
            if statement:
                cursor.execute(statement)
        cursor.close()
    finally:
        conn.close()


# =====================================================
# USERS
# =====================================================
def get_or_create_user(email, name=None, role="candidate"):
    """Return the user id for this email, creating the row if needed."""
    conn = get_connection()
    try:
        cursor = conn.cursor()
        cursor.execute("SELECT id FROM users WHERE email = %s", (email,))
        row = cursor.fetchone()
        if row:
            user_id = row[0]
            if name:
                cursor.execute("UPDATE users SET name = %s WHERE id = %s", (name, user_id))
            cursor.close()
            return user_id

        cursor.execute(
            "INSERT INTO users (name, email, role) VALUES (%s, %s, %s)",
            (name, email, role),
        )
        user_id = cursor.lastrowid
        cursor.close()
        return user_id
    finally:
        conn.close()


# =====================================================
# SKILLS
# =====================================================
def add_skills(user_id, skill_names, source="manual", category="Uncategorized", proficiency_level=None):
    """Insert one row per skill name (deduped list) for this user."""
    conn = get_connection()
    try:
        cursor = conn.cursor()
        seen = set()
        for raw in skill_names:
            name = raw.strip()
            key = name.lower()
            if not name or key in seen:
                continue
            seen.add(key)
            cursor.execute(
                """INSERT INTO skills (user_id, skill_name, category, proficiency_level, source)
                   VALUES (%s, %s, %s, %s, %s)""",
                (user_id, name, category, proficiency_level, source),
            )
        cursor.close()
    finally:
        conn.close()


def add_categorized_skills(user_id, categorized_skills, source="manual"):
    """Like add_skills, but each item is {'skill': name, 'category': category}
    (the shape returned by services categorization) instead of a plain name."""
    conn = get_connection()
    try:
        cursor = conn.cursor()
        seen = set()
        for item in categorized_skills:
            name = str(item.get("skill", "")).strip()
            key = name.lower()
            if not name or key in seen:
                continue
            seen.add(key)
            category = item.get("category") or "Uncategorized"
            cursor.execute(
                """INSERT INTO skills (user_id, skill_name, category, proficiency_level, source)
                   VALUES (%s, %s, %s, %s, %s)""",
                (user_id, name, category, None, source),
            )
        cursor.close()
    finally:
        conn.close()


# =====================================================
# ASSESSMENTS + REPORTS
#
# An assessment is written once as 'in_progress' the moment questions are
# generated (start_assessment), updated incrementally as the candidate
# answers (update_assessment_responses), and finalized on submit
# (complete_assessment). This is what makes a page refresh / dropped
# connection recoverable - the row survives even if the browser doesn't.
# =====================================================
def start_assessment(user_id, skills_snapshot, mcqs, ai_generated=True):
    """Create the in-progress row for a freshly generated assessment. Returns its id."""
    conn = get_connection()
    try:
        cursor = conn.cursor()
        cursor.execute(
            """INSERT INTO assessments
               (user_id, skills_snapshot, questions_json, responses_json,
                score, max_score, accuracy, ai_generated, completion_status)
               VALUES (%s, %s, %s, %s, 0, %s, 0.00, %s, 'in_progress')""",
            (
                user_id,
                skills_snapshot,
                json.dumps(mcqs),
                json.dumps({}),
                len(mcqs),
                ai_generated,
            ),
        )
        assessment_id = cursor.lastrowid
        cursor.close()
        return assessment_id
    finally:
        conn.close()


def update_assessment_responses(assessment_id, answers):
    """Persist the candidate's answers so far - called on every answer change."""
    conn = get_connection()
    try:
        cursor = conn.cursor()
        cursor.execute(
            "UPDATE assessments SET responses_json = %s WHERE id = %s",
            (json.dumps(answers), assessment_id),
        )
        cursor.close()
    finally:
        conn.close()


def complete_assessment(assessment_id, score, max_score):
    """Finalize an assessment: set score/accuracy and mark it completed."""
    conn = get_connection()
    try:
        cursor = conn.cursor()
        accuracy = round((score / max_score) * 100, 2) if max_score else 0.0
        cursor.execute(
            """UPDATE assessments
               SET score = %s, max_score = %s, accuracy = %s,
                   completion_status = 'completed', completed_at = %s
               WHERE id = %s""",
            (score, max_score, accuracy, datetime.now(), assessment_id),
        )
        cursor.close()
    finally:
        conn.close()


def get_in_progress_assessment(user_id):
    """Return the most recent in-progress assessment for this user, or None."""
    conn = get_connection()
    try:
        cursor = conn.cursor(dictionary=True)
        cursor.execute(
            """SELECT * FROM assessments
               WHERE user_id = %s AND completion_status = 'in_progress'
               ORDER BY started_at DESC LIMIT 1""",
            (user_id,),
        )
        row = cursor.fetchone()
        cursor.close()
        return row
    finally:
        conn.close()


def save_report(assessment_id, pdf_path, email_status="not_attempted"):
    conn = get_connection()
    try:
        cursor = conn.cursor()
        cursor.execute(
            "INSERT INTO reports (assessment_id, pdf_path, email_status) VALUES (%s, %s, %s)",
            (assessment_id, pdf_path, email_status),
        )
        report_id = cursor.lastrowid
        cursor.close()
        return report_id
    finally:
        conn.close()


def get_all_assessments():
    """Return every assessment joined with the candidate's email, newest first.
    Used by the admin dashboard and by Excel export - list of dicts."""
    conn = get_connection()
    try:
        cursor = conn.cursor(dictionary=True)
        cursor.execute(
            """SELECT a.id, u.email, a.score, a.max_score, a.accuracy,
                      a.ai_generated, a.completed_at
               FROM assessments a
               JOIN users u ON u.id = a.user_id
               ORDER BY a.completed_at DESC"""
        )
        rows = cursor.fetchall()
        cursor.close()
        return rows
    finally:
        conn.close()


def get_assessments_for_email(email):
    conn = get_connection()
    try:
        cursor = conn.cursor(dictionary=True)
        cursor.execute(
            """SELECT a.id, a.score, a.max_score, a.accuracy, a.ai_generated, a.completed_at
               FROM assessments a
               JOIN users u ON u.id = a.user_id
               WHERE u.email = %s
               ORDER BY a.completed_at DESC""",
            (email,),
        )
        rows = cursor.fetchall()
        cursor.close()
        return rows
    finally:
        conn.close()
