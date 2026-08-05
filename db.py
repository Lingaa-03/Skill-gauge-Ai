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

import bcrypt
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


def _ensure_column(cursor, table, column, ddl):
    """Idempotently add a column to an existing table (CREATE TABLE IF NOT
    EXISTS doesn't help once the table already exists from an earlier
    phase - this is the migration path for that)."""
    cursor.execute(
        """SELECT COUNT(*) FROM information_schema.columns
           WHERE table_schema = %s AND table_name = %s AND column_name = %s""",
        (config.MYSQL_DATABASE, table, column),
    )
    if cursor.fetchone()[0] == 0:
        cursor.execute(f"ALTER TABLE {table} ADD COLUMN {ddl}")


def init_db():
    """Create all tables if they don't already exist, run any pending
    column migrations, and seed the admin account. Safe to call every
    startup."""
    conn = get_connection()
    try:
        cursor = conn.cursor()
        for statement in _SCHEMA.strip().split(";"):
            statement = statement.strip()
            if statement:
                cursor.execute(statement)

        # Auth (Phase 4): admin/recruiter accounts log in with a
        # username + hashed password rather than email + skills.
        _ensure_column(cursor, "users", "username", "username VARCHAR(100) NULL UNIQUE")
        _ensure_column(cursor, "users", "password_hash", "password_hash VARCHAR(255) NULL")

        # Matching engine (Phase 5): distinguish internal mobility roles
        # from external candidate-facing job postings.
        _ensure_column(cursor, "jobs", "is_internal", "is_internal BOOLEAN NOT NULL DEFAULT FALSE")

        cursor.close()
    finally:
        conn.close()


# =====================================================
# AUTH (Phase 4)
# =====================================================
def create_account_if_missing(username, password, role="admin"):
    """Seed a username/password account (admin or recruiter) if one with
    this username doesn't already exist. Never overwrites an existing
    password - this only bootstraps the account once."""
    conn = get_connection()
    try:
        cursor = conn.cursor()
        cursor.execute("SELECT id FROM users WHERE username = %s", (username,))
        if cursor.fetchone():
            cursor.close()
            return

        password_hash = bcrypt.hashpw(password.encode("utf-8"), bcrypt.gensalt()).decode("utf-8")
        cursor.execute(
            """INSERT INTO users (name, email, role, username, password_hash)
               VALUES (%s, %s, %s, %s, %s)""",
            (username, f"{username}@internal.skillup", role, username, password_hash),
        )
        cursor.close()
    finally:
        conn.close()


def verify_login(username, password):
    """Check a username/password against the stored hash. Returns the
    user dict (id, role, ...) on success, or None on failure - never
    raises on bad credentials, so a login attempt can't leak timing/
    error-shape information about which part was wrong."""
    conn = get_connection()
    try:
        cursor = conn.cursor(dictionary=True)
        cursor.execute(
            "SELECT id, username, role, password_hash FROM users WHERE username = %s",
            (username,),
        )
        user = cursor.fetchone()
        cursor.close()
    finally:
        conn.close()

    if not user or not user["password_hash"]:
        return None
    if not bcrypt.checkpw(password.encode("utf-8"), user["password_hash"].encode("utf-8")):
        return None
    return user


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


def get_latest_completed_assessment(user_id):
    """Most recent completed assessment for this user, or None."""
    conn = get_connection()
    try:
        cursor = conn.cursor(dictionary=True)
        cursor.execute(
            """SELECT * FROM assessments
               WHERE user_id = %s AND completion_status = 'completed'
               ORDER BY completed_at DESC LIMIT 1""",
            (user_id,),
        )
        row = cursor.fetchone()
        cursor.close()
        return row
    finally:
        conn.close()


def get_skills_for_user(user_id):
    """All skill rows for this user (name + category), most recent first."""
    conn = get_connection()
    try:
        cursor = conn.cursor(dictionary=True)
        cursor.execute(
            "SELECT skill_name, category FROM skills WHERE user_id = %s ORDER BY created_at DESC",
            (user_id,),
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


# =====================================================
# TALENT MARKETPLACE (Phase 5)
#
# jobs/mentors/learning_resources are seeded once (seed_marketplace_data)
# with a small curated catalog. Embeddings are computed lazily by
# services/matching_engine.py and stored back via the set_*_embedding
# functions below, so a fresh row is matchable on the next lookup
# without needing a separate offline embedding job.
# =====================================================
_SEED_JOBS = [
    # (title, description, required_skills, department, is_internal)
    ("Backend Developer", "Build and maintain REST APIs and backend services.", "Python, SQL, REST APIs", "Engineering", False),
    ("Data Scientist", "Analyze data and build predictive models for product decisions.", "Python, Machine Learning, SQL, Statistics", "Data", False),
    ("Frontend Developer", "Build responsive web interfaces for customer-facing products.", "JavaScript, HTML/CSS, React", "Engineering", False),
    ("Java Backend Engineer", "Design and build scalable Java microservices.", "Java, SQL, Spring", "Engineering", False),
    ("Cloud Engineer", "Manage cloud infrastructure and deployment pipelines.", "AWS, Python, DevOps", "Infrastructure", False),
    ("QA Automation Engineer", "Build automated test suites for web applications.", "Java, Python, Selenium", "Quality", False),
    ("ML Engineer (Internal Mobility)", "Move from your current team into applied ML, productionizing models.", "Machine Learning, Python, Deep Learning", "Data", True),
    ("Database Administrator (Internal Mobility)", "Own performance tuning and reliability for core production databases.", "SQL, Database Optimization", "Infrastructure", True),
    ("DevOps Engineer (Internal Mobility)", "Take ownership of CI/CD pipelines and cloud reliability.", "AWS, CI/CD, Docker", "Infrastructure", True),
    ("Full Stack Developer (Internal Mobility)", "Grow into an end-to-end role across an internal product line.", "JavaScript, Python, SQL, HTML/CSS", "Engineering", True),
    ("Data Analyst (Internal Mobility)", "Support internal teams with dashboards and ad-hoc analysis.", "SQL, Python, Data Visualization", "Data", True),
    ("Senior Software Engineer (Internal Mobility)", "Step up into a senior role owning system design for a core service.", "Python, System Design, SQL, Java", "Engineering", True),
]

_SEED_MENTORS = [
    # (name, expertise, bio)
    ("Anita Rao", "Machine Learning, Python, Statistics", "Senior Data Scientist, 8 years building ML products."),
    ("Rahul Mehta", "AWS, DevOps, Python", "Cloud Architect specializing in scalable infrastructure."),
    ("Sara Khan", "Java, SQL, System Design", "Backend Lead with deep experience in distributed systems."),
    ("Vikram Nair", "JavaScript, React, HTML/CSS", "Frontend Lead focused on accessible, performant UI."),
    ("Priya Menon", "SQL, Database Optimization", "Database Specialist for high-throughput production systems."),
    ("Arjun Iyer", "Python, JavaScript, SQL", "Full Stack mentor supporting engineers moving across the stack."),
]

_SEED_LEARNING_RESOURCES = [
    # (title, skill_tag, url, description)
    # skill_tag must use the SAME vocabulary as jobs.required_skills, so
    # that a skill gap identified from a job match can actually resolve to
    # a resource here. Keep these aligned when adding either side.
    ("Advanced Python for Backend Development", "Python", "https://learning.hexacorp.internal/course/advanced-python", "Deep dive into backend patterns in Python."),
    ("Machine Learning Foundations", "Machine Learning", "https://learning.hexacorp.internal/course/ml-foundations", "Core ML concepts and applied modeling."),
    ("SQL Performance Tuning", "SQL", "https://learning.hexacorp.internal/course/sql-tuning", "Indexing, query plans, and optimization."),
    ("AWS Cloud Practitioner Path", "AWS", "https://learning.hexacorp.internal/course/aws-practitioner", "Foundational AWS services and architecture."),
    ("Modern JavaScript & React", "JavaScript", "https://learning.hexacorp.internal/course/modern-js-react", "Component-driven frontend development."),
    ("Java Spring Boot Masterclass", "Java", "https://learning.hexacorp.internal/course/spring-boot", "Building production Java services with Spring."),
    ("HTML/CSS Responsive Design", "HTML/CSS", "https://learning.hexacorp.internal/course/responsive-design", "Layout, accessibility, and responsive patterns."),
    ("Data Structures & System Design", "System Design", "https://learning.hexacorp.internal/course/system-design", "Designing scalable, maintainable systems."),
    ("Statistics for Data Science", "Statistics", "https://learning.hexacorp.internal/course/stats-for-ds", "Hypothesis testing, regression, and experimental design."),
    ("DevOps & CI/CD Essentials", "DevOps", "https://learning.hexacorp.internal/course/devops-cicd", "Pipelines, containers, and deployment automation."),
    ("Deep Learning with Neural Networks", "Deep Learning", "https://learning.hexacorp.internal/course/deep-learning", "Neural network architectures and training."),
    ("Data Visualization & Dashboards", "Data Visualization", "https://learning.hexacorp.internal/course/data-viz", "Building clear, effective dashboards and reports."),
    ("Designing REST APIs", "REST APIs", "https://learning.hexacorp.internal/course/rest-apis", "API design, versioning, and documentation."),
    ("Database Optimization Deep Dive", "Database Optimization", "https://learning.hexacorp.internal/course/db-optimization", "Advanced tuning for high-throughput databases."),
    ("Test Automation with Selenium", "Selenium", "https://learning.hexacorp.internal/course/selenium", "Automated browser testing at scale."),
    ("Docker & Containers in Practice", "Docker", "https://learning.hexacorp.internal/course/docker", "Containerizing and shipping applications."),
    ("CI/CD Pipeline Engineering", "CI/CD", "https://learning.hexacorp.internal/course/cicd-pipelines", "Designing reliable build and release pipelines."),
    ("Spring Framework Essentials", "Spring", "https://learning.hexacorp.internal/course/spring-essentials", "Dependency injection, data access, and web layers."),
]


def seed_marketplace_data():
    """Insert the curated jobs/mentors/learning-resources catalog if each
    table is still empty. Safe to call every startup - never duplicates."""
    conn = get_connection()
    try:
        cursor = conn.cursor()

        cursor.execute("SELECT COUNT(*) FROM jobs")
        if cursor.fetchone()[0] == 0:
            cursor.executemany(
                """INSERT INTO jobs (title, description, required_skills, department, is_internal)
                   VALUES (%s, %s, %s, %s, %s)""",
                _SEED_JOBS,
            )

        cursor.execute("SELECT COUNT(*) FROM mentors")
        if cursor.fetchone()[0] == 0:
            cursor.executemany(
                "INSERT INTO mentors (name, expertise, bio) VALUES (%s, %s, %s)",
                _SEED_MENTORS,
            )

        cursor.execute("SELECT COUNT(*) FROM learning_resources")
        if cursor.fetchone()[0] == 0:
            cursor.executemany(
                "INSERT INTO learning_resources (title, skill_tag, url, description) VALUES (%s, %s, %s, %s)",
                _SEED_LEARNING_RESOURCES,
            )

        cursor.close()
    finally:
        conn.close()


def get_jobs():
    conn = get_connection()
    try:
        cursor = conn.cursor(dictionary=True)
        cursor.execute("SELECT * FROM jobs")
        rows = cursor.fetchall()
        cursor.close()
        return rows
    finally:
        conn.close()


def get_mentors():
    conn = get_connection()
    try:
        cursor = conn.cursor(dictionary=True)
        cursor.execute("SELECT * FROM mentors")
        rows = cursor.fetchall()
        cursor.close()
        return rows
    finally:
        conn.close()


def get_learning_resources():
    conn = get_connection()
    try:
        cursor = conn.cursor(dictionary=True)
        cursor.execute("SELECT * FROM learning_resources")
        rows = cursor.fetchall()
        cursor.close()
        return rows
    finally:
        conn.close()


def set_job_embedding(job_id, embedding_json):
    _set_embedding("jobs", job_id, embedding_json)


def set_mentor_embedding(mentor_id, embedding_json):
    _set_embedding("mentors", mentor_id, embedding_json)


def set_learning_resource_embedding(resource_id, embedding_json):
    _set_embedding("learning_resources", resource_id, embedding_json)


def _set_embedding(table, row_id, embedding_json):
    conn = get_connection()
    try:
        cursor = conn.cursor()
        cursor.execute(f"UPDATE {table} SET embedding_json = %s WHERE id = %s", (embedding_json, row_id))
        cursor.close()
    finally:
        conn.close()
