import streamlit as st
import os, json, re, smtplib, random, time
from xml.sax.saxutils import escape
from datetime import datetime
from email.message import EmailMessage

import pandas as pd
import matplotlib.pyplot as plt

from PyPDF2 import PdfReader
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.units import inch

from openpyxl import Workbook

import config
import db
from services import ai_client, matching_engine

# =====================================================
# PAGE CONFIG & STYLING
# =====================================================
st.set_page_config(
    page_title="SkillGauge AI – Interview Assessment",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# =====================================================
# CUSTOM CSS (ONLY TEXT COLOR FIXED)
# =====================================================
st.markdown("""
<style>
* {
    font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
}

html, body, [data-testid="stAppViewContainer"] {
    background: linear-gradient(135deg, #667eea 0%, #764ba2 50%, #f093fb 100%) !important;
    min-height: 100vh;
}

/* Headings remain white */
h1, h2, h3 {
    color: #ffffff;
    font-weight: 700;
}

/* -----------------------------
   BLACK TEXT VISIBILITY FIX
------------------------------ */

/* Input text */
.stTextInput input,
.stTextArea textarea,
.stSelectbox select {
    color: #000000 !important;
    background-color: rgba(255,255,255,0.85) !important;
}

/* Placeholder text */
.stTextInput input::placeholder,
.stTextArea textarea::placeholder {
    color: #000000 !important;
    opacity: 0.7;
}

/* Labels */
label,
.stRadio label,
.stSelectbox label {
    color: #000000 !important;
    font-weight: 600;
}

/* Markdown text */
[data-testid="stMarkdownContainer"] {
    color: #000000 !important;
}

/* Download button text */
.stDownloadButton button {
    color: #000000 !important;
}

/* Metric labels */
.stMetric label {
    color: #000000 !important;
}

/* Buttons */
.stButton > button {
    background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
    color: #000000 !important;
    font-weight: 600;
}

/* Containers */
.stContainer {
    background: rgba(255, 255, 255, 0.75);
    backdrop-filter: blur(10px);
    padding: 24px;
    border-radius: 12px;
}

/* Alerts text */
.stSuccess, .stError, .stWarning {
    color: #000000 !important;
}
</style>
""", unsafe_allow_html=True)

# =====================================================
# DATABASE (MySQL is the source of truth; Excel is export-only - see
# export_excel_from_db() below)
# =====================================================
@st.cache_resource
def _init_db_once():
    db.init_db()
    # Seed the admin account from .env on first run only - this never
    # overwrites a password that's already been set.
    db.create_account_if_missing(config.ADMIN_USERNAME, config.ADMIN_PASSWORD, role="admin")
    db.create_account_if_missing(
        config.RECRUITER_USERNAME, config.RECRUITER_PASSWORD, role="recruiter"
    )
    # Seed the jobs/mentors/learning catalog the matching engine ranks against.
    db.seed_marketplace_data()
    return True

_init_db_once()

# =====================================================
# PATHS (from config.py, which reads .env)
# =====================================================
RESULT_DIR = config.RESULT_DIR
PDF_DIR = config.PDF_DIR
EXCEL_PATH = config.EXCEL_PATH

# =====================================================
# EMAIL CONFIG (from config.py, which reads .env)
# =====================================================
SENDER_EMAIL = config.SENDER_EMAIL
APP_PASSWORD = config.EMAIL_APP_PASSWORD

# =====================================================
# FALLBACK QUESTION BANK (50 Questions - 25 Random Selected if AI Fails)
# =====================================================
QUESTION_BANK = [
    # Python (7 questions)
    {
        "q": "What is the output of: print(type([]))",
        "options": {"A": "<class 'list'>", "B": "<class 'tuple'>", "C": "<class 'dict'>", "D": "<class 'array'>"},
        "answer": "A",
        "domain": "Python"
    },
    {
        "q": "Which keyword is used to create a function in Python?",
        "options": {"A": "function", "B": "def", "C": "define", "D": "func"},
        "answer": "B",
        "domain": "Python"
    },
    {
        "q": "What does len() function return for a string 'hello'?",
        "options": {"A": "4", "B": "5", "C": "6", "D": "None"},
        "answer": "B",
        "domain": "Python"
    },
    {
        "q": "How do you access the first element of a list in Python?",
        "options": {"A": "list[1]", "B": "list[0]", "C": "list.first()", "D": "list[:]"},
        "answer": "B",
        "domain": "Python"
    },
    {
        "q": "What is the correct syntax for creating a dictionary in Python?",
        "options": {"A": "{key: value}", "B": "[key: value]", "C": "(key: value)", "D": "key: value"},
        "answer": "A",
        "domain": "Python"
    },
    {
        "q": "Which method is used to add an element to a list?",
        "options": {"A": "add()", "B": "insert()", "C": "append()", "D": "push()"},
        "answer": "C",
        "domain": "Python"
    },
    {
        "q": "What does 'if __name__ == \"__main__\":' do in Python?",
        "options": {"A": "Creates a main function", "B": "Runs code only when script is executed directly", "C": "Creates a namespace", "D": "Imports the main module"},
        "answer": "B",
        "domain": "Python"
    },

    # SQL (7 questions)
    {
        "q": "Which SQL statement is used to retrieve data?",
        "options": {"A": "GET", "B": "RETRIEVE", "C": "SELECT", "D": "EXTRACT"},
        "answer": "C",
        "domain": "SQL"
    },
    {
        "q": "What does JOIN do in SQL?",
        "options": {"A": "Combines rows from two or more tables", "B": "Deletes duplicate rows", "C": "Sorts data", "D": "Groups data"},
        "answer": "A",
        "domain": "SQL"
    },
    {
        "q": "Which SQL keyword is used to filter results?",
        "options": {"A": "FILTER", "B": "WHERE", "C": "SEARCH", "D": "FIND"},
        "answer": "B",
        "domain": "SQL"
    },
    {
        "q": "What is the correct syntax for an INNER JOIN?",
        "options": {"A": "SELECT * FROM table1 INNER JOIN table2 ON condition", "B": "SELECT * FROM table1 JOIN table2 WHERE condition", "C": "SELECT * FROM table1 table2", "D": "SELECT * FROM table1 & table2"},
        "answer": "A",
        "domain": "SQL"
    },
    {
        "q": "Which aggregate function returns the average?",
        "options": {"A": "AVG()", "B": "AVERAGE()", "C": "MEAN()", "D": "SUM()"},
        "answer": "A",
        "domain": "SQL"
    },
    {
        "q": "What does GROUP BY do?",
        "options": {"A": "Sorts results ascending", "B": "Combines rows with same values", "C": "Limits results", "D": "Joins tables"},
        "answer": "B",
        "domain": "SQL"
    },
    {
        "q": "Which statement creates a new table in SQL?",
        "options": {"A": "NEW TABLE", "B": "CREATE TABLE", "C": "ADD TABLE", "D": "MAKE TABLE"},
        "answer": "B",
        "domain": "SQL"
    },

    # Machine Learning (7 questions)
    {
        "q": "What is the purpose of train-test split?",
        "options": {"A": "Reduce model size", "B": "Evaluate model on unseen data", "C": "Speed up training", "D": "Increase accuracy"},
        "answer": "B",
        "domain": "Machine Learning"
    },
    {
        "q": "Which metric is used for classification problems?",
        "options": {"A": "Mean Squared Error", "B": "R-squared", "C": "Accuracy", "D": "RMSE"},
        "answer": "C",
        "domain": "Machine Learning"
    },
    {
        "q": "What is overfitting in machine learning?",
        "options": {"A": "Model learns training data too well", "B": "Model fails to learn patterns", "C": "Model is too simple", "D": "Model uses too few features"},
        "answer": "A",
        "domain": "Machine Learning"
    },
    {
        "q": "Which algorithm is used for clustering?",
        "options": {"A": "K-Means", "B": "Linear Regression", "C": "SVM", "D": "Decision Trees"},
        "answer": "A",
        "domain": "Machine Learning"
    },
    {
        "q": "What does feature scaling do?",
        "options": {"A": "Removes duplicate features", "B": "Normalizes feature values to similar range", "C": "Increases model accuracy", "D": "Reduces computation time"},
        "answer": "B",
        "domain": "Machine Learning"
    },
    {
        "q": "Which loss function is used for regression?",
        "options": {"A": "Cross Entropy", "B": "Mean Squared Error", "C": "Hinge Loss", "D": "KL Divergence"},
        "answer": "B",
        "domain": "Machine Learning"
    },
    {
        "q": "What is cross-validation?",
        "options": {"A": "Training on full dataset", "B": "Splitting data into folds for evaluation", "C": "Combining multiple models", "D": "Removing outliers"},
        "answer": "B",
        "domain": "Machine Learning"
    },

    # HTML/CSS (7 questions)
    {
        "q": "Which tag is used for the largest heading in HTML?",
        "options": {"A": "<h6>", "B": "<head>", "C": "<h1>", "D": "<heading>"},
        "answer": "C",
        "domain": "HTML/CSS"
    },
    {
        "q": "Which CSS property is used to change text color?",
        "options": {"A": "font-color", "B": "text-color", "C": "color", "D": "background-color"},
        "answer": "C",
        "domain": "HTML/CSS"
    },
    {
        "q": "What does CSS stand for?",
        "options": {"A": "Computer Style Sheet", "B": "Cascading Style Sheet", "C": "Creative Style Sheet", "D": "Coded Style Sheet"},
        "answer": "B",
        "domain": "HTML/CSS"
    },
    {
        "q": "Which HTML tag is used to create a line break?",
        "options": {"A": "<break>", "B": "<line>", "C": "<br>", "D": "<newline>"},
        "answer": "C",
        "domain": "HTML/CSS"
    },
    {
        "q": "Which CSS property controls the space between lines?",
        "options": {"A": "letter-spacing", "B": "line-height", "C": "text-spacing", "D": "word-spacing"},
        "answer": "B",
        "domain": "HTML/CSS"
    },
    {
        "q": "What is the default display property for a <div> element?",
        "options": {"A": "inline", "B": "block", "C": "flex", "D": "grid"},
        "answer": "B",
        "domain": "HTML/CSS"
    },
    {
        "q": "Which attribute is used to specify alternative text for images?",
        "options": {"A": "title", "B": "alt", "C": "description", "D": "text"},
        "answer": "B",
        "domain": "HTML/CSS"
    },

    # Java (7 questions)
    {
        "q": "What does JVM stand for?",
        "options": {"A": "Java Virtual Memory", "B": "Java Virtual Machine", "C": "Java Verified Module", "D": "Java Valid Module"},
        "answer": "B",
        "domain": "Java"
    },
    {
        "q": "Which keyword is used to define a class in Java?",
        "options": {"A": "Class", "B": "class", "C": "CLASS", "D": "Clazz"},
        "answer": "B",
        "domain": "Java"
    },
    {
        "q": "What is an interface in Java?",
        "options": {"A": "A class with multiple implementations", "B": "A contract for classes to follow", "C": "A memory location", "D": "A type of variable"},
        "answer": "B",
        "domain": "Java"
    },
    {
        "q": "Which access modifier makes a method accessible from any class?",
        "options": {"A": "private", "B": "protected", "C": "public", "D": "default"},
        "answer": "C",
        "domain": "Java"
    },
    {
        "q": "What does 'this' keyword refer to in Java?",
        "options": {"A": "Current object", "B": "Parent class", "C": "Static variable", "D": "Method parameter"},
        "answer": "A",
        "domain": "Java"
    },
    {
        "q": "Which method is used to create a string from an object in Java?",
        "options": {"A": "valueOf()", "B": "toString()", "C": "stringify()", "D": "convert()"},
        "answer": "B",
        "domain": "Java"
    },
    {
        "q": "What is inheritance in Java?",
        "options": {"A": "Creating multiple objects", "B": "Acquiring properties from parent class", "C": "Creating interfaces", "D": "Memory allocation"},
        "answer": "B",
        "domain": "Java"
    },

    # JavaScript (5 questions)
    {
        "q": "How do you declare a variable in modern JavaScript?",
        "options": {"A": "var x = 5", "B": "let x = 5", "C": "const x = 5", "D": "Both B and C are modern"},
        "answer": "D",
        "domain": "JavaScript"
    },
    {
        "q": "What does the 'async' keyword do in JavaScript?",
        "options": {"A": "Runs code synchronously", "B": "Declares asynchronous function", "C": "Stops code execution", "D": "Creates a loop"},
        "answer": "B",
        "domain": "JavaScript"
    },
    {
        "q": "Which method is used to find an element by ID?",
        "options": {"A": "getElement()", "B": "findElement()", "C": "getElementById()", "D": "querySelector()"},
        "answer": "C",
        "domain": "JavaScript"
    },
    {
        "q": "What does JSON stand for?",
        "options": {"A": "JavaScript Object Notation", "B": "Java Serialized Object", "C": "JavaScript Original Notation", "D": "Java Object Network"},
        "answer": "A",
        "domain": "JavaScript"
    },
    {
        "q": "How do you add a comment in JavaScript?",
        "options": {"A": "# comment", "B": "// comment or /* */", "C": "-- comment", "D": "<!-- comment -->"},
        "answer": "B",
        "domain": "JavaScript"
    },

    # AWS (4 questions)
    {
        "q": "What does EC2 stand for in AWS?",
        "options": {"A": "Elastic Cloud Compute", "B": "Elastic Compute Cloud", "C": "Elastic Cloud Console", "D": "Elastic Cluster Compute"},
        "answer": "B",
        "domain": "AWS"
    },
    {
        "q": "Which AWS service is used for object storage?",
        "options": {"A": "EC2", "B": "RDS", "C": "S3", "D": "Lambda"},
        "answer": "C",
        "domain": "AWS"
    },
    {
        "q": "What is AWS Lambda?",
        "options": {"A": "Relational database", "B": "Serverless compute service", "C": "Storage service", "D": "Network service"},
        "answer": "B",
        "domain": "AWS"
    },
    {
        "q": "Which AWS service manages relational databases?",
        "options": {"A": "DynamoDB", "B": "RDS", "C": "S3", "D": "EC2"},
        "answer": "B",
        "domain": "AWS"
    }
]

# =====================================================
# SESSION STATE
# =====================================================
if "step" not in st.session_state:
    st.session_state.update({
        "step": 0,
        "mcqs": [],
        "answers": {},
        "score": 0,
        "email": "",
        "evaluated": False,
        "admin_logged": False,
        "admin_role": None,
        "ai_used": False,  # Track if AI or fallback was used
        "user_id": None,
        "skills_snapshot": "",
        "saved_to_db": False,
        "assessment_id": None,
        "assessment_start_time": None,
        "recommendations": None,
        "assessment_time_limit": 1500,  # seconds; recomputed as len(mcqs) * 60
        "current_question": 0,
    })

# =====================================================
# SAFE JSON EXTRACTION (FLASH-SAFE)
# =====================================================
def extract_json_flexible(text):
    text = re.sub(r"```json|```", "", text).strip()

    # Try direct JSON
    try:
        data = json.loads(text)
        if isinstance(data, list):
            return data
        if isinstance(data, dict):
            for k in ("mcqs", "questions"):
                if k in data and isinstance(data[k], list):
                    return data[k]
    except:
        pass

    # Extract JSON array from text
    match = re.search(r"\[[\s\S]*\]", text)
    if match:
        try:
            return json.loads(match.group())
        except:
            pass

    return None


def extract_mcqs(parsed):
    if isinstance(parsed, dict):
        for k in ("mcq", "mcqs", "questions"):
            if k in parsed and isinstance(parsed[k], list):
                return parsed[k]
    if isinstance(parsed, list):
        return parsed
    return None


# =====================================================
# SKILL EXTRACTION
# =====================================================
def extract_skills_from_resume(file):
    reader = PdfReader(file)
    text = " ".join(p.extract_text() or "" for p in reader.pages)

    prompt = f"""
Extract ONLY technical skills, programming languages, frameworks,
tools and domains (comma separated).
Resume:
{text[:6000]}
"""
    try:
        return ai_client.chat_complete(prompt).strip()
    except Exception as e:
        print(f"[SKILL EXTRACTION ERROR]: {e}")
        st.warning("⚠️ AI unavailable. Please enter skills manually.")
        return ""


def categorize_skills(skills_text):
    """
    Best-effort categorization for DB storage. Never blocks the flow -
    falls back to a deduped, uncategorized list if the AI call fails or
    returns something unparsable.
    """
    skill_list = [s.strip() for s in skills_text.split(",") if s.strip()]
    if not skill_list:
        return []

    prompt = f"""
Categorize each of these technical skills into exactly ONE category from:
Programming Language, Framework, Database, Cloud, Tool, Domain/Concept.

Skills: {", ".join(skill_list)}

Return ONLY a valid JSON array, no commentary, in this exact format:
[{{"skill": "Python", "category": "Programming Language"}}]
"""
    try:
        text = ai_client.chat_complete(prompt, max_tokens=3000)
        parsed = extract_json_flexible(text)
        if parsed and isinstance(parsed, list):
            seen, cleaned = set(), []
            for item in parsed:
                if not isinstance(item, dict):
                    continue
                name = str(item.get("skill", "")).strip()
                category = str(item.get("category", "")).strip() or "Uncategorized"
                key = name.lower()
                if name and key not in seen:
                    seen.add(key)
                    cleaned.append({"skill": name, "category": category})
            if cleaned:
                return cleaned
    except Exception as e:
        print(f"[SKILL CATEGORIZATION ERROR]: {e}")

    # Fallback: dedup only, no category
    seen, cleaned = set(), []
    for name in skill_list:
        key = name.lower()
        if key not in seen:
            seen.add(key)
            cleaned.append({"skill": name, "category": "Uncategorized"})
    return cleaned


def validate_mcqs(mcqs):
    """
    Guardrail pass over AI-generated questions before they're ever served
    to a candidate: exactly 4 non-empty options, a valid answer key, and
    no duplicate questions. Anything malformed is dropped rather than shown.
    """
    if not isinstance(mcqs, list):
        return []

    valid, seen_questions = [], set()
    for item in mcqs:
        if not isinstance(item, dict):
            continue
        q_text = str(item.get("q", "")).strip()
        options = item.get("options")
        answer = item.get("answer")

        if not q_text or not isinstance(options, dict):
            continue
        if set(options.keys()) != {"A", "B", "C", "D"}:
            continue
        if any(not str(v).strip() for v in options.values()):
            continue
        if answer not in options:
            continue

        key = q_text.lower()
        if key in seen_questions:
            continue
        seen_questions.add(key)
        valid.append(item)

    return valid


def get_fallback_mcqs(skills_text, count=25):
    """
    Skill-aware fallback: prefer curated questions whose domain matches
    the candidate's stated skills, then top up from the rest of the bank
    so a full assessment is always produced.
    """
    skill_tokens = [s.strip().lower() for s in skills_text.split(",") if s.strip()]

    def matches(question):
        domain = question["domain"].lower()
        return any(tok in domain or domain in tok for tok in skill_tokens)

    matched = [q for q in QUESTION_BANK if matches(q)]
    random.shuffle(matched)

    if len(matched) >= count:
        return matched[:count]

    remaining_pool = [q for q in QUESTION_BANK if q not in matched]
    random.shuffle(remaining_pool)
    return (matched + remaining_pool)[:count]


def generate_interview_mcqs(skills, count=25):
    """
    AI-first MCQ generation, validated by validate_mcqs() before being
    accepted, with a skill-aware fallback bank if the AI is unavailable
    or produces too few well-formed questions.
    """
    prompt = f"""
You are an experienced technical interviewer designing a screening assessment.

Generate EXACTLY {count} multiple-choice interview questions strictly
scoped to these candidate skills: {skills}

Rules:
- Fixed difficulty: intermediate (not beginner trivia, not expert-only edge cases).
- Every question has exactly 4 options (A, B, C, D) with exactly one correct answer.
- Do not repeat questions or ask near-duplicates of each other.
- Do not ask about topics outside the given skills.
- Avoid ambiguous wording - there must be exactly one unambiguous correct answer.

Return ONLY a valid JSON ARRAY, no commentary, in this exact format:
[
  {{
    "q": "Question",
    "options": {{
      "A": "...",
      "B": "...",
      "C": "...",
      "D": "..."
    }},
    "answer": "A"
  }}
]
"""

    for attempt in range(2):
        try:
            text = ai_client.chat_complete(prompt, retries=1, max_tokens=12000)
            mcqs = validate_mcqs(extract_json_flexible(text))

            if len(mcqs) >= count:
                st.session_state.ai_used = True
                return mcqs[:count]

        except Exception as e:
            print(f"[AI ERROR] Attempt {attempt+1}:", e)

    # FALLBACK - skill-aware, not fully random
    st.warning("⚠️ AI service temporarily unavailable. Using curated questions matched to your skills.")
    st.session_state.ai_used = False
    return get_fallback_mcqs(skills, count=count)


# =====================================================
# SEND EMAIL WITH PDF REPORT
# =====================================================
def send_email_report(recipient_email, pdf_path, student_name="Student"):
    """Send assessment report via email"""
    try:
        msg = EmailMessage()
        msg['Subject'] = "Your SkillGauge AI Assessment Report"
        msg['From'] = SENDER_EMAIL
        msg['To'] = recipient_email
        
        body = f"""
Hi {student_name},

Thank you for completing your SkillGauge AI Assessment!

Your personalized interview assessment report is attached. This report includes:
- Questions asked
- Your answers
- Correct answers
- Overall score and performance analysis

We hope this helps you prepare for your interview. If you have any questions, please feel free to reach out.

Best regards,
SkillGauge AI Team
"""
        
        msg.set_content(body)
        
        # Attach PDF
        with open(pdf_path, 'rb') as attachment:
            msg.add_attachment(
                attachment.read(),
                maintype='application',
                subtype='pdf',
                filename=os.path.basename(pdf_path)
            )
        
        # Send via Gmail
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
            smtp.login(SENDER_EMAIL, APP_PASSWORD)
            smtp.send_message(msg)
        
        return True
    except Exception as e:
        print(f"[v0] Email send failed: {str(e)}")
        return False

# =====================================================
# EXCEL EXPORT (backup/export only - MySQL is the source of truth)
# =====================================================
def export_excel_from_db():
    """Regenerate the Excel file from the latest DB records and return its path.
    Called on demand (admin dashboard), never as the write path for a result."""
    rows = db.get_all_assessments_detailed()

    wb = Workbook()
    ws = wb.active
    ws.title = "Assessments"
    ws.append([
        "Email", "Name", "Assessment ID", "Skills", "Score", "Max Marks",
        "Accuracy %", "Question Source", "Status", "Started", "Completed",
        "Report Path", "Email Status",
    ])
    for row in rows:
        started, completed = row["started_at"], row["completed_at"]
        ws.append([
            row["email"],
            row["name"] or "",
            row["assessment_id"],
            row["skills_snapshot"] or "",
            row["score"],
            row["max_score"],
            float(row["accuracy"]) if row["accuracy"] is not None else 0.0,
            "AI-Generated" if row["ai_generated"] else "Curated Fallback",
            row["completion_status"],
            started.strftime("%Y-%m-%d %H:%M:%S") if started else "",
            completed.strftime("%Y-%m-%d %H:%M:%S") if completed else "",
            row["pdf_path"] or "",
            row["email_status"] or "",
        ])

    # Second sheet: the skill-demand view of the talent pool
    ws_skills = wb.create_sheet("Skill Demand")
    ws_skills.append(["Skill", "Category", "Candidate Count"])
    for skill in db.get_skill_demand(limit=100):
        ws_skills.append([skill["skill_name"], skill["category"], skill["candidate_count"]])

    wb.save(EXCEL_PATH)
    return EXCEL_PATH

# =====================================================
# PDF GENERATION
# =====================================================
def generate_pdf(email, mcqs, answers, score):
    path = os.path.join(PDF_DIR, f"result_{int(datetime.now().timestamp())}.pdf")
    styles = getSampleStyleSheet()

    styles.add(ParagraphStyle(
        name="TitleX",
        alignment=TA_CENTER,
        fontSize=20,
        textColor="#1e3a5f"
    ))

    doc = SimpleDocTemplate(path)
    story = []

    story.append(Paragraph("SkillGauge AI – Interview Assessment Report", styles["TitleX"]))
    story.append(Spacer(1, 12))
    story.append(Paragraph(f"<b>Email:</b> {email}", styles["Normal"]))
    story.append(Paragraph(f"<b>Report Date:</b> {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}", styles["Normal"]))
    story.append(Spacer(1, 12))

    for i, q in enumerate(mcqs, 1):
        story.append(Spacer(1, 10))
        safe_question = escape(q['q'])
        story.append(Paragraph(f"<b>{i}. {safe_question}</b>", styles["Normal"]))

        
        user_ans = answers.get(i, "Not answered")
        correct_ans = q["answer"]
        is_correct = user_ans == correct_ans
        
        for k, v in q["options"].items():
            marker = "✓" if (k == correct_ans and is_correct) else "✗" if (k == user_ans and not is_correct) else ""
            safe_option = escape(f"{k}) {v} {marker}")
            story.append(Paragraph(safe_option, styles["Normal"]))
        
        story.append(Spacer(1, 6))

    total_questions = len(mcqs)
    story.append(Spacer(1, 12))
    story.append(Paragraph(f"<b>FINAL SCORE: {score}/{total_questions}</b>", styles["Normal"]))
    percentage = (score / total_questions) * 100 if total_questions else 0
    story.append(Paragraph(f"<b>Percentage: {percentage:.1f}%</b>", styles["Normal"]))

    doc.build(story)
    return path

# =====================================================
# HEADER WITH ADMIN BUTTON
# =====================================================
col1, col2 = st.columns([10, 1])
with col1:
    st.markdown("### 🎯 SkillGauge AI")
with col2:
    if st.button("🔐", help="Admin Login"):
        st.session_state.step = "admin_login"
        st.rerun()

st.divider()

# =====================================================
# ADMIN LOGIN
# =====================================================
if st.session_state.step == "admin_login":
    st.markdown("## 🔐 Admin Login")
    
    col1, col2, col3 = st.columns([1, 1, 1])
    
    with col2:
        with st.container():
            st.markdown("---")
            username = st.text_input("Username", placeholder="Enter username")
            password = st.text_input("Password", type="password", placeholder="Enter password")
            
            col_login, col_back = st.columns(2)
            with col_login:
                if st.button("✓ Login", use_container_width=True):
                    user = db.verify_login(username.strip(), password)
                    if user and user["role"] in ("admin", "recruiter"):
                        st.session_state.admin_logged = True
                        st.session_state.admin_role = user["role"]
                        st.session_state.step = "admin_dashboard"
                        st.rerun()
                    else:
                        st.error("❌ Invalid credentials. Please try again.")
            
            with col_back:
                if st.button("← Back", use_container_width=True):
                    st.session_state.step = 0
                    st.rerun()

# =====================================================
# ADMIN DASHBOARD
# =====================================================
elif st.session_state.step == "admin_dashboard":
    role = st.session_state.admin_role or "recruiter"
    is_admin = role == "admin"

    st.markdown(f"## 📊 {'Admin' if is_admin else 'Recruiter'} Dashboard")
    st.caption(f"Signed in as **{role}**" + ("" if is_admin else " — read-only access"))
    st.markdown("---")

    rows = db.get_all_assessments()

    if not rows:
        st.warning("⚠️ No assessment data yet.")
    else:
        stats = db.get_dashboard_stats()

        df = pd.DataFrame(rows)
        df["Date"] = df["completed_at"].apply(lambda d: d.strftime("%Y-%m-%d") if d else "")
        df["Time"] = df["completed_at"].apply(lambda d: d.strftime("%H:%M:%S") if d else "")
        df["Accuracy %"] = df["accuracy"].apply(lambda a: float(a) if a is not None else 0.0)
        df["Source"] = df["ai_generated"].apply(lambda v: "AI" if v else "Fallback")
        df = df.rename(columns={"email": "Email", "score": "Score", "max_score": "Max Marks"})
        df = df[["Email", "Score", "Max Marks", "Accuracy %", "Source", "Date", "Time"]]

        # --- Headline metrics (accuracy-based, so a variable question
        # count doesn't break the numbers the way a hardcoded /25 did) ---
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("Total Assessments", stats["total_assessments"])
        with col2:
            st.metric("Unique Candidates", stats["unique_candidates"])
        with col3:
            avg = stats["avg_accuracy"]
            st.metric("Avg Accuracy", f"{avg:.1f}%" if avg is not None else "—")
        with col4:
            total = stats["total_assessments"]
            st.metric(
                "Pass Rate (≥60%)",
                f"{(stats['passed_count'] / total * 100):.0f}%" if total else "—",
            )

        col5, col6, col7, col8 = st.columns(4)
        with col5:
            best = stats["best_accuracy"]
            st.metric("Best", f"{best:.1f}%" if best is not None else "—")
        with col6:
            worst = stats["worst_accuracy"]
            st.metric("Lowest", f"{worst:.1f}%" if worst is not None else "—")
        with col7:
            st.metric("AI-Generated", stats["ai_generated_count"])
        with col8:
            st.metric("Fallback Used", stats["fallback_count"])

        if stats["abandoned"]:
            st.info(f"↻ {stats['abandoned']} assessment(s) currently in progress / abandoned.")

        st.markdown("---")

        tab_records, tab_analytics, tab_candidate = st.tabs([
            "📋 Records", "📈 Analytics", "🔍 Candidate Detail",
        ])

        with tab_records:
            st.dataframe(df, use_container_width=True)

            if is_admin:
                # Regenerate the Excel export fresh from the DB every time,
                # so it never drifts from the source of truth.
                excel_path = export_excel_from_db()
                with open(excel_path, "rb") as f:
                    st.download_button(
                        label="📥 Download Full Excel Report",
                        data=f.read(),
                        file_name="skillgauge_assessments.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                    )
            else:
                st.caption("🔒 Excel export is restricted to admin accounts.")

        with tab_analytics:
            col_a, col_b = st.columns(2)

            with col_a:
                st.subheader("Score Distribution")
                buckets = db.get_score_distribution()
                if buckets:
                    order = ["0-19%", "20-39%", "40-59%", "60-79%", "80-100%"]
                    counts = {b["bucket"]: b["count"] for b in buckets}
                    labels = [b for b in order if b in counts]
                    values = [counts[b] for b in labels]

                    fig, ax = plt.subplots(figsize=(4.5, 3))
                    fig.patch.set_alpha(0)
                    ax.patch.set_alpha(0)
                    ax.bar(labels, values, color="#6366f1")
                    ax.set_ylabel("Candidates", color="white")
                    ax.tick_params(colors="white", labelsize=8)
                    for spine in ax.spines.values():
                        spine.set_color("white")
                    st.pyplot(fig, use_container_width=False)
                else:
                    st.info("Not enough data yet.")

            with col_b:
                st.subheader("Question Source Split")
                ai_count = stats["ai_generated_count"]
                fb_count = stats["fallback_count"]
                if ai_count or fb_count:
                    fig2, ax2 = plt.subplots(figsize=(4, 3))
                    fig2.patch.set_alpha(0)
                    ax2.patch.set_alpha(0)
                    ax2.pie(
                        [ai_count, fb_count],
                        labels=["AI-Generated", "Fallback"],
                        autopct="%1.0f%%",
                        colors=["#10b981", "#f59e0b"],
                        textprops={"fontsize": 9, "weight": "bold", "color": "white"},
                    )
                    ax2.axis("equal")
                    st.pyplot(fig2, use_container_width=False)
                else:
                    st.info("Not enough data yet.")

            st.markdown("---")
            st.subheader("🧠 Talent Pool Skill Demand")
            st.caption("What skills the candidate pool actually brings — useful for planning roles and training.")
            skill_rows = db.get_skill_demand(limit=12)
            if skill_rows:
                skill_df = pd.DataFrame(skill_rows).rename(columns={
                    "skill_name": "Skill",
                    "category": "Category",
                    "candidate_count": "Candidates",
                })
                st.dataframe(skill_df, use_container_width=True)
            else:
                st.info("No skills recorded yet.")

        with tab_candidate:
            st.subheader("🔍 Candidate Detail")
            email = st.selectbox("Select Candidate", df["Email"].unique())

            if email:
                detail = db.get_candidate_detail(email)
                if not detail:
                    st.warning("Candidate not found.")
                else:
                    user = detail["user"]
                    st.markdown(f"**Email:** {user['email']}")
                    if user["name"]:
                        st.markdown(f"**Name:** {user['name']}")
                    st.markdown(f"**First seen:** {user['created_at'].strftime('%Y-%m-%d %H:%M')}")

                    st.markdown("#### Skills on Profile")
                    if detail["skills"]:
                        skills_df = pd.DataFrame(detail["skills"]).rename(columns={
                            "skill_name": "Skill",
                            "category": "Category",
                            "source": "Source",
                            "created_at": "Added",
                        })
                        st.dataframe(skills_df, use_container_width=True)
                    else:
                        st.info("No skills recorded.")

                    st.markdown("#### Assessment History")
                    for att in detail["assessments"]:
                        acc = float(att["accuracy"]) if att["accuracy"] is not None else 0.0
                        status_icon = "✅" if att["completion_status"] == "completed" else "↻"
                        started = att["started_at"].strftime("%Y-%m-%d %H:%M") if att["started_at"] else "—"
                        st.markdown(
                            f"{status_icon} **Attempt #{att['id']}** — {att['score']}/{att['max_score']} "
                            f"({acc:.1f}%) · {'AI' if att['ai_generated'] else 'Fallback'} · "
                            f"{att['completion_status']} · started {started}"
                        )
                        if att["email_status"]:
                            mail_icon = {"sent": "📧 sent", "failed": "⚠️ failed"}.get(
                                att["email_status"], att["email_status"]
                            )
                            st.caption(f"Report: {mail_icon}")
                        if att["pdf_path"] and os.path.exists(att["pdf_path"]):
                            with open(att["pdf_path"], "rb") as pf:
                                st.download_button(
                                    "📄 Download this report",
                                    data=pf.read(),
                                    file_name=os.path.basename(att["pdf_path"]),
                                    mime="application/pdf",
                                    key=f"dl_report_{att['id']}",
                                )
                        elif att["pdf_path"]:
                            st.caption("📄 Report file no longer on disk.")

                    # Reuse the Phase 5 engine so a recruiter sees the same
                    # matches the candidate does - useful for a screening call.
                    st.markdown("#### Marketplace Fit")
                    if st.button("🎯 Compute role matches for this candidate", key=f"match_{user['id']}"):
                        with st.spinner("Matching..."):
                            candidate_recs = matching_engine.get_recommendations(user["id"])
                        for job in candidate_recs["external_jobs"][:3]:
                            st.markdown(
                                f"- **{job['title']}** — {job['match_percent']}% "
                                f"(gaps: {', '.join(job['missing_skills']) or 'none'})"
                            )
                        for job in candidate_recs["internal_jobs"][:2]:
                            st.markdown(
                                f"- 🔄 **{job['title']}** — {job['match_percent']}% "
                                f"(gaps: {', '.join(job['missing_skills']) or 'none'})"
                            )

        st.markdown("---")

    # Logout is available whether or not there's data
    if st.button("🚪 Logout", use_container_width=True):
        st.session_state.step = 0
        st.session_state.admin_logged = False
        st.session_state.admin_role = None
        st.rerun()

# =====================================================
# STUDENT FLOW - STEP 0: Input
# =====================================================
elif st.session_state.step == 0:
    st.markdown("## Welcome to SkillGauge AI")
    st.markdown("Professional Interview Assessment Platform")
    st.markdown("---")
    
    with st.container():
        method = st.radio(
            "How would you like to provide your profile?",
            ["📝 Enter Skills Manually", "📄 Upload Resume (PDF)"],
            horizontal=True
        )
        
        st.markdown("---")
        
        if method == "📝 Enter Skills Manually":
            st.subheader("Enter Your Technical Skills")
            skills = st.text_area(
                "List your skills (comma separated)",
                placeholder="e.g. Python, JavaScript, React, SQL, AWS, Machine Learning",
                height=100
            )
            st.markdown("**Example:** Python, JavaScript, React, SQL, AWS, Cloud Computing")
            
            st.markdown("---")
            
            st.subheader("Contact Information")
            email = st.text_input(
                "Your Email Address",
                placeholder="your.email@example.com"
            )
            st.session_state.email = email
            
        else:  # Resume Upload
            st.subheader("Upload Your Resume")
            file = st.file_uploader("Choose a PDF file", type="pdf")
            
            if file:
                with st.spinner("📖 Extracting skills from resume..."):
                    skills = extract_skills_from_resume(file)
                
                skills = st.text_area(
                    "Edit extracted skills if needed",
                    value=skills,
                    height=100
                )
                
                st.markdown("---")
                
                st.subheader("Confirm Contact Information")
                email = st.text_input(
                    "Your Email Address",
                    placeholder="your.email@example.com"
                )
                st.session_state.email = email
        
        st.markdown("---")
        
        # Validation and Submit
        col1, col2 = st.columns(2)
        with col1:
            if st.button("🚀 Start Interview", use_container_width=True):
                if not skills.strip():
                    st.error("❌ Please enter your skills.")
                elif not email.strip():
                    st.error("❌ Please enter your email address.")
                else:
                    source = "resume" if method == "📄 Upload Resume (PDF)" else "manual"
                    user_id = db.get_or_create_user(email.strip())
                    categorized = categorize_skills(skills)
                    db.add_categorized_skills(user_id, categorized, source=source)
                    st.session_state.user_id = user_id
                    st.session_state.email = email.strip()

                    # Session resilience (Q6): if this candidate already has an
                    # in-progress attempt (page refresh, dropped connection),
                    # resume it instead of generating a fresh set of questions.
                    existing = db.get_in_progress_assessment(user_id)
                    if existing:
                        st.session_state.assessment_id = existing["id"]
                        st.session_state.mcqs = json.loads(existing["questions_json"])
                        st.session_state.answers = {
                            int(k): v for k, v in json.loads(existing["responses_json"] or "{}").items()
                        }
                        st.session_state.ai_used = bool(existing["ai_generated"])
                        st.session_state.skills_snapshot = existing["skills_snapshot"] or skills.strip()
                        st.session_state.assessment_start_time = existing["started_at"].timestamp()
                        st.session_state.assessment_time_limit = len(st.session_state.mcqs) * 60
                        st.session_state.current_question = min(
                            len(st.session_state.answers), len(st.session_state.mcqs) - 1
                        )
                        st.info("↻ Resuming your in-progress assessment.")
                        st.session_state.step = 1
                        st.rerun()
                    else:
                        st.session_state.skills_snapshot = skills.strip()
                        with st.spinner("🤖 Generating interview questions..."):
                            mcqs = generate_interview_mcqs(skills)

                        if not mcqs:
                            st.error("❌ Failed to generate questions. Please retry.")
                        else:
                            st.session_state.mcqs = mcqs
                            st.session_state.assessment_id = db.start_assessment(
                                user_id=user_id,
                                skills_snapshot=skills.strip(),
                                mcqs=mcqs,
                                ai_generated=st.session_state.ai_used,
                            )
                            st.session_state.assessment_start_time = time.time()
                            st.session_state.assessment_time_limit = len(mcqs) * 60
                            st.session_state.current_question = 0
                            ai_status = "🤖 AI-Generated" if st.session_state.ai_used else "📚 Curated Questions"
                            st.success(f"✅ Questions ready! ({ai_status})")
                            st.session_state.step = 1
                            st.rerun()
        
        with col2:
            st.info("💡 Tip: More specific skills lead to better interview questions!")

# =====================================================
# STUDENT FLOW - STEP 1: MCQ Assessment (one question at a time, timed)
# =====================================================
elif st.session_state.step == 1:
    total_questions = len(st.session_state.mcqs)
    idx = min(st.session_state.current_question, total_questions - 1)

    # --- Timer: auto-submit once time is up ---
    elapsed = time.time() - st.session_state.assessment_start_time
    remaining = st.session_state.assessment_time_limit - elapsed
    if remaining <= 0:
        st.session_state.step = 2
        st.rerun()
    minutes, seconds = divmod(max(0, int(remaining)), 60)

    st.markdown("## 📝 Interview Assessment")
    top_col1, top_col2, top_col3 = st.columns(3)
    with top_col1:
        st.metric("Question", f"{idx + 1} of {total_questions}")
    with top_col2:
        ai_status = "🤖 AI-Generated" if st.session_state.ai_used else "📚 Curated Questions"
        st.metric("Source", ai_status)
    with top_col3:
        st.metric("Time Remaining", f"{minutes:02d}:{seconds:02d}")

    st.progress((idx + 1) / total_questions)
    st.markdown("---")

    q = st.session_state.mcqs[idx]
    st.markdown(f"### {q['q']}")

    options_list = list(q["options"].keys())
    current_answer = st.session_state.answers.get(idx + 1)
    default_index = options_list.index(current_answer) if current_answer in options_list else None

    selected = st.radio(
        "Select your answer:",
        options_list,
        format_func=lambda x: f"{x}) {q['options'][x]}",
        index=default_index,
        key=f"q_radio_{idx}",
        horizontal=False,
    )
    if selected is not None and st.session_state.answers.get(idx + 1) != selected:
        st.session_state.answers[idx + 1] = selected
        # Persist progress immediately - this is what makes a refresh or
        # dropped connection recoverable (Q6).
        db.update_assessment_responses(st.session_state.assessment_id, st.session_state.answers)

    st.markdown("---")
    nav_col1, nav_col2, nav_col3 = st.columns(3)
    with nav_col1:
        if st.button("⬅ Previous", use_container_width=True, disabled=(idx == 0)):
            st.session_state.current_question = idx - 1
            st.rerun()
    with nav_col2:
        if idx < total_questions - 1:
            if st.button("Next ➡", use_container_width=True):
                st.session_state.current_question = idx + 1
                st.rerun()
    with nav_col3:
        if st.button("✓ Submit Assessment", use_container_width=True):
            st.session_state.step = 2
            st.rerun()

# =====================================================
# STUDENT FLOW - STEP 2: Results
# =====================================================
elif st.session_state.step == 2:
    total_questions = len(st.session_state.mcqs)
    score = sum(
        1 for i, q in enumerate(st.session_state.mcqs, 1)
        if st.session_state.answers.get(i) == q["answer"]
    )

    percentage = (score / total_questions) * 100 if total_questions else 0

    pdf = generate_pdf(st.session_state.email, st.session_state.mcqs, st.session_state.answers, score)

    # Try to send email
    email_sent = send_email_report(st.session_state.email, pdf)

    # The assessment row already exists (created 'in_progress' when
    # questions were generated) - finalize it rather than inserting a new
    # one, so a submit after a resumed session updates the same record.
    if not st.session_state.get("saved_to_db"):
        db.complete_assessment(st.session_state.assessment_id, score=score, max_score=total_questions)
        db.save_report(
            st.session_state.assessment_id,
            pdf_path=pdf,
            email_status="sent" if email_sent else "failed",
        )
        st.session_state.saved_to_db = True
    
    # Result Header
    st.markdown("## ✅ Assessment Complete!")
    st.markdown("---")
    
    if email_sent:
        st.success("📧 Report sent to your email successfully!")
    else:
        st.info("📧 PDF report generated (email sending unavailable)")
    
    st.markdown("---")
    
    # Score Display
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("Your Score", f"{score}/{total_questions}")
    with col2:
        st.metric("Percentage", f"{percentage:.1f}%")
    with col3:
        if percentage >= 80:
            status = "🟢 Excellent"
        elif percentage >= 60:
            status = "🟡 Good"
        else:
            status = "🔴 Needs Improvement"
        st.metric("Performance", status)
    
    st.markdown("---")
    
    # Performance Analysis
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("Performance Chart")
        fig, ax = plt.subplots(figsize=(4, 3.5))
        fig.patch.set_alpha(0)
        ax.patch.set_alpha(0)
        colors = ['#10b981', '#ef4444']
        wedges, texts, autotexts = ax.pie(
            [score, total_questions - score],
            labels=["✓ Correct", "✗ Wrong"],
            autopct="%1.1f%%",
            startangle=90,
            colors=colors,
            explode=(0.05, 0.05),
            shadow=False,
            textprops={'fontsize': 10, 'weight': 'bold', 'color': 'white'}
        )
        for autotext in autotexts:
            autotext.set_color('white')
            autotext.set_fontsize(10)
        for text in texts:
            text.set_color('white')
            text.set_fontsize(10)
        ax.axis("equal")
        st.pyplot(fig, use_container_width=False)
    
    with col2:
        st.subheader("Answer Review")
        correct = 0
        for i, q in enumerate(st.session_state.mcqs, 1):
            user_ans = st.session_state.answers.get(i)
            correct_ans = q["answer"]
            is_correct = user_ans == correct_ans
            
            if is_correct:
                st.success(f"Q{i}: ✓ Correct")
            else:
                st.error(f"Q{i}: ✗ Wrong (Your: {user_ans}, Correct: {correct_ans})")
    
    st.markdown("---")
    
    # Download Report
    with open(pdf, "rb") as f:
        st.download_button(
            label="📄 Download Full Report (PDF)",
            data=f.read(),
            file_name=os.path.basename(pdf),
            mime="application/pdf",
            use_container_width=True
        )
    
    st.markdown("---")

    col_next1, col_next2 = st.columns(2)
    with col_next1:
        if st.button("🎯 View My Career Opportunities", use_container_width=True):
            st.session_state.step = 3
            st.rerun()

    with col_next2:
        # Reset Button
        if st.button("🔄 Take Another Assessment", use_container_width=True):
            st.session_state.step = 0
            st.session_state.mcqs = []
            st.session_state.answers = {}
            st.session_state.score = 0
            st.session_state.ai_used = False
            st.session_state.user_id = None
            st.session_state.skills_snapshot = ""
            st.session_state.saved_to_db = False
            st.session_state.recommendations = None
            st.rerun()

# =====================================================
# STUDENT FLOW - STEP 3: Talent Marketplace (Phase 5)
# Candidate-job matching, internal mobility, mentors,
# learning suggestions, and a career path.
# =====================================================
elif st.session_state.step == 3:
    st.markdown("## 🎯 Your Career Opportunities")
    st.markdown("Matched against your validated skill profile and assessment performance.")
    st.markdown("---")

    if not st.session_state.user_id:
        st.warning("⚠️ Complete an assessment first to see your matches.")
    else:
        # Cache in session state: the embedding/AI calls are the expensive
        # part, and a rerun (tab switch, button click) shouldn't re-run them.
        if not st.session_state.get("recommendations"):
            with st.spinner("🔎 Matching you to roles, mentors and learning paths..."):
                st.session_state.recommendations = matching_engine.get_recommendations(
                    st.session_state.user_id
                )

        recs = st.session_state.recommendations
        profile = recs["profile"]

        col_p1, col_p2, col_p3 = st.columns(3)
        with col_p1:
            st.metric("Skills On Profile", len(profile["skills"]))
        with col_p2:
            st.metric("Assessment Score", f"{profile['assessment_accuracy'] * 100:.0f}%")
        with col_p3:
            best = recs["external_jobs"][0]["match_percent"] if recs["external_jobs"] else 0
            st.metric("Best Role Match", f"{best:.0f}%")

        st.markdown("---")

        tab_jobs, tab_internal, tab_mentors, tab_learning, tab_path = st.tabs([
            "💼 Job Matches",
            "🔄 Internal Mobility",
            "👥 Mentors",
            "📚 Learning",
            "🧭 Career Path",
        ])

        def render_jobs(jobs, empty_msg):
            if not jobs:
                st.info(empty_msg)
                return
            for job in jobs:
                with st.container():
                    st.markdown(f"### {job['title']}  —  {job['match_percent']}% match")
                    st.progress(min(1.0, job["match_percent"] / 100))
                    st.markdown(f"**Department:** {job['department']}")
                    st.markdown(job["description"])
                    if job["matched_skills"]:
                        st.markdown(f"✅ **You already have:** {', '.join(job['matched_skills'])}")
                    if job["missing_skills"]:
                        st.markdown(f"📈 **To strengthen:** {', '.join(job['missing_skills'])}")
                    st.markdown("---")

        with tab_jobs:
            st.markdown("#### Roles that fit your current profile")
            render_jobs(recs["external_jobs"], "No job matches available yet.")

        with tab_internal:
            st.markdown("#### Internal moves you could grow into")
            render_jobs(recs["internal_jobs"], "No internal mobility options available yet.")

        with tab_mentors:
            st.markdown("#### Mentors who can help you grow")
            if not recs["mentors"]:
                st.info("No mentors available yet.")
            for mentor in recs["mentors"]:
                st.markdown(f"### {mentor['name']}  —  {mentor['relevance_percent']}% relevance")
                st.markdown(f"*{mentor['bio']}*")
                st.markdown(f"**Expertise:** {mentor['expertise']}")
                if mentor["shared_skills"]:
                    st.markdown(f"🤝 **Common ground:** {', '.join(mentor['shared_skills'])}")
                if mentor["can_teach"]:
                    st.markdown(f"🎓 **Can teach you:** {', '.join(mentor['can_teach'])}")
                st.markdown("---")

        with tab_learning:
            st.markdown("#### Recommended next, highest priority first")
            if not recs["learning"]:
                st.info("No learning resources available yet.")
            for item in recs["learning"]:
                badge = {3: "🔴 High", 2: "🟡 Medium", 1: "🟢 Optional"}.get(item["priority"], "")
                st.markdown(f"### {item['title']}")
                st.markdown(f"**Priority:** {badge}  |  **Skill:** {item['skill_tag']}")
                st.markdown(f"_{item['reason']}_")
                st.markdown(item["description"])
                if item["url"]:
                    st.markdown(f"[Open course]({item['url']})")
                st.markdown("---")

        with tab_path:
            st.markdown("#### Your suggested career path")
            st.markdown(recs["career_path"])

        st.markdown("---")
        col_b1, col_b2 = st.columns(2)
        with col_b1:
            if st.button("⬅ Back to Results", use_container_width=True):
                st.session_state.step = 2
                st.rerun()
        with col_b2:
            if st.button("🔄 Take Another Assessment", use_container_width=True, key="reset_from_marketplace"):
                st.session_state.step = 0
                st.session_state.mcqs = []
                st.session_state.answers = {}
                st.session_state.score = 0
                st.session_state.ai_used = False
                st.session_state.user_id = None
                st.session_state.skills_snapshot = ""
                st.session_state.saved_to_db = False
                st.session_state.recommendations = None
                st.rerun()
