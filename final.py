import streamlit as st
import os, json, re, smtplib, random
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

from openpyxl import Workbook, load_workbook
from google import genai

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
# PATHS
# =====================================================
RESULT_DIR = "results"
PDF_DIR = "pdf_reports"
EXCEL_PATH = os.path.join(RESULT_DIR, "student_scores.xlsx")

os.makedirs(RESULT_DIR, exist_ok=True)
os.makedirs(PDF_DIR, exist_ok=True)

# =====================================================
# GEMINI CONFIG
# =====================================================
# client = genai.Client(api_key="AIzaSyCS9x37uOK1CJkgrBnDdeU93OEjUDtLtwE")
# AIzaSyA167FYg2wh7EXEcIBK2KYkXp0pD7cx1Mk
MODEL_NAME = "models/gemini-flash-latest"

# =====================================================
# EMAIL CONFIG
# =====================================================
SENDER_EMAIL = "noreply.civicconnect@gmail.com"
APP_PASSWORD = "hvpwfvoouycoaqgo"

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
        "ai_used": False  # Track if AI or fallback was used
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
        res = client.models.generate_content(model=MODEL_NAME, contents=prompt)
        return res.text.strip()
    except:
        st.warning("⚠️ AI unavailable. Please enter skills manually.")
        return ""

def generate_interview_mcqs(skills):
    """
    AI-first MCQ generation with guaranteed fallback
    """

    prompt = f"""
Return ONLY a valid JSON ARRAY.

FORMAT:
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

Generate exactly 25 interview MCQs based on:
{skills}
"""

    for attempt in range(2):
        try:
            res = client.models.generate_content(
                model=MODEL_NAME,
                contents=prompt
            )

            mcqs = extract_json_flexible(res.text)

            if mcqs and isinstance(mcqs, list) and len(mcqs) >= 25:
                st.session_state.ai_used = True
                return mcqs[:25]

        except Exception as e:
            print(f"[AI ERROR] Attempt {attempt+1}:", e)

    # FALLBACK (UNCHANGED)
    st.warning("⚠️ AI service temporarily unavailable. Using curated question bank.")
    st.session_state.ai_used = False
    return random.sample(QUESTION_BANK, 25)


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
# SAVE RESULT
# =====================================================
def save_to_excel(email, score):
    if not os.path.exists(EXCEL_PATH):
        wb = Workbook()
        ws = wb.active
        ws.append(["Email", "Score", "Max Marks", "Date", "Time"])
        wb.save(EXCEL_PATH)

    wb = load_workbook(EXCEL_PATH)
    ws = wb.active
    now = datetime.now()
    ws.append([email, score, 25, now.strftime("%Y-%m-%d"), now.strftime("%H:%M:%S")])
    wb.save(EXCEL_PATH)

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

    story.append(Spacer(1, 12))
    story.append(Paragraph(f"<b>FINAL SCORE: {score}/25</b>", styles["Normal"]))
    percentage = (score / 25) * 100
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
                    if username == "admin" and password == "admin123":
                        st.session_state.admin_logged = True
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
    st.markdown("## 📊 Admin Dashboard")
    st.markdown("---")
    
    if not os.path.exists(EXCEL_PATH):
        st.warning("⚠️ No assessment data yet.")
    else:
        df = pd.read_excel(EXCEL_PATH)
        
        # Summary Metrics
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("Total Assessments", len(df))
        with col2:
            st.metric("Avg Score", f"{df['Score'].mean():.1f}/25")
        with col3:
            st.metric("Highest Score", f"{df['Score'].max()}/25")
        with col4:
            st.metric("Lowest Score", f"{df['Score'].min()}/25")
        
        st.markdown("---")
        
        # Data Table
        st.subheader("📋 Assessment Records")
        st.dataframe(df, use_container_width=True)
        
        st.markdown("---")
        
        # Download Button
        with open(EXCEL_PATH, "rb") as f:
            st.download_button(
                label="📥 Download Excel Report",
                data=f.read(),
                file_name="student_scores.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )   
        
        st.markdown("---")
        
        # Individual Analysis
        st.subheader("📈 Student Performance Analysis")
        email = st.selectbox("Select Student Email", df["Email"].unique())
        
        if email:
            latest = df[df["Email"] == email].iloc[-1]
            
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("Score", f"{int(latest['Score'])}/25")
            with col2:
                percentage = (latest['Score'] / 25) * 100
                st.metric("Percentage", f"{percentage:.1f}%")
            with col3:
                st.metric("Assessment Date", latest['Date'])
            
            # Pie Chart - Smaller Size
            fig, ax = plt.subplots(figsize=(4, 3))
            fig.patch.set_alpha(0)
            ax.patch.set_alpha(0)
            colors = ['#10b981', '#ef4444']
            wedges, texts, autotexts = ax.pie(
                [latest["Score"], 25 - latest["Score"]],
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
            ax.axis("equal")
            st.pyplot(fig, use_container_width=False)
        
        st.markdown("---")
        
        # Logout Button
        if st.button("🚪 Logout", use_container_width=True):
            st.session_state.step = 0
            st.session_state.admin_logged = False
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
                    with st.spinner("🤖 Generating interview questions..."):
                        mcqs = generate_interview_mcqs(skills)
                    
                    if not mcqs:
                        st.error("❌ Failed to generate questions. Please retry.")
                    else:
                        st.session_state.mcqs = mcqs
                        ai_status = "🤖 AI-Generated" if st.session_state.ai_used else "📚 Curated Questions"
                        st.success(f"✅ Questions ready! ({ai_status})")
                        st.session_state.step = 1
                        st.rerun()
        
        with col2:
            st.info("💡 Tip: More specific skills lead to better interview questions!")

# =====================================================
# STUDENT FLOW - STEP 1: MCQ Assessment
# =====================================================
elif st.session_state.step == 1:
    st.markdown("## 📝 Interview Assessment")
    st.markdown(f"**Total Questions:** {len(st.session_state.mcqs)}")
    ai_status = "🤖 AI-Generated" if st.session_state.ai_used else "📚 Curated Questions"
    st.markdown(f"**Question Source:** {ai_status}")
    st.markdown("---")
    
    progress = 0
    for i, q in enumerate(st.session_state.mcqs, 1):
        with st.container():
            # Progress indicator
            progress_percent = (i / len(st.session_state.mcqs)) * 100
            st.progress(progress_percent / 100)
            
            # Question number and text
            st.markdown(f"### Question {i} of {len(st.session_state.mcqs)}")
            st.markdown(f"**{q['q']}**", unsafe_allow_html=True)
            
            # Options
            options_list = list(q["options"].keys())
            default_index = 0
            if i in st.session_state.answers:
                try:
                    default_index = options_list.index(st.session_state.answers[i])
                except:
                    default_index = 0
            
            selected = st.radio(
                "Select your answer:",
                options_list,
                format_func=lambda x: f"{x}) {q['options'][x]}",
                index=default_index,
                key=f"q{i}",
                horizontal=False
            )
            st.session_state.answers[i] = selected
            st.markdown("---")
    
    # Submit Button
    if st.button("✓ Submit Assessment", use_container_width=True):
        st.session_state.step = 2
        st.rerun()

# =====================================================
# STUDENT FLOW - STEP 2: Results
# =====================================================
elif st.session_state.step == 2:
    score = sum(
        1 for i, q in enumerate(st.session_state.mcqs, 1)
        if st.session_state.answers.get(i) == q["answer"]
    )
    
    percentage = (score / 25) * 100
    
    pdf = generate_pdf(st.session_state.email, st.session_state.mcqs, st.session_state.answers, score)
    save_to_excel(st.session_state.email, score)
    
    # Try to send email
    email_sent = send_email_report(st.session_state.email, pdf)
    
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
        st.metric("Your Score", f"{score}/25")
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
            [score, 25 - score],
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
    
    # Reset Button
    if st.button("🔄 Take Another Assessment", use_container_width=True):
        st.session_state.step = 0
        st.session_state.mcqs = []
        st.session_state.answers = {}
        st.session_state.score = 0
        st.session_state.ai_used = False
        st.rerun()
