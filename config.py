"""
Centralized configuration, loaded from environment variables (.env).

Never hardcode secrets here or anywhere else in the codebase - add them
to .env (git-ignored) and read them through this module instead.
"""
import os
from dotenv import load_dotenv

load_dotenv()

# =====================================================
# AZURE OPENAI
# =====================================================
AZURE_OPENAI_ENDPOINT = os.environ.get("AZURE_OPENAI_ENDPOINT", "")
AZURE_OPENAI_API_KEY = os.environ.get("AZURE_OPENAI_API_KEY", "")
AZURE_OPENAI_API_VERSION = os.environ.get("AZURE_OPENAI_API_VERSION", "2024-12-01-preview")
AZURE_OPENAI_CHAT_DEPLOYMENT = os.environ.get("AZURE_OPENAI_CHAT_DEPLOYMENT", "")
AZURE_OPENAI_EMBEDDING_DEPLOYMENT = os.environ.get("AZURE_OPENAI_EMBEDDING_DEPLOYMENT", "")

# =====================================================
# EMAIL (Gmail SMTP)
# =====================================================
SENDER_EMAIL = os.environ.get("SENDER_EMAIL", "")
EMAIL_APP_PASSWORD = os.environ.get("EMAIL_APP_PASSWORD", "")

# =====================================================
# ADMIN (temporary - replaced by real auth in a later phase)
# =====================================================
ADMIN_USERNAME = os.environ.get("ADMIN_USERNAME", "admin")
ADMIN_PASSWORD = os.environ.get("ADMIN_PASSWORD", "admin123")

# Recruiter account - same dashboard, read-only (no exports/system view).
RECRUITER_USERNAME = os.environ.get("RECRUITER_USERNAME", "recruiter")
RECRUITER_PASSWORD = os.environ.get("RECRUITER_PASSWORD", "recruiter123")

# =====================================================
# ASSESSMENT
# =====================================================
# Total time limit = question count x this. Configurable so the duration
# can be tuned per cohort without a code change.
ASSESSMENT_SECONDS_PER_QUESTION = int(os.environ.get("ASSESSMENT_SECONDS_PER_QUESTION", "60"))

# =====================================================
# MYSQL
# =====================================================
MYSQL_HOST = os.environ.get("MYSQL_HOST", "localhost")
MYSQL_PORT = int(os.environ.get("MYSQL_PORT", "3306"))
MYSQL_DATABASE = os.environ.get("MYSQL_DATABASE", "skillup_db")
MYSQL_USER = os.environ.get("MYSQL_USER", "skillup_app")
MYSQL_PASSWORD = os.environ.get("MYSQL_PASSWORD", "")

# =====================================================
# PATHS
# =====================================================
RESULT_DIR = "results"
PDF_DIR = "pdf_reports"
# Generated on demand from the DB by export_excel_from_db(). Kept separate
# from any hand-maintained/historical spreadsheet in results/ so a download
# never overwrites data the DB isn't the source of truth for.
EXCEL_PATH = os.path.join(RESULT_DIR, "skillgauge_export.xlsx")

os.makedirs(RESULT_DIR, exist_ok=True)
os.makedirs(PDF_DIR, exist_ok=True)


def validate_azure_openai():
    """Raise a clear error early if required Azure OpenAI settings are missing."""
    missing = [
        name for name, val in [
            ("AZURE_OPENAI_ENDPOINT", AZURE_OPENAI_ENDPOINT),
            ("AZURE_OPENAI_API_KEY", AZURE_OPENAI_API_KEY),
            ("AZURE_OPENAI_CHAT_DEPLOYMENT", AZURE_OPENAI_CHAT_DEPLOYMENT),
            ("AZURE_OPENAI_EMBEDDING_DEPLOYMENT", AZURE_OPENAI_EMBEDDING_DEPLOYMENT),
        ]
        if not val
    ]
    if missing:
        raise RuntimeError(
            "Missing required .env settings: "
            + ", ".join(missing)
            + ". Copy .env.example to .env and fill in real values."
        )
