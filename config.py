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
EXCEL_PATH = os.path.join(RESULT_DIR, "student_scores.xlsx")

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
