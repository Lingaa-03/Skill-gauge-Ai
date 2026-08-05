"""
Preflight check - run this before starting the app if something's wrong.

Verifies the four things the app needs: Python packages, .env settings,
a reachable MySQL, and a working Azure OpenAI deployment. Prints a clear
PASS/FAIL per item instead of failing with a stack trace mid-request.

    python check_setup.py
"""
import importlib
import os
import sys

OK = "[ OK ]"
BAD = "[FAIL]"
failures = []


def report(ok, label, detail=""):
    print(f"{OK if ok else BAD} {label}" + (f" - {detail}" if detail else ""))
    if not ok:
        failures.append(label)


print("=" * 60)
print("SkillGauge AI - setup check")
print("=" * 60)

# ---------- 1. Packages ----------
print("\n1. Python packages")
for module, pip_name in [
    ("streamlit", "streamlit"),
    ("openai", "openai"),
    ("dotenv", "python-dotenv"),
    ("mysql.connector", "mysql-connector-python"),
    ("bcrypt", "bcrypt"),
    ("reportlab", "reportlab"),
    ("openpyxl", "openpyxl"),
    ("PyPDF2", "PyPDF2"),
    ("pandas", "pandas"),
    ("matplotlib", "matplotlib"),
]:
    try:
        importlib.import_module(module)
        report(True, pip_name)
    except ImportError:
        report(False, pip_name, "missing - run: python -m pip install -r requirements.txt")

if failures:
    print("\nInstall the missing packages first, then re-run this check.")
    sys.exit(1)

# ---------- 2. .env ----------
print("\n2. Configuration (.env)")
if not os.path.exists(".env"):
    report(False, ".env file", "not found - copy .env.example to .env and fill it in")
    sys.exit(1)
report(True, ".env file")

import config  # noqa: E402  (imported after the package check on purpose)

for name, value in [
    ("AZURE_OPENAI_ENDPOINT", config.AZURE_OPENAI_ENDPOINT),
    ("AZURE_OPENAI_API_KEY", config.AZURE_OPENAI_API_KEY),
    ("AZURE_OPENAI_CHAT_DEPLOYMENT", config.AZURE_OPENAI_CHAT_DEPLOYMENT),
    ("AZURE_OPENAI_EMBEDDING_DEPLOYMENT", config.AZURE_OPENAI_EMBEDDING_DEPLOYMENT),
    ("MYSQL_DATABASE", config.MYSQL_DATABASE),
    ("MYSQL_USER", config.MYSQL_USER),
    ("MYSQL_PASSWORD", config.MYSQL_PASSWORD),
]:
    # Never print secret values, only whether they're set.
    report(bool(value), name, "" if value else "not set in .env")

print(f"       assessment length: {config.ASSESSMENT_TOTAL_MINUTES} minutes")

# ---------- 3. MySQL ----------
print("\n3. MySQL")
try:
    import db

    conn = db.get_connection()
    cur = conn.cursor()
    cur.execute("SELECT DATABASE()")
    dbname = cur.fetchone()[0]
    cur.execute("SHOW TABLES")
    tables = sorted(r[0] for r in cur.fetchall())
    cur.close()
    conn.close()
    report(True, "connection", f"database '{dbname}'")
    report(True, "tables", f"{len(tables)} found: {', '.join(tables) if tables else '(none yet - created on first run)'}")
except Exception as e:
    report(False, "connection", f"{type(e).__name__}: {e}")
    print("       Is the MySQL80 service running? Are the credentials in .env correct?")

# ---------- 4. Azure OpenAI ----------
print("\n4. Azure OpenAI")
try:
    from services import ai_client

    text = ai_client.chat_complete("Reply with the single word: OK", max_tokens=2000, retries=1)
    report(bool(text and text.strip()), "chat deployment", config.AZURE_OPENAI_CHAT_DEPLOYMENT)
except Exception as e:
    report(False, "chat deployment", f"{type(e).__name__}: {e}")

try:
    vec = ai_client.embed("hello")
    report(len(vec) > 0, "embedding deployment", f"{len(vec)}-dim vector")
except Exception as e:
    report(False, "embedding deployment", f"{type(e).__name__}: {e}")

# ---------- Summary ----------
print("\n" + "=" * 60)
if failures:
    print(f"{len(failures)} problem(s): {', '.join(failures)}")
    sys.exit(1)
print("All checks passed. Start the app with:  python -m streamlit run final.py")
print("=" * 60)
