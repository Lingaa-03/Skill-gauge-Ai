# SkillGauge AI

AI-powered recruitment and talent marketplace. Candidates submit skills
(typed or parsed from a resume), take a timed AI-generated assessment, and
receive a scored PDF report plus matched job roles, internal mobility
options, learning recommendations and a career path. Recruiters and admins
get a dashboard with cohort analytics and per-candidate drill-down.

Built with Streamlit, MySQL and Azure OpenAI (`gpt-5.6-sol` for generation,
`text-embedding-3-large` for skill matching).

---

## Running it in VS Code

### One-time setup

**1. Install dependencies**

```bash
python -m pip install -r requirements.txt
```

**2. Make sure MySQL is running**

The app needs the local MySQL server up. Check it in PowerShell:

```bash
Get-Service MySQL80
```

If `Status` is `Stopped`, start it:

```bash
net start MySQL80
```

**3. Create the database and app user** (only needed once, in MySQL
Workbench or `mysql -u root -p`)

```sql
CREATE DATABASE IF NOT EXISTS skillup_db CHARACTER SET utf8mb4;
CREATE USER IF NOT EXISTS 'skillup_app'@'localhost' IDENTIFIED BY 'your-password';
GRANT ALL PRIVILEGES ON skillup_db.* TO 'skillup_app'@'localhost';
FLUSH PRIVILEGES;
```

Tables are created automatically on first app start.

**4. Create `.env`**

Copy `.env.example` to `.env` and fill in the real values (Azure key, MySQL
password). `.env` is git-ignored and must never be committed.

### Starting the app

Three ways, all equivalent — pick whichever you prefer:

| How | What to do |
|---|---|
| **F5** | Press `F5` (uses `.vscode/launch.json`) |
| **Ctrl+Shift+B** | Runs the default build task (`.vscode/tasks.json`) |
| **Terminal** | `python -m streamlit run final.py` |

Then open <http://localhost:8501>.

> **Do not run `python final.py`.** Streamlit apps must be started through
> the Streamlit CLI. Running the file directly executes the script outside
> Streamlit's runtime and floods the console with
> `missing ScriptRunContext!` warnings without ever serving the app.

To stop it: `Ctrl+C` in the terminal running it.

### If something breaks

Run the preflight check — it verifies packages, `.env`, MySQL connectivity
and both Azure deployments, and prints a clear PASS/FAIL for each:

```bash
python check_setup.py
```

It never prints secret values, only whether they're set.

---

## Logging in

| Role | Username | Password | Access |
|---|---|---|---|
| Candidate | — | — | No login; enters skills + email directly |
| Admin | `admin` | from `.env` | Full dashboard incl. Excel export |
| Recruiter | `recruiter` | from `.env` | Read-only dashboard, no export |

Reach the login via the 🔐 button in the top-right. Passwords are stored as
bcrypt hashes; the values in `.env` only seed the accounts on first run.

**Change the default passwords in `.env` before using this for real.**

---

## Configuration

All settings live in `.env` (see `.env.example` for the full list).

| Setting | Purpose |
|---|---|
| `AZURE_OPENAI_*` | Endpoint, key, and the chat/embedding deployment names |
| `MYSQL_*` | Database connection |
| `SENDER_EMAIL`, `EMAIL_APP_PASSWORD` | Gmail SMTP for emailing reports |
| `ADMIN_*`, `RECRUITER_*` | Seeded dashboard accounts |
| `ASSESSMENT_TOTAL_MINUTES` | Time limit per assessment (default 30) |

Config is read at import time, so **restart the app after editing `.env`**.

---

## Project layout

```
final.py                      Streamlit app (all UI + flow)
config.py                     Settings loaded from .env
db.py                         MySQL schema, queries, auth, seed data
check_setup.py                Preflight diagnostics
services/
  ai_client.py                Azure OpenAI wrapper (chat + embeddings)
  matching_engine.py          Job/mobility/learning/career-path matching
.streamlit/config.toml        Theme (pinned light, app palette)
.vscode/                      F5 / Ctrl+Shift+B run configs
pdf_reports/                  Generated reports (git-ignored)
results/                      Generated Excel export (git-ignored)
```

MySQL is the source of truth. Excel and PDFs are generated from it on
demand, which is why they're git-ignored rather than committed.
