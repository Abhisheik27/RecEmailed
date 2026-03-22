# 📧 RecEmailed

Create Gmail draft emails for recruiter outreach — powered by Google Sheets.

Manage your recruiter list, customise email templates, and create drafts in Gmail. **Nothing is sent automatically** — you review and send each draft yourself.

![Python](https://img.shields.io/badge/Python-3.10+-blue?logo=python)
![Streamlit](https://img.shields.io/badge/Streamlit-1.30+-red?logo=streamlit)
![Gmail API](https://img.shields.io/badge/Gmail-API-green?logo=gmail)
![Google Sheets](https://img.shields.io/badge/Google_Sheets-API-blue?logo=googlesheets)

---

## Features

- 🔗 **Google Sheets integration** — your recruiter data lives in a Sheet, not a local file
- ➕ **Add / edit / search / delete** recruiters directly from the app
- 📋 **Status tracking** — Not Contacted → Drafted → In-contact
- 🎯 **Dual templates** — role-specific (when role is known) and generic (when it's not)
- ⚠️ **Duplicate protection** — warns before re-drafting already-contacted recruiters
- 📜 **History log** — tracks every run (date, count, recruiter names) in a History tab
- 📎 **Resume attachment** — PDF attached to every draft
- 🚀 **Streamlit Cloud ready** — deploy and access from any browser

---

## Quick Start

### 1. Clone & Install

```bash
git clone https://github.com/Abhisheik27/RecEmailed.git
cd RecEmailed
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
```

### 2. Set Up Google Cloud (One-Time)

1. Go to [Google Cloud Console](https://console.cloud.google.com/) → Create a project
2. Enable **Gmail API** and **Google Sheets API**
3. Set up **OAuth consent screen** (External) → add your email as a test user
4. Create **OAuth client ID** (Desktop app) → download the JSON
5. Rename to `credentials.json` and place in the project folder

### 3. Create Your Google Sheet

Create a new Google Sheet with these headers in Row 1:

| Name | Email | Company | Role | Status |
|------|-------|---------|------|--------|

### 4. Run the App

```bash
./run.sh
```

Or manually:
```bash
source venv/bin/activate
streamlit run app.py
```

On first run, log in with Google when prompted. Your token is cached after that.

---

## How It Works

1. **Paste** your Google Sheet URL in the sidebar → Connect
2. **Add** recruiters (in-app or directly in Google Sheets)
3. **Upload** your resume PDF
4. **Edit** the email template (role-specific or generic)
5. **Select** who to draft → Click **Create Drafts**
6. **Open Gmail → Drafts** → review and send

---

## Deploy to Streamlit Cloud

1. Push to GitHub
2. Go to [share.streamlit.io](https://share.streamlit.io) → New app → select your repo
3. Add secrets in **Settings → Secrets**:
   ```toml
   sheet_url = "https://docs.google.com/spreadsheets/d/YOUR_SHEET_ID/edit"
   google_token = '<paste contents of token.json>'
   ```
4. Deploy 🚀

---

## Project Structure

```
RecEmailed/
├── app.py                          # Streamlit web app (main entry point)
├── requirements.txt                # Python dependencies
├── run.sh                          # Local launcher script
├── .gitignore                      # Keeps credentials out of git
└── .streamlit/
    ├── config.toml                 # Theme + server config
    └── secrets.toml.example        # Template for cloud deployment secrets
```

---

## Security

- `credentials.json` and `token.json` are gitignored — never committed
- The app requests only `gmail.compose` + `spreadsheets` scopes
- **Drafts only** — nothing is sent automatically, ever

---

## License

MIT
