# 📧 RecEmailed

Create Gmail draft emails for recruiter outreach — with a simple web UI.

Upload your recruiter data (Excel) and resume (PDF), customise your email template, and create drafts in your Gmail inbox. **Nothing is sent automatically** — you review and send each draft yourself.

![Python](https://img.shields.io/badge/Python-3.10+-blue?logo=python)
![Streamlit](https://img.shields.io/badge/Streamlit-1.30+-red?logo=streamlit)
![Gmail API](https://img.shields.io/badge/Gmail-API-green?logo=gmail)

---

## How It Works

1. **Upload** your Excel file (with recruiter Name, Email, Company, Role) and resume PDF
2. **Customise** the email template with `{Name}`, `{Company}`, `{Role}` placeholders
3. **Preview** what the email looks like for the first recruiter
4. **Click "Create Drafts"** → drafts appear in your Gmail → Drafts folder
5. **Open Gmail**, review each draft, and hit Send when you're ready

---

## Quick Start

### 1. Clone & Install

```bash
git clone https://github.com/YOUR_USERNAME/RecEmailed.git
cd RecEmailed
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
```

### 2. Set Up Gmail API (One-Time)

1. Go to [Google Cloud Console](https://console.cloud.google.com/)
2. Create a project → Enable **Gmail API**
3. Go to **APIs & Services → Credentials**
4. Click **Create Credentials → OAuth client ID → Desktop app**
5. Download the JSON file and save it as **`credentials.json`** in this folder

> **Note:** If asked to configure a consent screen, choose "External", fill in the app name, and add your email as a test user.

### 3. Run the App

```bash
source venv/bin/activate
streamlit run app.py
```

A browser tab opens with the UI. On the first run, you'll be asked to log in with Google — after that, your token is cached.

---

## Excel File Format

Your `.xlsx` file needs these exact column headers:

| Name         | Email               | Company | Role              |
|--------------|---------------------|---------|-------------------|
| Jane Smith   | jane@example.com    | Google  | Software Engineer |
| John Doe     | john@example.com    | Meta    | Backend Developer |

Run `python create_sample_data.py` to generate a sample file.

---

## CLI Mode (Optional)

If you prefer a no-UI approach, you can also run the script directly:

```bash
python draft_emails.py
```

This reads from `recruiters.xlsx` and `resume.pdf` in the project folder.

---

## Project Structure

```
RecEmailed/
├── app.py                 # Streamlit web UI (main entry point)
├── draft_emails.py        # CLI version (alternative)
├── create_sample_data.py  # Generates a sample recruiters.xlsx
├── requirements.txt       # Python dependencies
├── credentials.json       # 🔑 YOUR Gmail OAuth credentials (not committed)
├── token.json             # 🔒 Auto-generated after first login (not committed)
└── .gitignore             # Keeps credentials out of git
```

---

## Security

- `credentials.json` and `token.json` are in `.gitignore` — they won't be committed
- The app only requests `gmail.compose` scope (create drafts) — no access to read your inbox
- **Nothing is sent automatically** — you always send manually from Gmail

---

## License

MIT — do whatever you want with it.
