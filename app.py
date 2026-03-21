"""
app.py
------
Streamlit UI for RecEmailed — Gmail Draft Creator for Recruiter Outreach.

Features:
    - Google Sheets integration as the persistent data source
    - Add / search / delete recruiters directly from the app
    - Status tracking: Not Contacted → Drafted → In-contact
    - Select which recruiters to create drafts for
    - Upload resume PDF (attached to every draft)
    - Editable email template with {Name}, {Company}, {Role} placeholders
    - Preview emails before creating drafts
    - Auto-updates status in the Sheet after creating drafts

Run with:
    streamlit run app.py
"""

import base64
import json
import os
import re
from datetime import datetime
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

import gspread
import pandas as pd
import streamlit as st
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build


# ---------------------------------------------------------------------------
# CONSTANTS
# ---------------------------------------------------------------------------

# Gmail + Google Sheets scopes (both use the same OAuth token)
SCOPES = [
    "https://www.googleapis.com/auth/gmail.compose",
    "https://www.googleapis.com/auth/spreadsheets",
]

# OAuth file paths
CREDENTIALS_FILE = "credentials.json"
TOKEN_FILE = "token.json"

# Status values
STATUS_NOT_CONTACTED = "Not Contacted"
STATUS_DRAFTED = "Drafted"
STATUS_IN_CONTACT = "In-contact"
ALL_STATUSES = [STATUS_NOT_CONTACTED, STATUS_DRAFTED, STATUS_IN_CONTACT]

# Required columns in the Google Sheet
REQUIRED_COLUMNS = ["Name", "Email", "Company", "Role", "Status"]

# Default email template
DEFAULT_SUBJECT = "Interested in {Role} opportunity at {Company}"

DEFAULT_TEMPLATE = """\
Hi {Name},

I hope this message finds you well! I came across the {Role} position \
at {Company} and I'm very excited about the opportunity.

I believe my skills and experience align well with what your team is \
looking for. I've attached my resume for your reference and would love \
the chance to discuss how I can contribute to {Company}.

Please let me know if there's a convenient time to chat. I'm happy to \
work around your schedule.

Thank you for your time, and I look forward to hearing from you!

Best regards,
[Your Name]\
"""


# ---------------------------------------------------------------------------
# AUTH & CONNECTIONS
# ---------------------------------------------------------------------------


def get_credentials():
    """
    Get OAuth2 credentials, refreshing or creating as needed.

    Supports two modes:
        1. LOCAL: Uses credentials.json + token.json files
        2. STREAMLIT CLOUD: Reads from st.secrets (no files needed)

    Returns:
        google.oauth2.credentials.Credentials object, or None if
        credentials are missing.
    """
    creds = None

    # --- Mode 1: Streamlit Cloud (secrets-based) ---
    # If running on Streamlit Cloud, credentials are stored in secrets.toml
    # Wrapped in try/except because st.secrets throws if no secrets file exists
    try:
        if "google_token" in st.secrets:
            try:
                token_data = json.loads(st.secrets["google_token"])
                creds = Credentials.from_authorized_user_info(token_data, SCOPES)

                # Refresh if expired
                if creds and creds.expired and creds.refresh_token:
                    creds.refresh(Request())

                return creds
            except Exception:
                return None
    except Exception:
        pass  # No secrets file — fall through to local auth

    # --- Mode 2: Local (file-based) ---
    if os.path.exists(TOKEN_FILE):
        creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if not os.path.exists(CREDENTIALS_FILE):
                return None
            flow = InstalledAppFlow.from_client_secrets_file(
                CREDENTIALS_FILE, SCOPES
            )
            creds = flow.run_local_server(port=0)

        with open(TOKEN_FILE, "w") as f:
            f.write(creds.to_json())

    return creds


def get_gmail_service(creds):
    """Build Gmail API service from credentials."""
    return build("gmail", "v1", credentials=creds)


def get_gsheet(creds, sheet_url: str):
    """
    Open a Google Sheet by URL using the OAuth credentials.

    Args:
        creds:     OAuth2 credentials
        sheet_url: Full URL of the Google Sheet

    Returns:
        Tuple of (first_worksheet, spreadsheet_object).
        We return the spreadsheet too so we can access the History tab.
    """
    gc = gspread.authorize(creds)
    spreadsheet = gc.open_by_url(sheet_url)
    return spreadsheet.sheet1, spreadsheet


# ---------------------------------------------------------------------------
# GOOGLE SHEET OPERATIONS
# ---------------------------------------------------------------------------


def load_recruiters(worksheet) -> pd.DataFrame:
    """
    Load all recruiter data from the Google Sheet into a DataFrame.

    Uses get_all_values() instead of get_all_records() to handle
    sheets with extra empty columns (which cause duplicate header errors).

    If the sheet is empty, initialises it with the required headers.
    """
    all_values = worksheet.get_all_values()

    # Empty sheet — set up headers
    if not all_values:
        worksheet.update("A1:E1", [REQUIRED_COLUMNS])
        return pd.DataFrame(columns=REQUIRED_COLUMNS)

    headers = [str(h).strip().title() for h in all_values[0]]

    # Check if required columns exist
    required_set = {"Name", "Email", "Company", "Role"}
    if not required_set.issubset(set(headers)):
        # Headers don't match — set them up
        if all(h == "" for h in headers):
            worksheet.update("A1:E1", [REQUIRED_COLUMNS])
            return pd.DataFrame(columns=REQUIRED_COLUMNS)

    # Find column indices for required fields
    col_indices = {}
    for col_name in REQUIRED_COLUMNS:
        if col_name in headers:
            col_indices[col_name] = headers.index(col_name)

    # If Status column doesn't exist, we'll add it
    has_status = "Status" in col_indices

    # Parse data rows
    rows = []
    for row in all_values[1:]:
        if not any(cell.strip() for cell in row):
            continue  # Skip completely empty rows

        record = {}
        for col_name in ["Name", "Email", "Company", "Role"]:
            if col_name in col_indices and col_indices[col_name] < len(row):
                record[col_name] = str(row[col_indices[col_name]]).strip()
            else:
                record[col_name] = ""

        if has_status and col_indices["Status"] < len(row):
            status = str(row[col_indices["Status"]]).strip()
            record["Status"] = status if status else STATUS_NOT_CONTACTED
        else:
            record["Status"] = STATUS_NOT_CONTACTED

        rows.append(record)

    if not rows:
        return pd.DataFrame(columns=REQUIRED_COLUMNS)

    df = pd.DataFrame(rows)
    return df


def add_recruiter(worksheet, name: str, email: str, company: str, role: str):
    """
    Add a new recruiter row to the Google Sheet.

    Appends the row at the bottom. The Apps Script in the Sheet
    can auto-sort by Company if you've set that up.
    """
    worksheet.append_row(
        [name, email, company, role, STATUS_NOT_CONTACTED],
        value_input_option="USER_ENTERED",
    )


def delete_recruiter(worksheet, row_index: int):
    """
    Delete a recruiter from the Google Sheet by row index.

    Args:
        row_index: 0-based index in the DataFrame → row_index + 2 in sheet
                   (because row 1 = headers, and gspread is 1-indexed)
    """
    sheet_row = row_index + 2  # +1 for header, +1 for 1-based indexing
    worksheet.delete_rows(sheet_row)


def update_status(worksheet, row_index: int, new_status: str):
    """
    Update the Status column for a specific recruiter in the Sheet.

    The Status column is E (column 5).
    """
    sheet_row = row_index + 2
    worksheet.update_cell(sheet_row, 5, new_status)


def batch_update_status(worksheet, row_indices: list[int], new_status: str):
    """
    Update Status for multiple recruiters at once (batch operation).

    Uses batch_update to minimise API calls — important when creating
    drafts for many recruiters at once.
    """
    cells = []
    for idx in row_indices:
        sheet_row = idx + 2
        cells.append(gspread.Cell(sheet_row, 5, new_status))

    if cells:
        worksheet.update_cells(cells)


def save_edits_to_sheet(worksheet, original_df: pd.DataFrame, edited_df: pd.DataFrame):
    """
    Compare original and edited DataFrames, then write any changes
    back to the Google Sheet.

    Only updates cells that actually changed — doesn't rewrite the
    entire sheet, which keeps API usage low.

    Returns:
        Number of cells updated.
    """
    # Column mapping: DataFrame column name → sheet column number (1-indexed)
    col_map = {"Name": 1, "Email": 2, "Company": 3, "Role": 4, "Status": 5}

    cells_to_update = []

    for idx in edited_df.index:
        if idx not in original_df.index:
            continue

        for col_name, sheet_col in col_map.items():
            old_val = str(original_df.at[idx, col_name])
            new_val = str(edited_df.at[idx, col_name])

            if old_val != new_val:
                sheet_row = idx + 2  # +1 header, +1 for 1-based
                cells_to_update.append(
                    gspread.Cell(sheet_row, sheet_col, new_val)
                )

    if cells_to_update:
        worksheet.update_cells(cells_to_update)

    return len(cells_to_update)


def get_history_sheet(spreadsheet):
    """
    Get or create the 'History' tab in the spreadsheet.

    The History tab logs every draft creation run with:
        - Date & Time
        - Drafts Created (count)
        - Recruiters (comma-separated names)
    """
    try:
        history_ws = spreadsheet.worksheet("History")
    except gspread.exceptions.WorksheetNotFound:
        # Create the History tab with headers
        history_ws = spreadsheet.add_worksheet(
            title="History", rows=1000, cols=3
        )
        history_ws.update(
            "A1:C1",
            [["Date & Time", "Drafts Created", "Recruiters"]],
        )

    return history_ws


def log_history(spreadsheet, num_drafts: int, recruiter_names: list[str]):
    """
    Log a draft creation run to the History tab.

    Args:
        spreadsheet:     The gspread Spreadsheet object.
        num_drafts:      Number of drafts successfully created.
        recruiter_names: List of recruiter names that were drafted.
    """
    history_ws = get_history_sheet(spreadsheet)

    timestamp = datetime.now().strftime("%B %d, %Y at %I:%M %p")
    names_str = ", ".join(recruiter_names)

    history_ws.append_row(
        [timestamp, num_drafts, names_str],
        value_input_option="USER_ENTERED",
    )


def load_history(spreadsheet) -> pd.DataFrame:
    """
    Load history data from the History tab.

    Returns:
        DataFrame with columns: Date & Time, Drafts Created, Recruiters
    """
    history_ws = get_history_sheet(spreadsheet)
    data = history_ws.get_all_records()

    if not data:
        return pd.DataFrame(columns=["Date & Time", "Drafts Created", "Recruiters"])

    return pd.DataFrame(data)


# ---------------------------------------------------------------------------
# EMAIL HELPERS
# ---------------------------------------------------------------------------


def validate_email(email: str) -> bool:
    """Quick regex check for valid email format."""
    pattern = r"^[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}$"
    return bool(re.match(pattern, email))


def build_mime_message(
    to_email: str,
    subject: str,
    body: str,
    resume_bytes: bytes,
    resume_filename: str,
) -> dict:
    """Build a Gmail-compatible MIME message with a PDF attachment."""
    msg = MIMEMultipart()
    msg["to"] = to_email
    msg["subject"] = subject

    msg.attach(MIMEText(body, "plain"))

    pdf_part = MIMEApplication(resume_bytes, _subtype="pdf")
    pdf_part.add_header(
        "Content-Disposition", "attachment", filename=resume_filename
    )
    msg.attach(pdf_part)

    raw = base64.urlsafe_b64encode(msg.as_bytes()).decode("utf-8")
    return {"raw": raw}


def create_gmail_draft(service, message_body: dict) -> str:
    """Create a single draft in Gmail. Returns the draft ID."""
    draft = (
        service.users()
        .drafts()
        .create(userId="me", body={"message": message_body})
        .execute()
    )
    return draft.get("id", "unknown")


# ---------------------------------------------------------------------------
# STREAMLIT UI
# ---------------------------------------------------------------------------


def main():
    """Main Streamlit app."""

    st.set_page_config(
        page_title="RecEmailed",
        page_icon="📧",
        layout="centered",
    )

    st.title("📧 RecEmailed")
    st.caption(
        "Manage your recruiter list and create Gmail drafts. "
        "Nothing is sent — you review and send each draft yourself."
    )

    # ==================================================================
    # SIDEBAR — Google Sheet Connection
    # ==================================================================
    with st.sidebar:
        st.header("🔗 Google Sheet")
        st.caption(
            "Paste your Google Sheet URL below. "
            "Make sure the sheet has columns: "
            "**Name, Email, Company, Role, Status**"
        )

        sheet_url = st.text_input(
            "Sheet URL",
            value=st.session_state.get("sheet_url", ""),
            placeholder="https://docs.google.com/spreadsheets/d/...",
            label_visibility="collapsed",
        )

        if sheet_url:
            st.session_state.sheet_url = sheet_url

        connect_btn = st.button(
            "🔌 Connect",
            use_container_width=True,
            disabled=not sheet_url,
        )

        # Check for credentials.json
        if not os.path.exists(CREDENTIALS_FILE):
            st.error(
                f"⚠️ `{CREDENTIALS_FILE}` not found. "
                "Download from Google Cloud Console."
            )

        st.divider()
        st.caption(
            "**Required APIs:**\n"
            "- Gmail API\n"
            "- Google Sheets API\n\n"
            "Enable both in Google Cloud Console."
        )

    # ==================================================================
    # CONNECT TO SHEET
    # ==================================================================
    if connect_btn and sheet_url:
        with st.spinner("🔐 Connecting..."):
            creds = get_credentials()
            if creds is None:
                st.error(f"❌ `{CREDENTIALS_FILE}` not found.")
                return

            try:
                worksheet, spreadsheet = get_gsheet(creds, sheet_url)
                st.session_state.creds = creds
                st.session_state.worksheet = worksheet
                st.session_state.spreadsheet = spreadsheet
                st.session_state.connected = True
                # Force a fresh load of data
                st.session_state.pop("df", None)
            except Exception as e:
                st.error(f"❌ Could not connect to sheet: {e}")
                st.session_state.connected = False

    # ==================================================================
    # MAIN CONTENT — only show if connected
    # ==================================================================
    if not st.session_state.get("connected"):
        st.info("👈 Paste your Google Sheet URL in the sidebar and click Connect.")
        return

    worksheet = st.session_state.worksheet
    spreadsheet = st.session_state.spreadsheet
    creds = st.session_state.creds

    # Load data (cache in session state, refresh with button)
    if "df" not in st.session_state or st.session_state.df is None:
        with st.spinner("Loading recruiter data..."):
            st.session_state.df = load_recruiters(worksheet)

    df = st.session_state.df

    # ------------------------------------------------------------------
    # SECTION 1: Manage Recruiters
    # ------------------------------------------------------------------
    st.header("1️⃣ Manage Recruiters")

    # --- Add New Recruiter ---
    with st.expander("➕ Add a New Recruiter", expanded=False):
        with st.form("add_recruiter_form", clear_on_submit=True):
            cols = st.columns(4)
            new_name = cols[0].text_input("Name", placeholder="Jane Smith")
            new_email = cols[1].text_input("Email", placeholder="jane@google.com")
            new_company = cols[2].text_input("Company", placeholder="Google")
            new_role = cols[3].text_input("Role", placeholder="SWE")

            submitted = st.form_submit_button(
                "➕ Add Recruiter", use_container_width=True
            )

            if submitted:
                # Validate inputs
                if not all([new_name.strip(), new_email.strip(),
                            new_company.strip(), new_role.strip()]):
                    st.error("All fields are required.")
                elif not validate_email(new_email.strip()):
                    st.error(f"Invalid email: {new_email}")
                else:
                    add_recruiter(
                        worksheet,
                        new_name.strip(),
                        new_email.strip(),
                        new_company.strip(),
                        new_role.strip(),
                    )
                    st.success(f"✅ Added {new_name.strip()}!")
                    # Refresh data
                    st.session_state.df = load_recruiters(worksheet)
                    st.rerun()

    # --- Search & Filter ---
    search_col, filter_col, refresh_col = st.columns([3, 2, 1])

    with search_col:
        search_query = st.text_input(
            "🔍 Search",
            placeholder="Search by name, email, or company...",
            label_visibility="collapsed",
        )

    with filter_col:
        status_filter = st.multiselect(
            "Filter by Status",
            options=ALL_STATUSES,
            default=ALL_STATUSES,
            label_visibility="collapsed",
        )

    with refresh_col:
        if st.button("🔄", help="Refresh data from Google Sheet"):
            st.session_state.df = load_recruiters(worksheet)
            st.rerun()

    # Apply search and filter
    display_df = df.copy()

    if search_query:
        query = search_query.lower()
        mask = (
            display_df["Name"].str.lower().str.contains(query, na=False)
            | display_df["Email"].str.lower().str.contains(query, na=False)
            | display_df["Company"].str.lower().str.contains(query, na=False)
            | display_df["Role"].str.lower().str.contains(query, na=False)
        )
        display_df = display_df[mask]

    if status_filter:
        display_df = display_df[display_df["Status"].isin(status_filter)]

    # --- Status Summary ---
    if not df.empty:
        status_counts = df["Status"].value_counts()
        metric_cols = st.columns(len(ALL_STATUSES) + 1)
        metric_cols[0].metric("Total", len(df))
        for i, status in enumerate(ALL_STATUSES):
            count = status_counts.get(status, 0)
            metric_cols[i + 1].metric(status, count)

    # --- Recruiter Table with Selection ---
    if display_df.empty:
        if search_query:
            st.warning(f"No recruiters found matching '{search_query}'.")
        else:
            st.info("No recruiters yet. Add one above! ☝️")
    else:
        # Add checkbox column — checked only for "Not Contacted"
        display_df = display_df.copy()
        display_df.insert(
            0, "Select",
            display_df["Status"].apply(lambda s: s == STATUS_NOT_CONTACTED),
        )

        st.caption(
            "📝 All fields are editable — click any cell to modify. "
            "Hit **Save Changes** below to sync edits to your Google Sheet."
        )

        edited_df = st.data_editor(
            display_df,
            use_container_width=True,
            hide_index=True,
            column_config={
                "Select": st.column_config.CheckboxColumn(
                    "✅", help="Check to create a draft", width="small",
                ),
                "Status": st.column_config.SelectboxColumn(
                    "Status", options=ALL_STATUSES, width="medium",
                ),
                "Name": st.column_config.TextColumn("Name", width="medium"),
                "Email": st.column_config.TextColumn("Email", width="medium"),
                "Company": st.column_config.TextColumn("Company", width="medium"),
                "Role": st.column_config.TextColumn("Role", width="medium"),
            },
            num_rows="fixed",
        )

        # Save edited state
        st.session_state.edited_display = edited_df

        # --- Save Changes Button ---
        if st.button("💾 Save Changes to Sheet", use_container_width=True):
            # Compare edited_df with the display_df (before edits) to find changes
            # We skip the "Select" column — it's only for draft selection
            original_compare = display_df.drop(columns=["Select"])
            edited_compare = edited_df.drop(columns=["Select"])

            with st.spinner("Saving changes to Google Sheet..."):
                num_updated = save_edits_to_sheet(
                    worksheet, original_compare, edited_compare
                )

            if num_updated > 0:
                st.success(f"✅ Saved {num_updated} change(s) to Google Sheet!")
                # Refresh local data
                st.session_state.df = load_recruiters(worksheet)
                st.rerun()
            else:
                st.info("No changes detected.")

        # --- Delete a Recruiter ---
        with st.expander("🗑️ Delete a Recruiter"):
            # Search within the delete section
            delete_search = st.text_input(
                "Search recruiter to delete",
                placeholder="Type name, email, or company...",
                key="delete_search",
                label_visibility="collapsed",
            )

            # Filter options based on delete search
            delete_pool = display_df.drop(columns=["Select"])
            if delete_search:
                dq = delete_search.lower()
                delete_mask = (
                    delete_pool["Name"].str.lower().str.contains(dq, na=False)
                    | delete_pool["Email"].str.lower().str.contains(dq, na=False)
                    | delete_pool["Company"].str.lower().str.contains(dq, na=False)
                )
                delete_pool = delete_pool[delete_mask]

            recruiter_options = [
                f"{row['Name']} — {row['Email']} ({row['Company']})"
                for _, row in delete_pool.iterrows()
            ]

            if recruiter_options:
                selected_to_delete = st.selectbox(
                    "Pick one",
                    options=recruiter_options,
                    label_visibility="collapsed",
                )

                if st.button("🗑️ Delete", type="secondary"):
                    # Find the original index
                    del_idx = recruiter_options.index(selected_to_delete)
                    original_idx = delete_pool.index[del_idx]

                    name = df.at[original_idx, "Name"]
                    delete_recruiter(worksheet, original_idx)
                    st.success(f"🗑️ Deleted {name}.")
                    st.session_state.df = load_recruiters(worksheet)
                    st.rerun()
            else:
                st.info("No recruiters match your search.")

    st.divider()

    # ------------------------------------------------------------------
    # SECTION 2: Upload Resume
    # ------------------------------------------------------------------
    st.header("2️⃣ Resume")
    resume_file = st.file_uploader(
        "📎 Upload your resume PDF",
        type=["pdf"],
        help="Attached to every draft email.",
    )

    st.divider()

    # ------------------------------------------------------------------
    # SECTION 3: Email Template
    # ------------------------------------------------------------------
    st.header("3️⃣ Email Template")
    st.caption("Placeholders: **{Name}**, **{Company}**, **{Role}**")

    subject_template = st.text_input(
        "Subject Line", value=DEFAULT_SUBJECT,
    )

    body_template = st.text_area(
        "Email Body", value=DEFAULT_TEMPLATE, height=250,
    )

    # Preview
    if (
        st.session_state.get("edited_display") is not None
        and not st.session_state.edited_display.empty
    ):
        selected_rows = st.session_state.edited_display[
            st.session_state.edited_display["Select"]
        ]
        if not selected_rows.empty:
            first = selected_rows.iloc[0]
            with st.expander("👁️ Preview email for first selected recruiter"):
                try:
                    st.markdown(f"**To:** {first['Email']}")
                    st.markdown(
                        f"**Subject:** {subject_template.format(Name=first['Name'], Company=first['Company'], Role=first['Role'])}"
                    )
                    st.divider()
                    st.text(
                        body_template.format(
                            Name=first["Name"],
                            Company=first["Company"],
                            Role=first["Role"],
                        )
                    )
                except KeyError as e:
                    st.error(f"Invalid placeholder: {e}")

    st.divider()

    # ------------------------------------------------------------------
    # SECTION 4: Create Drafts
    # ------------------------------------------------------------------
    st.header("4️⃣ Create Drafts")

    # Check prerequisites
    ready = True
    edited = st.session_state.get("edited_display")

    if edited is None or edited.empty:
        st.info("No recruiter data loaded.")
        ready = False
    elif edited["Select"].sum() == 0:
        st.warning("☝️ Select at least one recruiter.")
        ready = False

    if not resume_file:
        st.info("⬆️ Upload your resume PDF first.")
        ready = False

    if not os.path.exists(CREDENTIALS_FILE):
        st.warning(f"⚠️ `{CREDENTIALS_FILE}` not found.")
        ready = False

    if ready:
        selected_df = edited[edited["Select"]].copy()
        st.info(f"**{len(selected_df)}** recruiter(s) selected for drafts.")

        if st.button(
            "🚀 Create Drafts in Gmail",
            type="primary",
            use_container_width=True,
        ):
            resume_bytes = resume_file.getvalue()
            resume_filename = resume_file.name

            # Get Gmail service
            with st.spinner("🔐 Authenticating with Gmail..."):
                gmail_service = get_gmail_service(creds)

            st.success("✅ Authenticated!")

            # Create drafts
            progress = st.progress(0, text="Creating drafts...")
            results = []
            drafted_indices = []

            for i, (idx, row) in enumerate(selected_df.iterrows()):
                name = row["Name"]
                email = row["Email"]
                company = row["Company"]
                role = row["Role"]

                progress.progress(
                    i / len(selected_df),
                    text=f"Creating draft for {name} ({email})...",
                )

                try:
                    subject = subject_template.format(
                        Name=name, Company=company, Role=role
                    )
                    body = body_template.format(
                        Name=name, Company=company, Role=role
                    )

                    message_body = build_mime_message(
                        to_email=email,
                        subject=subject,
                        body=body,
                        resume_bytes=resume_bytes,
                        resume_filename=resume_filename,
                    )

                    draft_id = create_gmail_draft(gmail_service, message_body)
                    drafted_indices.append(idx)

                    results.append({
                        "Recruiter": name,
                        "Email": email,
                        "Result": "✅ Created",
                        "Draft ID": draft_id,
                    })

                except Exception as e:
                    results.append({
                        "Recruiter": name,
                        "Email": email,
                        "Result": f"❌ {e}",
                        "Draft ID": "—",
                    })

            progress.progress(1.0, text="Done!")

            # Update statuses in Google Sheet (batch)
            if drafted_indices:
                with st.spinner("📝 Updating statuses in Google Sheet..."):
                    batch_update_status(
                        worksheet, drafted_indices, STATUS_DRAFTED
                    )

                # Log to history
                drafted_names = [
                    r["Recruiter"] for r in results
                    if r["Result"].startswith("✅")
                ]
                with st.spinner("📜 Logging to history..."):
                    log_history(spreadsheet, len(drafted_names), drafted_names)

                # Refresh local data
                st.session_state.df = load_recruiters(worksheet)

            # Show results
            st.divider()
            st.subheader("📊 Results")
            st.dataframe(results, use_container_width=True, hide_index=True)

            successes = sum(1 for r in results if r["Result"].startswith("✅"))
            failures = len(results) - successes

            if successes > 0:
                st.success(
                    f"🎉 {successes} draft(s) created! "
                    "Open **Gmail → Drafts** to review and send."
                )
            if failures > 0:
                st.error(f"⚠️ {failures} failed. Check the table above.")

    # ------------------------------------------------------------------
    # SECTION 5: History
    # ------------------------------------------------------------------
    st.divider()
    st.header("📜 History")
    st.caption("Log of all draft creation runs.")

    try:
        history_df = load_history(spreadsheet)

        if history_df.empty:
            st.info("No history yet. Create some drafts and it'll show up here!")
        else:
            # Show most recent runs first
            history_df = history_df.iloc[::-1].reset_index(drop=True)

            # Summary stats
            total_runs = len(history_df)
            total_drafted = history_df["Drafts Created"].astype(int).sum()
            h_cols = st.columns(2)
            h_cols[0].metric("Total Runs", total_runs)
            h_cols[1].metric("Total Drafts Ever Created", total_drafted)

            # History table
            st.dataframe(
                history_df,
                use_container_width=True,
                hide_index=True,
                column_config={
                    "Date & Time": st.column_config.TextColumn(
                        "📅 Date & Time", width="medium"
                    ),
                    "Drafts Created": st.column_config.NumberColumn(
                        "📧 Drafts", width="small"
                    ),
                    "Recruiters": st.column_config.TextColumn(
                        "👥 Recruiters", width="large"
                    ),
                },
            )
    except Exception:
        st.info("History will appear after your first draft creation run.")

    # ------------------------------------------------------------------
    # Footer
    # ------------------------------------------------------------------
    st.divider()
    st.caption(
        "RecEmailed · Your data lives in Google Sheets · "
        "Drafts saved in Gmail — nothing sent automatically."
    )


if __name__ == "__main__":
    main()
