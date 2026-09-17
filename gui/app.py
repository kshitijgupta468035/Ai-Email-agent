import sys
import os
from io import BytesIO

import pandas as pd
import streamlit as st


# =========================================================
# PROJECT PATH
# =========================================================

PROJECT_ROOT = os.path.dirname(
    os.path.dirname(
        os.path.abspath(__file__)
    )
)

if PROJECT_ROOT not in sys.path:
    sys.path.insert(0, PROJECT_ROOT)


# =========================================================
# TOOLS
# =========================================================

from tools.excel_tool import (
    read_tasks,
    validate_task_columns,
    validate_task_data,
    prepare_email_tracking_columns,
    get_pending_tasks,
)

from tools.gmail_tool import (
    authenticate_gmail,
    send_email,
)

from tools.gemini_tool import (
    create_gemini_client,
    generate_text,
)


# =========================================================
# FILE PATHS
# =========================================================

INPUT_FOLDER = os.path.join(PROJECT_ROOT, "input")
OUTPUT_FILE = os.path.join(
    PROJECT_ROOT,
    "output",
    "output_tasks.xlsx",
)


# =========================================================
# PAGE CONFIGURATION
# =========================================================

st.set_page_config(
    page_title="AI Email Agent",
    page_icon="📧",
    layout="wide",
    initial_sidebar_state="expanded",
)


# =========================================================
# UI STYLING
# =========================================================

st.markdown(
    """
    <style>

    /* ---------- Main page ---------- */

    .block-container {
        padding-top: 2rem;
        padding-bottom: 3rem;
        max-width: 1400px;
    }

    .app-title {
        font-size: 2.35rem;
        font-weight: 750;
        line-height: 1.15;
        margin: 0;
    }

    .app-subtitle {
        color: #6b7280;
        font-size: 0.98rem;
        margin-top: 0.35rem;
        margin-bottom: 1.8rem;
    }

    .section-title {
        font-size: 1.45rem;
        font-weight: 700;
        margin-top: 0.4rem;
        margin-bottom: 0.25rem;
    }

    .section-description {
        color: #6b7280;
        margin-bottom: 1.2rem;
    }

    /* ---------- Sidebar ---------- */

    [data-testid="stSidebar"] {
        border-right: 1px solid rgba(128, 128, 128, 0.18);
    }

    .sidebar-title {
        font-size: 1.3rem;
        font-weight: 750;
        margin-bottom: 0.1rem;
    }

    .sidebar-subtitle {
        color: #6b7280;
        font-size: 0.82rem;
        margin-bottom: 1rem;
    }

    .status-card {
        border: 1px solid rgba(128, 128, 128, 0.20);
        border-radius: 10px;
        padding: 0.75rem 0.8rem;
        margin-top: 0.8rem;
        margin-bottom: 0.8rem;
    }

    .status-label {
        color: #6b7280;
        font-size: 0.78rem;
        margin-bottom: 0.15rem;
    }

    .status-value {
        font-size: 0.9rem;
        font-weight: 650;
    }

    /* ---------- Metric cards ---------- */

    [data-testid="stMetric"] {
        border: 1px solid rgba(128, 128, 128, 0.18);
        border-radius: 12px;
        padding: 0.9rem 1rem;
        background: rgba(128, 128, 128, 0.035);
    }

    /* ---------- Workflow cards ---------- */

    .workflow-card {
        border: 1px solid rgba(128, 128, 128, 0.18);
        border-radius: 12px;
        padding: 1rem 1.1rem;
        min-height: 125px;
        margin-bottom: 0.8rem;
    }

    .workflow-number {
        color: #6b7280;
        font-size: 0.78rem;
        font-weight: 650;
        text-transform: uppercase;
        letter-spacing: 0.04em;
    }

    .workflow-name {
        font-size: 1.05rem;
        font-weight: 700;
        margin-top: 0.25rem;
    }

    .workflow-text {
        color: #6b7280;
        font-size: 0.87rem;
        margin-top: 0.25rem;
    }

    /* ---------- Email preview ---------- */

    .email-card {
        border: 1px solid rgba(128, 128, 128, 0.22);
        border-radius: 12px;
        padding: 1rem 1.1rem;
        margin-bottom: 1rem;
    }

    .email-recipient {
        font-size: 1.05rem;
        font-weight: 700;
    }

    .email-meta {
        color: #6b7280;
        font-size: 0.86rem;
        margin-top: 0.2rem;
    }

    /* ---------- Small text ---------- */

    .muted {
        color: #6b7280;
        font-size: 0.88rem;
    }

    </style>
    """,
    unsafe_allow_html=True,
)


# =========================================================
# SESSION STATE
# =========================================================

DEFAULT_STATE = {
    "tasks_df": None,
    "gmail": None,
    "gemini_client": None,
    "gemini_api_key": "",
    "email_previews": {},
    "approved_emails": set(),
    "send_emails": True,
    "max_emails": 50,
    "last_result": None,
}

for key, value in DEFAULT_STATE.items():
    if key not in st.session_state:
        st.session_state[key] = value


# =========================================================
# HELPERS
# =========================================================

def create_task_key(row):
    """Create a stable identifier for restoring sent status."""
    return (
        str(row["Name"]).strip().lower(),
        str(row["Email"]).strip().lower(),
        str(row["Task"]).strip().lower(),
        str(row["Due_Date"]).strip(),
    )


def restore_previous_state(df):
    """
    Restore only previously Sent email states from the
    application's output file.
    """
    if not os.path.exists(OUTPUT_FILE):
        return df, False

    try:
        previous_state = read_tasks(OUTPUT_FILE)
        previous_state = prepare_email_tracking_columns(
            previous_state
        )

        previous_status = {}

        for _, row in previous_state.iterrows():
            previous_status[create_task_key(row)] = {
                "Email_Status": row["Email_Status"],
                "Last_Email_Sent": row["Last_Email_Sent"],
            }

        restored = False

        for index, row in df.iterrows():
            key = create_task_key(row)

            if key not in previous_status:
                continue

            old_status = previous_status[key]

            if (
                str(old_status["Email_Status"])
                .strip()
                .lower()
                == "sent"
            ):
                df.at[index, "Email_Status"] = "Sent"
                df.at[index, "Last_Email_Sent"] = (
                    old_status["Last_Email_Sent"]
                )
                restored = True

        return df, restored

    except Exception as e:
        st.warning(
            f"Previous output could not be restored: {e}"
        )
        return df, False


def load_uploaded_excel(uploaded_file):
    """Read, validate, prepare, and restore an uploaded Excel file."""
    df = read_tasks(uploaded_file)
    validate_task_columns(df)
    validate_task_data(df)
    df = prepare_email_tracking_columns(df)
    df, restored = restore_previous_state(df)
    return df, restored


def get_task_counts(df):
    """Return total, pending/not-sent, and sent counts."""
    pending_tasks = get_pending_tasks(df)

    total = len(df)

    sent = len(
        df[
            df["Email_Status"]
            .astype(str)
            .str.strip()
            .str.lower()
            .eq("sent")
        ]
    )

    pending_not_sent = len(
        pending_tasks[
            ~pending_tasks["Email_Status"]
            .astype(str)
            .str.strip()
            .str.lower()
            .eq("sent")
        ]
    )

    return total, pending_not_sent, sent


def get_unsent_pending_tasks(df):
    """Return only Pending tasks that have not already been sent."""
    pending = get_pending_tasks(df)

    return pending[
        ~pending["Email_Status"]
        .astype(str)
        .str.strip()
        .str.lower()
        .eq("sent")
    ]


def generate_preview_for_row(index, row):
    """Generate one email preview with Gemini."""
    prompt = f"""
Write a professional and friendly reminder email.

Recipient name: {row["Name"]}
Task: {row["Task"]}
Due date: {row["Due_Date"]}

Keep the email concise.

Do not make up any information.

Include a subject line.
"""

    body = generate_text(
        st.session_state.gemini_client,
        prompt,
    )

    return {
        "name": row["Name"],
        "email": row["Email"],
        "task": row["Task"],
        "due_date": row["Due_Date"],
        "body": body,
    }


# =========================================================
# SIDEBAR
# =========================================================

with st.sidebar:

    st.markdown(
        '<div class="sidebar-title">📧 AI Email Agent</div>',
        unsafe_allow_html=True,
    )

    st.markdown(
        '<div class="sidebar-subtitle">'
        'Smart task reminders powered by AI'
        '</div>',
        unsafe_allow_html=True,
    )

    st.divider()

    page = st.radio(
        "Go to",
        [
            "📊 Dashboard",
            "📋 Tasks",
            "🤖 Gemini",
            "📧 Gmail",
            "✉️ Email Preview",
            "⚙️ Settings",
        ],
        label_visibility="visible",
    )

    st.divider()

    gmail_status = (
        "🟢 Connected"
        if st.session_state.gmail is not None
        else "🔴 Not connected"
    )

    gemini_status = (
        "🟢 Ready"
        if st.session_state.gemini_client is not None
        else "🔴 Not connected"
    )

    st.markdown(
        f"""
        <div class="status-card">
            <div class="status-label">Gmail</div>
            <div class="status-value">{gmail_status}</div>
        </div>

        <div class="status-card">
            <div class="status-label">Gemini</div>
            <div class="status-value">{gemini_status}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.caption(
        "Sending is "
        + ("ON" if st.session_state.send_emails else "OFF")
    )


# =========================================================
# MAIN HEADER
# =========================================================

st.markdown(
    '<div class="app-title">📧 AI Email Agent</div>',
    unsafe_allow_html=True,
)

st.markdown(
    '<div class="app-subtitle">'
    'Automate task reminders with Excel, Gemini and Gmail'
    '</div>',
    unsafe_allow_html=True,
)


# =========================================================
# DASHBOARD
# =========================================================

if page == "📊 Dashboard":

    st.markdown(
        '<div class="section-title">📊 Dashboard</div>',
        unsafe_allow_html=True,
    )

    st.markdown(
        '<div class="section-description">'
        'Manage your reminder workflow from one place.'
        '</div>',
        unsafe_allow_html=True,
    )

    if st.session_state.tasks_df is not None:
        total, pending, sent = get_task_counts(
            st.session_state.tasks_df
        )
    else:
        total, pending, sent = 0, 0, 0

    col1, col2, col3 = st.columns(3)

    with col1:
        st.metric("Total Tasks", total)

    with col2:
        st.metric("Pending / Not Sent", pending)

    with col3:
        st.metric("Emails Sent", sent)

    st.subheader("Workflow")

    c1, c2, c3 = st.columns(3)

    with c1:
        st.markdown(
            """
            <div class="workflow-card">
                <div class="workflow-number">Step 01</div>
                <div class="workflow-name">📂 Upload Tasks</div>
                <div class="workflow-text">
                    Upload and validate your Excel task file.
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    with c2:
        st.markdown(
            """
            <div class="workflow-card">
                <div class="workflow-number">Step 02</div>
                <div class="workflow-name">🤖 Generate</div>
                <div class="workflow-text">
                    Use Gemini to create reminder email previews.
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    with c3:
        st.markdown(
            """
            <div class="workflow-card">
                <div class="workflow-number">Step 03</div>
                <div class="workflow-name">🚀 Approve & Send</div>
                <div class="workflow-text">
                    Review approved emails and send them through Gmail.
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    st.subheader("Upload Task File")

    uploaded_file = st.file_uploader(
        "Choose an Excel file",
        type=["xlsx"],
        help=(
            "Required columns: Name, Email, Task, Due_Date, Status"
        ),
    )

    if uploaded_file is not None:

        try:
            with st.spinner("Reading and validating Excel..."):
                df, restored = load_uploaded_excel(
                    uploaded_file
                )

            st.session_state.tasks_df = df
            st.session_state.email_previews = {}
            st.session_state.approved_emails = set()

            if restored:
                st.success(
                    "Excel loaded. Previous sent-email history "
                    "has been restored."
                )
            else:
                st.success(
                    "Excel loaded and validated successfully."
                )

        except ValueError as e:
            st.error(f"❌ Validation error: {e}")

        except FileNotFoundError as e:
            st.error(f"❌ File error: {e}")

        except PermissionError as e:
            st.error(f"❌ Permission error: {e}")

        except Exception as e:
            st.error(f"❌ Something went wrong: {e}")

    elif st.session_state.tasks_df is None:
        st.info(
            "Upload an Excel file to start the workflow."
        )


# =========================================================
# TASKS
# =========================================================

elif page == "📋 Tasks":

    st.markdown(
        '<div class="section-title">📋 Tasks</div>',
        unsafe_allow_html=True,
    )

    st.markdown(
        '<div class="section-description">'
        'View the currently loaded tasks and their email status.'
        '</div>',
        unsafe_allow_html=True,
    )

    df = st.session_state.tasks_df

    if df is None:
        st.info(
            "No task file is loaded. Go to Dashboard and upload Excel."
        )
    else:
        total, pending, sent = get_task_counts(df)

        col1, col2, col3 = st.columns(3)

        with col1:
            st.metric("Total", total)

        with col2:
            st.metric("Pending / Not Sent", pending)

        with col3:
            st.metric("Sent", sent)

        st.dataframe(
            df,
            use_container_width=True,
            hide_index=True,
        )


# =========================================================
# GEMINI
# =========================================================

elif page == "🤖 Gemini":

    st.markdown(
        '<div class="section-title">🤖 Gemini</div>',
        unsafe_allow_html=True,
    )

    st.markdown(
        '<div class="section-description">'
        'Configure and test the Gemini model used to generate emails.'
        '</div>',
        unsafe_allow_html=True,
    )

    st.text_input(
        "Gemini API Key",
        type="password",
        placeholder="Enter your Gemini API key",
        key="gemini_api_key_input",
    )

    gemini_api_key = st.session_state.gemini_api_key_input

    if st.button (
        "🔌 Test Gemini Connection",
        use_container_width=True,
    ):

        try:
            if st.session_state.gemini_client is None:
                
                if not gemini_api_key.strip():
                    raise ValueError(
                        "Please enter your Gemini API key."
                    )

                with st.spinner("Connecting to Gemini..."):
                    client = create_gemini_client(
                        gemini_api_key
                    )

                    response = generate_text(
                        client,
                        "Reply with exactly: Gemini connection successful.",
                    )

                st.session_state.gemini_client = client
                st.session_state.gemini_api_key = gemini_api_key

                st.success("🟢 Gemini connected successfully.")
                st.caption("Test response:")
                st.code(response)
            
        except ValueError as e:
            st.error(f"❌ {e}")

        except Exception as e:
            st.session_state.gemini_client = None
            st.error(
                f"❌ Gemini connection failed: {e}"
            )

        if st.session_state.gemini_client is not None:
            st.success("🟢 Gemini is ready for email generation.")


# =========================================================
# GMAIL
# =========================================================

elif page == "📧 Gmail":

    st.markdown(
        '<div class="section-title">📧 Gmail</div>',
        unsafe_allow_html=True,
    )

    st.markdown(
        '<div class="section-description">'
        'Connect your Gmail account before sending approved emails.'
        '</div>',
        unsafe_allow_html=True,
    )

    

    if st.button(
        "🔌 Connect Gmail",
        use_container_width=True,
    ):

        try:
            if st.session_state.gmail is None:

                with st.spinner("Connecting to Gmail..."):
                    gmail = authenticate_gmail()

                st.session_state.gmail = gmail

        except TimeoutError as e:
            st.session_state.gmail = None
            st.error(f"⏱️ {e}")

        except Exception as e:
            st.session_state.gmail = None
            st.error(
                f"❌ Gmail connection failed: {e}"
            )

    if st.session_state.gmail is None:
        st.warning(
                    "🔴 Gmail is not connected."
                )
    else:
        st.success("🟢 Gmail connected successfully.")


    st.info(
        "Gmail authentication opens Google's authorization page "
        "when authorization is required."
    )


# =========================================================
# EMAIL PREVIEW
# =========================================================

elif page == "✉️ Email Preview":

    st.markdown(
        '<div class="section-title">✉️ Email Preview</div>',
        unsafe_allow_html=True,
    )

    st.markdown(
        '<div class="section-description">'
        'Generate, review and approve reminder emails before sending.'
        '</div>',
        unsafe_allow_html=True,
    )

    df = st.session_state.tasks_df

    if df is None:
        st.info(
            "Please upload and validate an Excel file from Dashboard."
        )

    elif st.session_state.gemini_client is None:
        st.warning(
            "Please connect Gemini from the Gemini page first."
        )

    else:

        unsent = get_unsent_pending_tasks(df)

        col1, col2 = st.columns([3, 1])

        with col1:
            st.write(
                f"Unsent pending tasks: **{len(unsent)}**"
            )

        with col2:
            st.write(
                f"Approved: **{len(st.session_state.approved_emails)}**"
            )

        if st.button(
            "✨ Generate Email Previews",
            use_container_width=True,
        ):

            try:
                if unsent.empty:
                    st.info(
                        "There are no unsent pending tasks."
                    )
                    
                    st.session_state.email_previews = {}
                    st.session_state.approved_emails = set()

                else:
                    if not st.session_state.email_previews:
                        limited_tasks = unsent.head(
                            st.session_state.max_emails
                        )

                        previews = {}
                        progress = st.progress(0)

                        for position, (index, row) in enumerate(
                            limited_tasks.iterrows(),
                            start=1,
                        ):
                            with st.spinner(
                                f"Generating email {position} "
                                f"of {len(limited_tasks)}..."
                            ):
                                previews[index] = (
                                    generate_preview_for_row(
                                        index,
                                        row,
                                    )
                                )

                            progress.progress(
                                position / len(limited_tasks)
                            )

                        st.session_state.email_previews = previews
                        st.session_state.approved_emails = set()

                        st.success(
                            f"Generated {len(previews)} email preview(s)."
                        )
                    else:
                        st.success(
                            "🟢 Email Previews are already genereated."
                        )

            except Exception as e:
                st.error(
                    f"❌ Failed to generate previews: {e}"
                )

        previews = st.session_state.email_previews

        if previews:
            st.subheader("Review Emails")

            for index, preview in previews.items():

                st.markdown(
                    f"""
                    <div class="email-card">
                        <div class="email-recipient">
                            {preview["name"]}
                        </div>
                        <div class="email-meta">
                            {preview["email"]}
                        </div>
                        <div class="email-meta">
                            Task: {preview["task"]}
                        </div>
                        <div class="email-meta">
                            Due Date: {preview["due_date"]}
                        </div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )

                st.write(
                    f"**Subject:** Reminder: {preview['task']}"
                )

                email_body = st.text_area(
                    "Email Body",
                    value=preview["body"],
                    height=220,
                    key=f"email_body_{index}",
                )

                approved = st.checkbox(
                    "Approve this email for sending",
                    key=f"approve_{index}",
                )

                if approved:
                    st.session_state.approved_emails.add(index)
                else:
                    st.session_state.approved_emails.discard(index)

                st.divider()

            approved_count = len(
                st.session_state.approved_emails
            )

            st.subheader("Send Approved Emails")

            if not st.session_state.send_emails:
                st.warning(
                    "Email sending is currently OFF in Settings."
                )

            st.write(
                f"Approved emails: **{approved_count}**"
            )

            if st.button(
                "🚀 Send Approved Emails",
                disabled=(
                    approved_count == 0
                    or not st.session_state.send_emails
                ),
                use_container_width=True,
            ):

                try:
                    if st.session_state.gmail is None:
                        raise ValueError(
                            "Please connect Gmail before sending."
                        )

                    if approved_count == 0:
                        raise ValueError(
                            "No emails have been approved."
                        )

                    sent_count = 0
                    failed_count = 0

                    for index in list(
                        st.session_state.approved_emails
                    ):

                        preview = previews[index]

                        current_status = (
                            str(
                                df.at[
                                    index,
                                    "Email_Status",
                                ]
                            )
                            .strip()
                            .lower()
                        )

                        if current_status == "sent":
                            st.warning(
                                f"Skipped {preview['email']} "
                                "because it was already sent."
                            )
                            continue

                        try:
                            email_body = st.session_state.get(
                                f"email_body_{index}",
                                preview["body"],
                            )

                            with st.spinner(
                                f"Sending email to "
                                f"{preview['email']}..."
                            ):
                                send_email(
                                    gmail=st.session_state.gmail,
                                    recipient=preview["email"],
                                    subject=(
                                        f"Reminder: "
                                        f"{preview['task']}"
                                    ),
                                    body=email_body,
                                )

                            df.at[
                                index,
                                "Email_Status",
                            ] = "Sent"

                            df.at[
                                index,
                                "Last_Email_Sent",
                            ] = (
                                pd.Timestamp.now().strftime(
                                    "%Y-%m-%d %H:%M:%S"
                                )
                            )

                            sent_count += 1

                        except Exception as e:
                            failed_count += 1
                            st.error(
                                f"Failed to send to "
                                f"{preview['email']}: {e}"
                            )

                    os.makedirs(
                        os.path.dirname(OUTPUT_FILE),
                        exist_ok=True,
                    )

                    df.to_excel(
                        OUTPUT_FILE,
                        index=False,
                    )

                    st.session_state.tasks_df = df
                    st.session_state.approved_emails = set()

                    if sent_count:
                        st.success(
                            f"✅ Successfully sent "
                            f"{sent_count} email(s)."
                        )

                    if failed_count:
                        st.warning(
                            f"⚠️ {failed_count} email(s) failed."
                        )

                    st.info(
                        "Updated task state has been saved."
                    )

                except ValueError as e:
                    st.error(f"❌ {e}")

                except Exception as e:
                    st.error(
                        f"❌ Sending process failed: {e}"
                    )


    # =========================================================
    # DOWNLOAD UPDATED EXCEL
    # =========================================================

    if st.session_state.tasks_df is not None:

        st.divider()

        st.subheader("📥 Download Updated Excel")

        output_buffer = BytesIO()

        with pd.ExcelWriter(
            output_buffer,
            engine="openpyxl",
        ) as writer:

            st.session_state.tasks_df.to_excel(
                writer,
                index=False,
                sheet_name="Tasks",
            )

        output_buffer.seek(0)

        st.download_button(
            label="📥 Download Updated Excel",
            data=output_buffer,
            file_name="output_tasks.xlsx",
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
            use_container_width=True,
        )

# =========================================================
# SETTINGS
# =========================================================

elif page == "⚙️ Settings":

    st.markdown(
        '<div class="section-title">⚙️ Settings</div>',
        unsafe_allow_html=True,
    )

    st.markdown(
        '<div class="section-description">'
        'Control how the email workflow behaves.'
        '</div>',
        unsafe_allow_html=True,
    )

    st.subheader("Email Sending")

    send_emails = st.toggle(
        "Enable email sending",
        value=st.session_state.send_emails,
        help=(
            "When OFF, emails can be generated and reviewed "
            "but cannot be sent."
        ),
    )

    st.session_state.send_emails = send_emails

    if send_emails:
        st.success(
            "🟢 Sending is enabled."
        )
    else:
        st.warning(
            "🟡 Sending is disabled. The Send button will remain disabled."
        )

    st.subheader("Generation Limit")

    max_emails = st.number_input(
        "Maximum emails to generate at one time",
        min_value=1,
        max_value=500,
        value=int(st.session_state.max_emails),
        step=1,
        help=(
            "Limits how many unsent pending tasks are processed "
            "when generating previews."
        ),
    )

    st.session_state.max_emails = int(max_emails)

    st.subheader("Application State")

    if st.session_state.tasks_df is None:
        st.caption("No Excel file loaded.")
    else:
        st.caption(
            f"Loaded tasks: {len(st.session_state.tasks_df)}"
        )

    st.caption(
        "Output file: "
        + OUTPUT_FILE
    )

    if st.button(
        "🧹 Clear Email Previews",
        use_container_width=True,
    ):
        st.session_state.email_previews = {}
        st.session_state.approved_emails = set()
        st.success(
            "Email previews and approvals were cleared."
        )

# =========================================================
# SETTINGS
# =========================================================

    st.subheader("📁 Email History")

    st.write(
        "The output file stores the email-sending history "
        "of your tasks."
    )

    if os.path.exists(OUTPUT_FILE):

        st.warning(
            "⚠️ An output file currently exists. "
            "Deleting it will remove the saved email history."
        )

        if st.button(
            "🗑️ Delete Output File",
            type="secondary"
        ):
            st.session_state.confirm_delete_output = True

    else:

        st.info(
            "No output file currently exists."
        )

    # -----------------------------------------------------
    # DELETE CONFIRMATION
    # -----------------------------------------------------

    if st.session_state.get(
        "confirm_delete_output",
        False
    ):

        st.error(
            "⚠️ Are you sure you want to delete "
            "the output file?"
        )

        st.write(
            "After deletion, previously sent tasks will "
            "no longer have their saved sending history. "
            "They can therefore be processed again."
        )

        col1, col2 = st.columns(2)

        with col1:

            if st.button(
                "❌ Cancel",
                use_container_width=True
            ):

                st.session_state.confirm_delete_output = False
                st.rerun()

        with col2:

            if st.button(
                "🗑️ Yes, Delete Output File",
                type="primary",
                use_container_width=True
            ):

                try:

                    if os.path.exists(OUTPUT_FILE):

                        os.remove(OUTPUT_FILE)

                        # Clear the confirmation state
                        st.session_state.confirm_delete_output = False

                        # Clear the current task state
                        if "tasks_df" in st.session_state:
                            st.session_state.tasks_df = None

                        # Clear generated email previews
                        st.session_state.email_previews = {}

                        # Clear approvals
                        st.session_state.approved_emails = set()

                        st.success(
                            "✅ Output file deleted successfully. "
                            "Email history has been reset."
                        )

                    else:

                        st.info(
                            "The output file does not exist."
                        )

                except PermissionError:

                    st.error(
                        "❌ Could not delete the output file. "
                        "Please make sure output_tasks.xlsx "
                        "is not open in Excel."
                    )

                except Exception as e:

                    st.error(
                        f"❌ Failed to delete output file: {e}"
                    )


