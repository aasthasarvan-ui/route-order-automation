from datetime import datetime
import email
from email.header import decode_header
import imaplib
import os
import openpyxl
import pandas as pd
import streamlit as st

# Page Configuration with modern layout and branding
st.set_page_config(
    page_title="Nestle DR Downloader & Combiner",
    page_icon="📥",
    layout="centered",
)

# Custom CSS for polished styling, rounded cards, clean typography, and professional look
st.markdown("""
    <style>
    .main {
        background-color: #f4f6f9;
    }
    .stButton>button {
        width: 100%;
        border-radius: 8px;
        font-weight: 600;
        height: 3.2em;
        background: linear-gradient(135deg, #FF4B4B 0%, #FF6B6B 100%);
        color: white;
        border: none;
        box-shadow: 0 4px 12px rgba(255, 75, 75, 0.3);
        transition: all 0.3s ease;
    }
    .stButton>button:hover {
        transform: translateY(-2px);
        box-shadow: 0 6px 16px rgba(255, 75, 75, 0.4);
    }
    .metric-container {
        background-color: #ffffff;
        padding: 20px;
        border-radius: 12px;
        box-shadow: 0 4px 6px rgba(0,0,0,0.02);
        border: 1px solid #eef2f6;
        text-align: center;
        margin-bottom: 20px;
    }
    .card {
        background: #ffffff;
        padding: 24px;
        border-radius: 12px;
        box-shadow: 0 4px 12px rgba(0,0,0,0.03);
        border: 1px solid #eef2f6;
        margin-bottom: 20px;
    }
    </style>
""", unsafe_allow_html=True)

# App Header Section
st.title("📥 Nestle Bot DR Downloader & Merger")
st.markdown("""
    <p style='color: #4a5568; font-size: 16px; font-weight: 400; line-height: 1.5;'>
    High-speed enterprise automation tool to fetch emails, extract attachments starting with <b>'OutputBlankDR'</b>, 
    and consolidate data seamlessly into the master Excel template starting from row 6.
    </p>
""", unsafe_allow_html=True)
st.markdown("<hr style='margin: 20px 0; border: none; border-top: 1px solid #e2e8f0;'>", unsafe_allow_html=True)

# Load email password securely from Streamlit secrets if available
default_password = ""
try:
    if hasattr(st, "secrets") and "email" in st.secrets and "app_password" in st.secrets["email"]:
        default_password = st.secrets["email"]["app_password"]
except Exception:
    pass

# Sidebar Settings & Credentials
with st.sidebar:
    st.header("⚙️ Configuration & Auth")
    email_user = st.text_input("Email Address", value="harmesh.kumar@aseesfeed.com")
    email_pass = st.text_input(
        "Email App Password",
        value=default_password,
        type="password",
        help="Enter your Gmail App Password or store it securely in Streamlit Secrets."
    )
    imap_server = st.selectbox(
        "IMAP Server",
        ["imap.gmail.com", "imap.mail.yahoo.com", "imap.office365.com"]
    )
    mailbox_name = st.text_input(
        "Mailbox / Folder Name",
        value="BOT",
        help="Folder name where bot emails are received (e.g., BOT)"
    )
    
    st.markdown("---")
    st.markdown("### 📊 Quick Guide")
    st.markdown("""
    1. Select your target date.
    2. Optional subject filter.
    3. Click **Process & Combine**.
    4. Download the compiled Excel report.
    """)

# Main Search Parameters Section
st.subheader("🔍 Search & Filter Parameters")
col1, col2 = st.columns(2)

with col1:
    selected_date = st.date_input("Select Target Date", value=datetime.today())

with col2:
    subject_filter = st.text_input(
        "Subject Filter (Optional)",
        value="Nestle Bot Partailly Completed & Missing DR Codes",
        help="Leave blank to process all emails for the selected date."
    )

TEMPLATE_FILENAME = "Output.xlsx"
st.markdown("<hr style='margin: 20px 0; border: none; border-top: 1px solid #e2e8f0;'>", unsafe_allow_html=True)

# Action Button & Execution Engine
if st.button("🚀 Process & Combine Files (High Speed)", type="primary"):
    if not email_pass:
        st.warning("⚠️ Please provide your Email App Password in the sidebar.")
    elif not os.path.exists(TEMPLATE_FILENAME):
        st.error(f"❌ Template file '{TEMPLATE_FILENAME}' not found in the repository root!")
    else:
        download_folder = "temp_downloads"
        os.makedirs(download_folder, exist_ok=True)

        formatted_date = selected_date.strftime("%d-%m-%Y")
        imap_date = selected_date.strftime("%d-%b-%Y")

        # Initialize progress bar for visual feedback
        progress_bar = st.progress(0)
        status_text = st.empty()

        try:
            status_text.text(f"Connecting to {imap_server} and accessing '{mailbox_name}'...")
            progress_bar.progress(15)

            mail = imaplib.IMAP4_SSL(imap_server)
            mail.login(email_user, email_pass)

            rv, data = mail.select(mailbox_name)
            if rv != "OK":
                st.error(f"Could not select folder '{mailbox_name}'. Please verify the folder name.")
                mail.logout()
                progress_bar.empty()
            else:
                status_text.text(f"Searching emails for date: {formatted_date}...")
                progress_bar.progress(30)
                
                search_query = f'(ON "{imap_date}")'
                status, messages = mail.search(None, search_query)

                if status != "OK" or not messages[0]:
                    st.warning(f"No emails found in '{mailbox_name}' folder for date ({formatted_date}).")
                    mail.logout()
                    progress_bar.empty()
                else:
                    msg_nums = messages[0].split()
                    matching_nums = []

                    status_text.text("Filtering emails by subject criteria...")
                    progress_bar.progress(50)

                    if subject_filter.strip():
                        for num in msg_nums:
                            try:
                                res, h_data = mail.fetch(num, "(BODY.PEEK[HEADER.FIELDS (Subject)])")
                                for response_part in h_data:
                                    if isinstance(response_part, tuple):
                                        msg_header = email.message_from_bytes(response_part[1])
                                        raw_subject = msg_header["Subject"]
                                        if raw_subject:
                                            decoded_subj_list = decode_header(raw_subject)
                                            subject_str = ""
                                            for sub, enc in decoded_subj_list:
                                                if isinstance(sub, bytes):
                                                    subject_str += sub.decode(enc or "utf-8", errors="ignore")
                                                else:
                                                    subject_str += str(sub)

                                            if subject_filter.strip().lower() in subject_str.lower():
                                                matching_nums.append(num)
                            except Exception:
                                continue
                    else:
                        matching_nums = msg_nums

                    if not matching_nums:
                        st.warning("Found emails for this date, but none matched the subject criteria.")
                        mail.logout()
                        progress_bar.empty()
                    else:
                        status_text.text(f"Found {len(matching_nums)} matching email(s). Extracting attachments...")
                        progress_bar.progress(70)
                        
                        downloaded_files = []
                        for num in matching_nums:
                            res, msg_data = mail.fetch(num, "(RFC822)")
                            for response_part in msg_data:
                                if isinstance(response_part, tuple):
                                    msg = email.message_from_bytes(response_part[1])
                                    for part in msg.walk():
                                        if part.get_content_maintype() == "multipart":
                                            continue
                                        if part.get("Content-Disposition") is None:
                                            continue

                                        filename = part.get_filename()
                                        if filename:
                                            decoded_header_val = decode_header(filename)
                                            filename = decoded_header_val[0][0]
                                            if isinstance(filename, bytes):
                                                filename = filename.decode(
                                                    decoded_header_val[0][1] or "utf-8", errors="ignore"
                                                )

                                            if filename.startswith("OutputBlankDR") and filename.endswith(".xlsx"):
                                                filepath = os.path.join(download_folder, filename)
                                                with open(filepath, "wb") as f:
                                                    f.write(part.get_payload(decode=True))
                                                downloaded_files.append(filepath)

                        mail.logout()
                        progress_bar.progress(85)

                        if downloaded_files:
                            status_text.text("Merging data into template starting from row 6...")
                            
                            output_filename = f"Master_OutputBlankDR_{formatted_date.replace('-', '_')}.xlsx"
                            output_path = os.path.join(download_folder, output_filename)

                            wb = openpyxl.load_workbook(TEMPLATE_FILENAME)

                            if "Order Data" in wb.sheetnames:
                                ws = wb["Order Data"]
                                current_row = 6

                                for file in downloaded_files:
                                    try:
                                        df_source = pd.read_excel(file, sheet_name="Order Data", header=None)
                                        data_to_paste = df_source.iloc[5:].values.tolist()

                                        for row_data in data_to_paste:
                                            for col_idx, val in enumerate(row_data, start=1):
                                                ws.cell(row=current_row, column=col_idx, value=val)
                                            current_row += 1
                                    except Exception as e:
                                        st.error(f"Error reading file {file}: {e}")

                                wb.save(output_path)
                                progress_bar.progress(100)
                                status_text.text("Processing complete!")

                                st.success(f"🎉 Successfully processed {len(downloaded_files)} attachment(s) and combined data into master file!")
                                
                                # Display metrics in clean cards
                                m1, m2 = st.columns(2)
                                with m1:
                                    st.metric(label="📁 Files Consolidated", value=len(downloaded_files))
                                with m2:
                                    st.metric(label="📊 Total Rows Added", value=current_row - 6)

                                st.markdown("---")
                                st.info("📥 Click the button below to download your master Excel report:")
                                with open(output_path, "rb") as f:
                                    st.download_button(
                                        label="📥 Download Master Excel File",
                                        data=f,
                                        file_name=output_filename,
                                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                    )
                            else:
                                st.error("Sheet 'Order Data' was not found in the template file.")
                        else:
                            progress_bar.empty()
                            st.warning("Matching emails were found, but none contained attachments starting with 'OutputBlankDR'.")

        except Exception as e:
            progress_bar.empty()
            st.error(f"An unexpected error occurred: {e}")
