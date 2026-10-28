from datetime import datetime
import email
from email.header import decode_header
import imaplib
import os
import openpyxl
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Nestle DR Bot Automation Hub",
    page_icon="🤖",
    layout="wide",
)

st.title("🤖 Nestle Bot DR Automation Hub")
st.markdown(
    "Production-grade high-speed tool to parse email attachments starting with"
    " **OutputBlankDR**, prevent duplicate orders, and merge data starting from"
    " **row 6** in `Output.xlsx`."
)

# 1. Secure Credential Management & Persistence via Session State
if "email_pass" not in st.session_state:
  default_pass = ""
  try:
    if (
        hasattr(st, "secrets")
        and "email" in st.secrets
        and "app_password" in st.secrets["email"]
    ):
      default_pass = st.secrets["email"]["app_password"]
  except Exception:
    pass
  st.session_state.email_pass = default_pass

st.sidebar.header("🔐 Credentials & Config")
email_user = st.sidebar.text_input(
    "Email Address", value="harmesh.kumar@aseesfeed.com"
)
email_pass = st.sidebar.text_input(
    "Email App Password",
    value=st.session_state.email_pass,
    type="password",
    help="Enter Gmail App Password or store in Streamlit secrets.",
)
st.session_state.email_pass = email_pass

imap_server = st.sidebar.selectbox(
    "IMAP Server", ["imap.gmail.com", "imap.mail.yahoo.com", "imap.office365.com"]
)
mailbox_name = st.sidebar.text_input(
    "Mailbox / Folder Name",
    value="BOT",
    help="Folder where bot emails arrive (e.g., BOT)",
)

st.subheader("⚙️ Processing Parameters")
col1, col2 = st.columns(2)
with col1:
  selected_date = st.date_input("Select Target Date", value=datetime.today())
with col2:
  subject_filter = st.text_input(
      "Subject Filter (Optional)",
      value="Nestle Bot Partailly Completed & Missing DR Codes",
      help="Leave blank to process all emails on this date.",
  )

deduplicate_flag = st.checkbox(
    "Prevent Duplicate Orders (Unique Check)",
    value=True,
    help=(
        "Ensures identical rows (Order ID + Product + Item) are not added twice."
    ),
)

TEMPLATE_FILENAME = "Output.xlsx"

# Audit log storage in session state
if "audit_logs" not in st.session_state:
  st.session_state.audit_logs = []

if st.button("🚀 Run Automation & Merge Data", type="primary"):
  if not email_pass:
    st.warning("Please enter your Email App Password.")
  elif not os.path.exists(TEMPLATE_FILENAME):
    st.error(
        f"Template file '{TEMPLATE_FILENAME}' not found in the repository!"
    )
  else:
    download_folder = "temp_downloads"
    os.makedirs(download_folder, exist_ok=True)

    formatted_date = selected_date.strftime("%d-%m-%Y")
    imap_date = selected_date.strftime("%d-%b-%Y")

    progress_bar = st.progress(0)
    status_text = st.empty()

    try:
      status_text.text("Connecting to mail server...")
      progress_bar.progress(10)

      mail = imaplib.IMAP4_SSL(imap_server)
      mail.login(email_user, email_pass)

      rv, data = mail.select(mailbox_name)
      if rv != "OK":
        st.error(
            f"Could not select folder '{mailbox_name}'. Please verify the name."
        )
        mail.logout()
      else:
        status_text.text("Searching emails for the selected date...")
        progress_bar.progress(30)

        search_query = f'(ON "{imap_date}")'
        status, messages = mail.search(None, search_query)

        if status != "OK" or not messages[0]:
          st.warning(
              f"No emails found in folder '{mailbox_name}' for date"
              f" {formatted_date}."
          )
          mail.logout()
        else:
          msg_nums = messages[0].split()
          matching_nums = []

          status_text.text("Filtering emails by subject...")
          progress_bar.progress(50)

          if subject_filter.strip():
            for num in msg_nums:
              try:
                res, h_data = mail.fetch(
                    num, "(BODY.PEEK[HEADER.FIELDS (Subject)])"
                )
                for response_part in h_data:
                  if isinstance(response_part, tuple):
                    msg_header = email.message_from_bytes(response_part[1])
                    raw_subject = msg_header["Subject"]
                    if raw_subject:
                      decoded_subj_list = decode_header(raw_subject)
                      subject_str = ""
                      for sub, enc in decoded_subj_list:
                        if isinstance(sub, bytes):
                          subject_str += sub.decode(
                              enc or "utf-8", errors="ignore"
                          )
                        else:
                          subject_str += str(sub)

                      if (
                          subject_filter.strip().lower() in subject_str.lower()
                      ):
                        matching_nums.append(num)
              except Exception:
                continue
          else:
            matching_nums = msg_nums

          if not matching_nums:
            st.warning("Found emails for this date, but none matched the criteria.")
            mail.logout()
          else:
            status_text.text(
                f"Found {len(matching_nums)} matching emails. Downloading"
                " attachments..."
            )
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
                            decoded_header_val[0][1] or "utf-8",
                            errors="ignore",
                        )

                      # Strict filter for attachments
                      if filename.startswith("OutputBlankDR") and filename.endswith(".xlsx"):
                        filepath = os.path.join(download_folder, filename)
                        with open(filepath, "wb") as f:
                          f.write(part.get_payload(decode=True))
                        downloaded_files.append(filepath)

            mail.logout()

            if downloaded_files:
              status_text.text("Merging data and updating template...")
              progress_bar.progress(90)

              output_filename = f"Master_OutputBlankDR_{formatted_date.replace('-', '_')}.xlsx"
              output_path = os.path.join(download_folder, output_filename)

              wb = openpyxl.load_workbook(TEMPLATE_FILENAME)

              if "Order Data" in wb.sheetnames:
                ws = wb["Order Data"]
                current_row = 6

                existing_keys = set()
                if deduplicate_flag:
                  # Read existing rows in master to prevent duplicates
                  for r in range(6, ws.max_row + 1):
                    order_id = ws.cell(row=r, column=1).value
                    product_id = ws.cell(row=r, column=15).value
                    item_id = ws.cell(row=r, column=14).value
                    if order_id:
                      existing_keys.add((str(order_id), str(product_id), str(item_id)))

                total_added = 0
                total_duplicates = 0

                for file in downloaded_files:
                  try:
                    df_source = pd.read_excel(
                        file, sheet_name="Order Data", header=None
                    )
                    data_to_paste = df_source.iloc[5:].values.tolist()

                    for row_data in data_to_paste:
                      if len(row_data) >= 15:
                        order_id = str(row_data[0]) if row_data[0] is not None else ""
                        product_id = str(row_data[14]) if len(row_data) > 14 and row_data[14] is not None else ""
                        item_id = str(row_data[13]) if len(row_data) > 13 and row_data[13] is not None else ""
                        unique_key = (order_id, product_id, item_id)

                        if deduplicate_flag and unique_key in existing_keys:
                          total_duplicates += 1
                          continue

                        existing_keys.add(unique_key)

                      for col_idx, val in enumerate(row_data, start=1):
                        ws.cell(row=current_row, column=col_idx, value=val)
                      current_row += 1
                      total_added += 1

                  except Exception as e:
                    st.error(f"Error reading file {file}: {e}")

                wb.save(output_path)
                progress_bar.progress(100)
                status_text.text("Processing complete!")

                st.success(
                    f"Successfully processed {len(downloaded_files)} file(s)."
                    f" Added {total_added} row(s) starting from row 6."
                )
                if deduplicate_flag and total_duplicates > 0:
                  st.info(f"Skipped {total_duplicates} duplicate row(s).")

                # Log activity in audit trail
                st.session_state.audit_logs.insert(
                    0,
                    {
                        "Timestamp": datetime.now().strftime(
                            "%Y-%m-%d %H:%M:%S"
                        ),
                        "Date": formatted_date,
                        "Files Processed": len(downloaded_files),
                        "Rows Added": total_added,
                        "Duplicates Skipped": total_duplicates,
                    },
                )

                with open(output_path, "rb") as f:
                  st.download_button(
                      label="📥 Download Final Master Excel File",
                      data=f,
                      file_name=output_filename,
                      mime=(
                          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                      ),
                  )
              else:
                st.error("Sheet 'Order Data' not found in template.")
            else:
              progress_bar.progress(100)
              status_text.text("Done.")
              st.warning(
                  "Matching emails found, but no attachments started with"
                  " 'OutputBlankDR'."
              )

    except Exception as e:
        progress_bar.progress(100)
        st.error(f"An unexpected error occurred: {e}")

# 3. Audit Log History Section
if st.session_state.audit_logs:
  st.markdown("---")
  st.subheader("📋 Session Audit Trail / History")
  log_df = pd.DataFrame(st.session_state.audit_logs)
  st.dataframe(log_df, use_container_width=True)
