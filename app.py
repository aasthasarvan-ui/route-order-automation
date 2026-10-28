from datetime import datetime
import email
from email.header import decode_header
import imaplib
import os
import openpyxl
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Nestle DR Downloader & Combiner",
    page_icon="📥",
    layout="centered",
)

st.title("📥 Nestle Bot DR Downloader & Merger")
st.write(
    "High-speed tool to download attachments starting with 'OutputBlankDR'"
    " (subject is optional) and paste data starting from row 6 in the template"
    " Excel file (Output.xlsx)."
)

# Load email password securely from Streamlit secrets if available, otherwise input field
default_password = ""
try:
  if (
      hasattr(st, "secrets")
      and "email" in st.secrets
      and "app_password" in st.secrets["email"]
  ):
    default_password = st.secrets["email"]["app_password"]
except Exception:
  pass

st.sidebar.header("Email Credentials & Settings")
email_user = st.sidebar.text_input(
    "Email Address", value="harmesh.kumar@aseesfeed.com"
)
email_pass = st.sidebar.text_input(
    "Email App Password",
    value=default_password,
    type="password",
    help=(
        "Enter your Gmail App Password. You can also save it in Streamlit Secrets"
        " (.streamlit/secrets.toml)"
    ),
)

imap_server = st.sidebar.selectbox(
    "IMAP Server", ["imap.gmail.com", "imap.mail.yahoo.com", "imap.office365.com"]
)
mailbox_name = st.sidebar.text_input(
    "Mailbox / Folder Name",
    value="BOT",
    help="Enter the folder name where emails are located (e.g., BOT)",
)

st.subheader("Search Parameters")
selected_date = st.date_input("Select Date", value=datetime.today())
# Subject filter is optional (leave blank to process all emails for that date)
subject_filter = st.text_input(
    "Subject Filter (Optional - Leave blank to ignore subject)",
    value="Nestle Bot Partailly Completed & Missing DR Codes",
)

TEMPLATE_FILENAME = "Output.xlsx"

if st.button("Process & Combine Files (High Speed)", type="primary"):
  if not email_pass:
    st.warning("Please enter your Email App Password.")
  elif not os.path.exists(TEMPLATE_FILENAME):
    st.error(
        f"Template file '{TEMPLATE_FILENAME}' not found in the GitHub repository!"
        " Please check the file name."
    )
  else:
    download_folder = "temp_downloads"
    os.makedirs(download_folder, exist_ok=True)

    formatted_date = selected_date.strftime("%d-%m-%Y")
    imap_date = selected_date.strftime("%d-%b-%Y")

    with st.spinner(
        f"Connecting to server and searching '{mailbox_name}' folder..."
    ):
      try:
        mail = imaplib.IMAP4_SSL(imap_server)
        mail.login(email_user, email_pass)

        rv, data = mail.select(mailbox_name)
        if rv != "OK":
          st.error(
              f"Could not select folder '{mailbox_name}'. Please verify the"
              " folder name."
          )
          mail.logout()
        else:
          search_query = f'(ON "{imap_date}")'
          status, messages = mail.search(None, search_query)

          if status != "OK" or not messages[0]:
            st.warning(
                f"No emails found in '{mailbox_name}' folder for the date"
                f" ({formatted_date})."
            )
            mail.logout()
          else:
            msg_nums = messages[0].split()
            matching_nums = []

            # If subject filter is provided, filter by subject. If blank, accept all emails for that date.
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
                            subject_filter.strip().lower()
                            in subject_str.lower()
                        ):
                          matching_nums.append(num)
                except Exception:
                  continue
            else:
              matching_nums = msg_nums  # Take all messages if subject is blank

            if not matching_nums:
              st.warning(
                  f"Found emails for this date, but none matched the subject"
                  f" criteria."
              )
              mail.logout()
            else:
              st.info(
                  f"Found {len(matching_nums)} email(s) to process. Checking"
                  " attachments..."
              )
              downloaded_files = []

              # Download attachments strictly starting with 'OutputBlankDR'
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

                        # STRICT FILTER: Filename must start with 'OutputBlankDR'
                        if filename.startswith("OutputBlankDR") and filename.endswith(".xlsx"):
                          filepath = os.path.join(download_folder, filename)
                          with open(filepath, "wb") as f:
                            f.write(part.get_payload(decode=True))
                          downloaded_files.append(filepath)

              mail.logout()

              if downloaded_files:
                st.success(
                    f"Successfully downloaded {len(downloaded_files)} Excel"
                    " attachment(s) starting with 'OutputBlankDR'!"
                )

                output_filename = f"Master_OutputBlankDR_{formatted_date.replace('-', '_')}.xlsx"
                output_path = os.path.join(download_folder, output_filename)

                # Load template file Output.xlsx
                wb = openpyxl.load_workbook(TEMPLATE_FILENAME)

                if "Order Data" in wb.sheetnames:
                  ws = wb["Order Data"]
                  current_row = 6

                  for file in downloaded_files:
                    try:
                      df_source = pd.read_excel(
                          file, sheet_name="Order Data", header=None
                      )
                      data_to_paste = df_source.iloc[5:].values.tolist()

                      for row_data in data_to_paste:
                        for col_idx, val in enumerate(row_data, start=1):
                          ws.cell(row=current_row, column=col_idx, value=val)
                        current_row += 1  # Correctly incremented inside row loop
                    except Exception as e:
                      st.error(f"Error reading file {file}: {e}")

                  wb.save(output_path)

                  st.info(
                      "Data has been successfully pasted starting from row 6 in"
                      " the template file! Click the button below to download:"
                  )
                  with open(output_path, "rb") as f:
                    st.download_button(
                        label="📥 Download Master Excel File",
                        data=f,
                        file_name=output_filename,
                        mime=(
                            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        ),
                    )
                else:
                  st.error(
                      "Sheet 'Order Data' was not found in the template file."
                  )
              else:
                st.warning(
                    "No attachments starting with 'OutputBlankDR' were found in"
                    " the searched emails."
                )

      except Exception as e:
        st.error(f"An error occurred: {e}")
