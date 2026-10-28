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
    "Template Excel file (output.xlsx) me mail ke attachments ka data row 6 se"
    " paste karne ka tool."
)

st.sidebar.header("Email Credentials & Settings")
email_user = st.sidebar.text_input(
    "Email Address", value="harmesh.kumar@aseesfeed.com"
)
email_pass = st.sidebar.text_input(
    "Email App Password", type="password", help="Gmail App Password use karein"
)

imap_server = st.sidebar.selectbox(
    "IMAP Server", ["imap.gmail.com", "imap.mail.yahoo.com", "imap.office365.com"]
)
mailbox_name = st.sidebar.text_input(
    "Mailbox / Folder Name",
    value="BOT",
    help="Jis folder me mails hain uska naam dalein (jaise BOT)",
)

st.subheader("Search Parameters")
selected_date = st.date_input("Select Date", value=datetime.today())
subject_filter = st.text_input(
    "Subject Filter", value="Nestle Bot Partailly Completed & Missing DR Codes"
)

# GitHub par upload ki gayi template file ka naam yahan 'output.xlsx' set hai
TEMPLATE_FILENAME = "output.xlsx"

if st.button("Process & Combine Files", type="primary"):
  if not email_pass:
    st.warning("Kripya apna Email App Password enter karein.")
  elif not os.path.exists(TEMPLATE_FILENAME):
    st.error(
        f"Template file '{TEMPLATE_FILENAME}' GitHub repository me nahi mili!"
        " Kripya file ka naam check karein."
    )
  else:
    download_folder = "temp_downloads"
    os.makedirs(download_folder, exist_ok=True)

    formatted_date = selected_date.strftime("%d-%m-%Y")
    imap_date = selected_date.strftime("%d-%b-%Y")

    with st.spinner(
        f"Connecting to email server and checking '{mailbox_name}' folder..."
    ):
      try:
        mail = imaplib.IMAP4_SSL(imap_server)
        mail.login(email_user, email_pass)

        rv, data = mail.select(mailbox_name)
        if rv != "OK":
          st.error(
              f"'{mailbox_name}' folder select nahi ho paya. Kripya folder ka"
              " naam check karein."
          )
          mail.logout()
        else:
          search_query = f'(ON "{imap_date}")'
          status, messages = mail.search(None, search_query)

          if status != "OK" or not messages[0]:
            st.warning(
                f"Is date ({formatted_date}) par '{mailbox_name}' folder me koi"
                " mail nahi mila."
            )
            mail.logout()
          else:
            downloaded_files = []
            matched_mails_count = 0

            for num in messages[0].split():
              res, msg_data = mail.fetch(num, "(RFC822)")
              for response_part in msg_data:
                if isinstance(response_part, tuple):
                  msg = email.message_from_bytes(response_part[1])

                  raw_subject = msg["Subject"]
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
                  else:
                    subject_str = ""

                  if subject_filter.strip().lower() in subject_str.lower():
                    matched_mails_count += 1
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

                        if (
                            "OutputBlankDR" in filename
                            and filename.endswith(".xlsx")
                        ):
                          filepath = os.path.join(download_folder, filename)
                          with open(filepath, "wb") as f:
                            f.write(part.get_payload(decode=True))
                          downloaded_files.append(filepath)

            mail.logout()

            if matched_mails_count == 0:
              st.warning(
                  f"Is date par mails mile, lekin unka subject '{subject_filter}'"
                  " se match nahi hua."
              )
            elif downloaded_files:
              st.success(
                  f"Total {matched_mails_count} matching mails mile aur"
                  f" {len(downloaded_files)} files successfully download ho"
                  " gayi hain!"
              )

              output_filename = f"Master_OutputBlankDR_{formatted_date.replace('-', '_')}.xlsx"
              output_path = os.path.join(download_folder, output_filename)

              # 'output.xlsx' template file load karein
              wb = openpyxl.load_workbook(TEMPLATE_FILENAME)

              if "Order Data" in wb.sheetnames:
                ws = wb["Order Data"]

                # Har downloaded file se row 6 (index 5) se aage ka data append karein
                for file in downloaded_files:
                  try:
                    df_source = pd.read_excel(
                        file, sheet_name="Order Data", header=None
                    )
                    data_to_paste = df_source.iloc[5:].values.tolist()

                    for row_data in data_to_paste:
                      ws.append(row_data)
                  except Exception as e:
                    st.error(f"Error reading file {file}: {e}")

                wb.save(output_path)

                st.info(
                    "Template file me row 6 se data successfully paste ho gaya"
                    " hai! Neeche button se download karein:"
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
                st.error("Template file me 'Order Data' sheet nahi mili.")
            else:
              st.warning(
                  "Matching emails mile, lekin unme koi 'OutputBlankDR' excel"
                  " file attachment nahi thi."
              )

      except Exception as e:
        st.error(f"An error occurred: {e}")
