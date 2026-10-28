from datetime import datetime
import email
from email.header import decode_header
import imaplib
import os
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Nestle DR Downloader & Combiner",
    page_icon="📥",
    layout="centered",
)

st.title("📥 Nestle Bot DR Downloader & Merger")
st.write(
    "Mozilla Thunderbird / Gmail se specific folder aur date ke"
    " **OutputBlankDR** excel files ko download aur combine karne ke liye tool."
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

if st.button("Process & Combine Files", type="primary"):
  if not email_pass:
    st.warning("Kripya apna Email App Password enter karein.")
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

              all_order_data = []
              template_file = downloaded_files[
                  0
              ]  # Pehli file ka structure template ke liye use hoga

              # Sabhi files se 'Order Data' read karna
              for idx, file in enumerate(downloaded_files):
                try:
                  df_raw = pd.read_excel(
                      file, sheet_name="Order Data", header=None
                  )
                  # Row 4 ke baad ka data (yani row 5 index se jo Excel ki row 6 hai) extract karna
                  data_rows = df_raw.iloc[4:]
                  all_order_data.append(data_rows)
                except Exception as e:
                  st.error(f"Error reading file {file}: {e}")

              if all_order_data:
                # Pehli file ki pehli 4 rows (instructions aur template headers) ko base banana
                df_template = pd.read_excel(
                    template_file, sheet_name="Order Data", header=None
                )
                header_structure = df_template.iloc[0:4]

                # Saare data rows ko ek sath concatenate karna
                combined_data = pd.concat(all_order_data, ignore_index=True)

                # Final DataFrame banana jisme row 6 se data shuru ho
                final_order_df = pd.concat(
                    [header_structure, combined_data], ignore_index=True
                )

                output_filename = f"Master_OutputBlankDR_{formatted_date.replace('-', '_')}.xlsx"
                output_path = os.path.join(download_folder, output_filename)

                # Excel file ko multi-sheet format me save karna (Field List aur Order Data dono ke sath)
                with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
                  # Field List sheet copy karna pehli file se
                  df_field_list = pd.read_excel(
                      template_file, sheet_name="Field List", header=None
                  )
                  df_field_list.to_excel(
                      writer,
                      sheet_name="Field List",
                      index=False,
                      header=False,
                  )

                  # Order Data sheet save karna jisme row 6 se data paste ho
                  final_order_df.to_excel(
                      writer,
                      sheet_name="Order Data",
                      index=False,
                      header=False,
                  )

                st.info(
                    "Files successfully combine ho chuki hain aur row 6 se data"
                    " paste ho gaya hai. Neeche button se download karein:"
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
              st.warning(
                  "Matching emails mile, lekin unme koi 'OutputBlankDR' excel"
                  " file attachment nahi thi."
              )

      except Exception as e:
        st.error(f"An error occurred: {e}")
