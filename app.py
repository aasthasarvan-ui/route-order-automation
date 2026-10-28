from datetime import datetime
import email
from email.header import decode_header
import imaplib
import os
import pandas as pd
import streamlit as st

# Streamlit Page Configuration
st.set_page_config(
    page_title="Nestle DR Downloader & Combiner",
    page_icon="📥",
    layout="centered",
)

st.title("📥 Nestle Bot DR Downloader & Merger")
st.write(
    "Mozilla Thunderbird / Gmail se specific date ke **OutputBlankDR** excel"
    " files ko download aur combine karne ke liye tool."
)

# Sidebar Credentials & Settings
st.sidebar.header("Email Credentials")
email_user = st.sidebar.text_input(
    "Email Address", value="harmesh.kumar@aseesfeed.com"
)
email_pass = st.sidebar.text_input(
    "Email App Password", type="password", help="Gmail App Password use karein"
)

imap_server = st.sidebar.selectbox(
    "IMAP Server", ["imap.gmail.com", "imap.mail.yahoo.com", "imap.office365.com"]
)

# Main UI inputs
st.subheader("Search Parameters")
selected_date = st.date_input("Select Date", value=datetime.today())
subject_filter = st.text_input(
    "Subject Filter", value="Nestle Bot Partailly Completed & Missing DR Codes"
)

if st.button("Process & Combine Files", type="primary"):
  if not email_pass:
  else:
    download_folder = "temp_downloads"
    os.makedirs(download_folder, exist_ok=True)

    formatted_date = selected_date.strftime("%d-%m-%Y")
    imap_date = selected_date.strftime("%d-%b-%Y")

    with st.spinner("Connecting to email server and searching emails..."):
      try:
        mail = imaplib.IMAP4_SSL(imap_server)
        mail.login(email_user, email_pass)
        mail.select("inbox")

        search_query = f'(ON "{imap_date}" SUBJECT "{subject_filter}")'
        status, messages = mail.search(None, search_query)

        if status != "OK" or not messages[0]:
          st.warning(
              f"Is date ({formatted_date}) par koi matching mail nahi mila."
          )
          mail.logout()
        else:
          downloaded_files = []
          for num in messages[0].split():
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

                    if (
                        "OutputBlankDR" in filename
                        and filename.endswith(".xlsx")
                    ):
                      filepath = os.path.join(download_folder, filename)
                      with open(filepath, "wb") as f:
                        f.write(part.get_payload(decode=True))
                      downloaded_files.append(filepath)

          mail.logout()

          if downloaded_files:
            st.success(
                f"Total {len(downloaded_files)} files successfully download ho"
                " gayi hain!"
            )

            # Combine Excel Files
            all_data = []
            for file in downloaded_files:
              try:
                df = pd.read_excel(file)
                all_data.append(df)
              except Exception as e:
                st.error(f"Error reading file {file}: {e}")

            if all_data:
              combined_df = pd.concat(all_data, ignore_index=True)
              output_filename = f"Master_OutputBlankDR_{formatted_date.replace('-', '_')}.xlsx"
              output_path = os.path.join(download_folder, output_filename)
              combined_df.to_excel(output_path, index=False)

              st.info("Files combine ho chuki hain. Neeche button se download karein:")
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
                "Matching emails mile, lekin unme koi 'OutputBlankDR' excel file"
                " attachment nahi thi."
            )

      except Exception as e:
        st.error(f"An error occurred: {e}")
