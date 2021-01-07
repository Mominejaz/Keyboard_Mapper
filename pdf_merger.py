import imaplib
import os
import email
import time
import glob
import mailparser
from PyPDF2 import PdfFileMerger
from email.message import EmailMessage
import smtplib

email_user = "quality.graham@gmail.com"
email_pw = "mominejaz"


def email_auto_script(to_email, subject, body, file_path):
    print('Sending Email!!!')

    gmail_user = "quality.graham@gmail.com"
    gmail_password = "mominejaz"

    try:
        msg = EmailMessage()
        msg["Subject"] = subject
        msg["From"] = gmail_user
        msg["To"] = to_email
        msg.set_content(body)
        # Raw text for path needs to be added here
        if file_path == []:
            for file in file_path:
                print(file)
                with open(file, "rb") as f:
                    file_data = f.read()
                    file_name = os.path.split(file)
                    file_name = file_name[-1]

                    msg.add_attachment(
                        file_data,
                        maintype="application",
                        subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        filename=file_name,
                    )
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
            smtp.login(gmail_user, gmail_password)
            smtp.send_message(msg)
            print("Email sent!")
    except:
        print("Something went wrong...Email not Sent.")


def merger(merged_file_name, input_paths):
    pdf_merger = PdfFileMerger()
    file_handles = []

    for path in input_paths:
        pdf_merger.append(path)
    with open(merged_file_name, 'wb') as fileobj:
        pdf_merger.write(fileobj)

def pdf_merger(input_path):
    try:
        folder_name = os.path.split(input_path)
        folder_name = folder_name[-1]
        merged_file_name = input_path + '\\' + folder_name + '.pdf'
        pdf_list = glob.glob(f'{input_path}/*.pdf')
        if pdf_list == []:
            raise Exception
        else:
            pdf_list.sort()
            merger(merged_file_name, pdf_list)
            print(f'Merged PDF: {merged_file_name}')
    except:
        print('No pdf to merge!')

def fast_scandir(dirname):
    subfolders= [f.path for f in os.scandir(dirname) if f.is_dir()]
    for dirname in list(subfolders):
        subfolders.extend(fast_scandir(dirname))
    return subfolders

def main():
    try:

        while True:
            mail = imaplib.IMAP4_SSL("imap.gmail.com")
            mail.login(email_user, email_pw)
            mail.select('Inbox')

            typ, items = mail.search(None, '(UNSEEN SUBJECT PDF_Merge)')

            if typ == "OK":
                mail_ids = items[0]
                items = items[0].split()
                if len(items) > 0:
                    for num in items:
                        typ, data = mail.fetch(num, '(RFC822)')
                        raw_email = data[0][1]
                        raw_email_string = raw_email.decode('utf-8')
                        email_message = email.message_from_string(raw_email_string)
                        #mailparser
                        mail = mailparser.parse_from_string(raw_email_string)
                        body = mail.body.splitlines()
                        email_from = mail.from_
                        root_path = body[0]
                        root_path = r'{}'.format(root_path)
                        folder_list = fast_scandir(root_path)
                        if folder_list == []:
                            pdf_merger(root_path)
                        else:
                            for folder in folder_list:
                                pdf_merger(folder)
                                time.sleep(2)
                            pdf_merger(root_path)
                else:
                    raise Exception


    except WindowsError:
        email_body = "Hello,\n\n" \
                     "PDF Merging failed\n" \
                     f"The directory cannot be accessed. [{root_path}] \n" \
                     "Please put the folders in another directory and re-send the email!\n\n" \
                     "Best Regards,\n" \
                     "Graham Automation"
        email_auto_script([email_from[0][1]],"PDF_Merge Error", email_body,file_path=[])
    except Exception:
        print("No request to merge pdfs found restarting in 30s")

    finally:
        time.sleep(30)
        main()

if __name__ == '__main__':
    main()