import imaplib
import base64
import os
import email
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
# import pypdftk
import time
import sys
# from pdf2image import pdf2image
from datetime import datetime
import glob
import mailparser
from PyPDF2 import PdfFileMerger

email_user = "quality.graham@gmail.com"
email_pw = "mominejaz"


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
        pdf_list.sort()

        merger(merged_file_name, pdf_list)
    except:
        pass

def fast_scandir(dirname):
    subfolders= [f.path for f in os.scandir(dirname) if f.is_dir()]
    for dirname in list(subfolders):
        subfolders.extend(fast_scandir(dirname))
    return subfolders

def main():
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
                root_path = body[0]
                root_path = r'{}'.format(root_path)
                folder_list = fast_scandir(root_path)
                if folder_list == []:
                    pdf_merger(root_path)
                else:
                    for folder in folder_list:
                        folder_path = root_path+'\\'+folder
                        pdf_merger(folder_path)
    # 
    # 
    #             # instance of MIMEMultipart
    #             msg = MIMEMultipart()
    # 
    #             # storing the senders email address
    #             msg['From'] = email_user
    # 
    #             # storing the receivers email address
    #             msg['To'] = toaddr
    # 
    #             # storing the subject
    #             msg['Subject'] = "{} - Merged and Flattened PDF at {}".format(tosubject, current_time)
    # 
    #             # string to store the body of the mail
    #             body = "Please view the attached PDF generated at {}.".format(current_time)
    # 
    #             # attach the body with the msg instance
    #             msg.attach(MIMEText(body, 'plain'))
    # 
    #             # open the file to be sent
    #             attachment = open(merged_file, "rb")
    # 
    #             # instance of MIMEBase and named as p
    #             p = MIMEBase('application', 'octet-stream')
    # 
    #             # To change the payload into encoded form
    #             p.set_payload((attachment).read())
    # 
    #             # encode into base64
    #             encoders.encode_base64(p)
    # 
    #             p.add_header('Content-Disposition', "attachment; filename= %s" % merged_file_name)
    # 
    #             # attach the instance 'p' to instance 'msg'
    #             msg.attach(p)
    # 
    #             # creates SMTP session
    #             s = smtplib.SMTP('smtp.gmail.com', 587)
    # 
    #             # start TLS for security
    #             s.starttls()
    # 
    #             # Authentication
    #             s.login(email_user, email_pw)
    # 
    #             # Converts the Multipart msg into a string
    #             text = msg.as_string()
    # 
    #             # sending the mail
    #             s.sendmail(email_user, toaddr, text)
    # 
    #             # terminating the session
    #             s.quit()
    # 
    #             attachment.close()
    # 
    #             os.remove(merged_file)
    # 
    #             files_list = []
    #             files_list_unspaced = []
    #             files_list_flattened = []
    #             bad_files_list = []
    # tosubject = ''
    # toaddr = ''
    # mail.close()
    print('Merged PDFs!!!')
    # time.sleep(120)

if __name__ == '__main__':
    main()