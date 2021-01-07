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
                        pdf_merger(folder)
                    pdf_merger(root_path)
    print('Script Finished running!!!')
    # time.sleep(120)

if __name__ == '__main__':
    main()