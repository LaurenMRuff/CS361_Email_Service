# AUTHOR: Lauren Ruff
# Email: ruffl@oregonstate.edu
# Assignment: 8, Integration
# Due Date: February 28, 2022
# Version: 1.0
# File: Email_Service.py
# Description: This is an email microservice created for Ethan's Project and also used for Lauren's project. It will
#              allow the user to send an email using content from a text file. It is required that the email comes from
#              a gmail account, as the user must log in to the account to send an email. Recipient can be any email
#              provider.

import base64
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
import tkinter as tk
from tkinter import ttk
import os.path
import time
import pickle
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import platform


def popup(msg):
    # popup window to tell user password was incorrect
    popup_wind = tk.Tk()
    popup_wind.geometry("400x50")
    popup_wind.eval('tk::PlaceWindow . center')
    ttk.Label(popup_wind, text=msg).pack()

    popup_wind.after(5000, lambda: popup_wind.destroy())  # destroy window after 5 seconds

    popup_wind.mainloop()


def connect_to_gmail(s):
    # COMMENT: Using smtp to log in to gmail code adapted from code found at link
    # DATE: February 25, 2022
    # SOURCE: https://docs.python.org/3/library/email.examples.html
    # https://stackoverflow.com/questions/28421887/django-email-with-smtp-gmail-smtpauthenticationerror-534-application-specific-pa

    SCOPES = 'https://www.googleapis.com/auth/gmail.send'

    creds = None

    email = s.split('@')

    pickle_file = "token_" + email[0] + ".pickle"

    if os.path.exists(pickle_file):
        with open(pickle_file, 'rb') as token:
            creds = pickle.load(token)
        # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open(pickle_file, 'wb') as token:
            pickle.dump(creds, token)

    service = build('gmail', 'v1', credentials=creds)

    return service


def generate_email(input_file, dir_path, slash):

    # open the text file and process it
    with open(input_file) as f:
        lines = f.readlines()

    sender = lines[0].rstrip()
    recipient = lines[1].rstrip()
    subject = lines[2].rstrip()
    attachment = lines[3].rstrip()
    message = ''.join(lines[4:]).rstrip()

    email = MIMEMultipart('mixed')

    email['to'] = recipient
    email['from'] = sender
    email['Subject'] = subject

    message_alt = MIMEMultipart('alternative')
    message_alt.attach(MIMEText(message, 'plain'))

    email.attach(message_alt)

    if len(attachment) > 0:
        try:
            with open(attachment, 'rb') as img_f:
                img_data = img_f.read()
                img_f.close()
            email.attach(MIMEImage(img_data))
        except FileNotFoundError:
            popup("ERROR: File could not be found. File will not be attached to email.")

    server = connect_to_gmail(sender)

    msg_raw = {'raw': base64.urlsafe_b64encode(email.as_string().encode()).decode()}

    try:
        server.users().messages().send(userId="me", body=msg_raw).execute()
        popup('Email sent successfully!')
        f = open(dir_path + slash + "success.txt", 'w')
        f.close()

    except server.SMTPAuthenticationError:
        popup('ERROR: Email failed to send')
        f = open(dir_path + slash + "fail.txt", 'w')
        f.close()

    except server.SMTPDataError:
        popup('ERROR: Issue with message data')
        f = open(dir_path + slash + "fail.txt", 'w')
        f.close()

    except server.SMTPSenderRefused:
        popup('ERROR: Recipient email is invalid')
        f = open(dir_path + slash + "fail.txt", 'w')
        f.close()


def get_system():

    local_dir = ''
    slash = ''

    if platform.system() == "Darwin":  # for Mac
        local_dir = os.environ['HOME'] + '/Desktop/email_service_data'
        slash = '/'
    elif platform.system() == "Windows":
        local_dir = os.environ['USERPROFILE'] + '\\Desktop\\email_service_data'
        slash = '\\'

    return local_dir, slash


def email_service():

    # COMMENT: code for waiting for a file to exist copied from source link
    # DATE: February 26, 2022
    # SOURCE : https://stackoverflow.com/questions/21746750/check-and-wait-until-a-file-exists-to-read-it

    dir_path, slash = get_system()

    file_path = dir_path + slash + "email_data.txt"

    while True:
        # wait for the file to exist
        while not os.path.exists(file_path):
            time.sleep(1)

        # if it is the file we are looking for, process it. Otherwise, print an error and exit
        if os.path.isfile(file_path):
            generate_email(file_path, dir_path, slash)

        else:
            popup("email_data.txt was an invalid file.")

        os.remove(file_path)


if __name__ == "__main__":
    email_service()
