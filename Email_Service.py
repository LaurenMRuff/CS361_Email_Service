# AUTHOR: Lauren Ruff
# Email: ruffl@oregonstate.edu
# Assignment: 10, Portfolio
# Due Date: March 18, 2022
# Version: 2.0
# File: Email_Service.py
# Description: This is an email microservice created for Ethan's Project and also used for Lauren's project. It will
#              allow the user to send an email using content from a text file. It is required that the email comes from
#              a gmail account, as the user must log in to the account to send an email. Recipient can be any email
#              provider.

import base64
import sys
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
from datetime import datetime


def popup(msg):
    """
    this is a popup message gui that shows the passed message to the user
    :param msg: message to be shown in the popup
    """
    popup_wind = tk.Tk()
    popup_wind.geometry("400x50")
    popup_wind.eval('tk::PlaceWindow . center')
    ttk.Label(popup_wind, text=msg).pack()

    popup_wind.after(5000, lambda: popup_wind.destroy())  # destroy window after 5 seconds

    popup_wind.mainloop()


def connect_to_gmail(s):
    """
    uses oauth2 to connect to gmail securely
    :param s: the senders email
    :return: the authorized connection to gmail
    """
    # COMMENT: Using smtp to log in to gmail code adapted from code found at link
    # DATE: February 25, 2022
    # SOURCE: https://docs.python.org/3/library/email.examples.html
    # https://stackoverflow.com/questions/28421887/django-email-with-smtp-gmail-smtpauthenticationerror-534-application-specific-pa

    SCOPES = 'https://www.googleapis.com/auth/gmail.send'

    creds = None

    email = s.split('@')

    pickle_file = "token_" + email[0] + ".pickle"

    if os.path.exists(pickle_file):
        if os.stat(pickle_file).st_mtime < time.time() - 3 * 86400:
            # if the pickle file is more than 3 days old, it is deleted and the user needs to re-authenticate
            os.remove(pickle_file)
            creds = None
        else:
            # else, the file is opened and the token is read
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
    """
    This method parses the text file, then generates and sends the email with the contents of the text file
    :param input_file: location of the file with the email data
    :param dir_path: path to the directory of this application
    :param slash: direction of the file structure slashes in relation to the operating system
    """
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
        # if there is an attachment, add it
        try:
            with open(attachment, 'rb') as img_f:
                image_to_attach = MIMEImage(img_f.read(), name=os.path.basename(attachment))
                img_f.close()
            email.attach(image_to_attach)
        except FileNotFoundError:
            popup("ERROR: File could not be found. File will not be attached to email.")

    server = connect_to_gmail(sender)

    # generate the raw message
    msg_raw = {'raw': base64.urlsafe_b64encode(email.as_string().encode()).decode()}

    try:
        server.users().messages().send(userId="me", body=msg_raw).execute()
        popup('Email sent successfully!')

        # write the data from the email_data.txt file to a new text file to keep as a record of what was send
        now = datetime.now()
        email_sent_path = dir_path + slash + "email_data_" + now.strftime("%m.%d.%Y_%H%M") + ".txt"
        with open(email_sent_path, 'w') as f:
            for line in lines:
                f.write(line)
        f.close()

    # all possible errors
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
    """
    This method checks for the operating system on the source computer to appropriately structure files
    Works for Windows and Mac
    :return: the local directory, and slash direction
    """
    local_dir = ''
    slash = ''

    if platform.system() == "Darwin":  # for Mac
        local_dir = os.environ['HOME'] + '/Desktop/email_service_data'
        slash = '/'
    elif platform.system() == "Windows":  # for windows
        local_dir = os.environ['USERPROFILE'] + '\\Desktop\\email_service_data'
        slash = '\\'

    return local_dir, slash


def email_service():
    """
    main method for starting the email service
    gets the system and file path to the file it is looking for with the email data
    checks for a new file, creates a new email if the file is less than 3 seconds old
    """

    # COMMENT: code for waiting for a file to exist copied from source link
    # DATE: February 26, 2022
    # SOURCE : https://stackoverflow.com/questions/21746750/check-and-wait-until-a-file-exists-to-read-it

    dir_path, slash = get_system()

    file_path = dir_path + slash + "email_data.txt"

    sleepTime = 3

    while True:

        # get the current time
        timeNow = time.time()

        # if it is the file we are looking for, process it. Otherwise, print an error and exit
        if os.path.exists(file_path) and os.path.isfile(file_path):
            # get the time since epoch for the last time the file was modified
            lastTimeSent = os.path.getmtime(file_path)
            if lastTimeSent > timeNow - sleepTime:
                # if the file is less than 3 seconds old, the email will be sent
                generate_email(file_path, dir_path, slash)

        else:
            popup("email_data.txt was an invalid file.")

        time.sleep(sleepTime)


if __name__ == "__main__":
    try:
        email_service()
    except KeyboardInterrupt:
        sys.exit(0)
