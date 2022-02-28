# OSU CS 361 - Software Engineering I
## Microservice : Email Service (Email_Service.py)

### Overview of this Microservice
This email service uses OAuth 2.0 to send emails from a gmail account to a recipient<br>
<br>
This code looks for a text file in the location "C:[USER]/Desktop/email_service_data/email_data.txt"<br>
<br>
At this point, if the email_data.txt file is found it is deleted (even if it was not successfully read) and a fail.txt or success.txt file is added to the folder in its place.<br>
<br>
The first time you use the service it will open an internet browser and verify that you want to authenticate the email being used<br>
<br>
### email_data.txt structure
Line [0] - Sender's email, required<br>
Line [1] - Recipient's email, required<br>
Line [2] - Subject line for the email, optional<br>
Line [3] - full file path for any attachment to be added to the email, optional<br>
Line [4:] - message contents, optional, can be on multiple lines<br>

Any lines marked as optional must be left as a blank line.
