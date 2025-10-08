import pandas as pd
from message  import create_email_body
import config
import smtplib   # its laibrary to sends mails
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import random
import time


SERVER = "smtp.gmail.com"
PORT = 587
FROM = config.EMAIL_ADDRESS
PASS = config.EMAIL_PASSWORD    #use app password not ur actual pass, becasue of (smtplib.SMTPAuthenticationError).


server = smtplib.SMTP(SERVER, PORT) # create connection with server useing(server and ports).
server.set_debuglevel(1) #optinal to monitror the detals throug the terminal
server.ehlo() #its like cheack hand phase betwwen you and the server.
server.starttls() # create encrypt connection so that any message we will send it will be encrypted.
server.login(FROM, PASS) #To login to the account useing what we give him (from,pass)


emails_list = pd.read_excel("Emails_info.xlsx")
names = emails_list["Company_Name"]
emails = emails_list["Email"]
#to count all the emails in the file
total_emails = len(emails)

pdf_path = "Mohammed_AlQahtani_v3.pdf"
with open(pdf_path, "rb") as attachment_file:
    part = MIMEBase("application", "octet-stream") #use for file contents like(pdf), for attachments
    part.set_payload(attachment_file.read())

# for encode the header of the file that we will send
encoders.encode_base64(part)
part.add_header(
    'Content-Disposition',
    f'attachment; filename={pdf_path}',#tilte for our files so that the reciver know what we sent
)

#counter to count the emails 
sent_count = 0
#LOOP throug emails 
for send in range(len(emails)):
    name = names[send]
    email = emails[send]
    if pd.notna(email) and pd.notna(name): #to check if there's empty row in excle file.
        msg = MIMEMultipart()      # use to create new meassage to detect the doublect info like (body,files,etc)
        msg["Subject"] = "Job Application"
        msg["From"] = FROM
        msg["To"] = emails[send]
        message = create_email_body(name) # call the function to replacy evrey name of the company we wanna sent, and avoid the spam from the email.
        msg.attach(MIMEText(message, "plain"))
        msg.attach(part)
        server.sendmail(FROM, [email], msg.as_string())
        sent_count += 1
        print(f"{sent_count}/ {total_emails} Email sent to {emails[send]}.")
    else:
        print("Skipping a row because email or name is empty.")
    random_wait_time = random.randint(15, 25)
    print(f"Wait for {random_wait_time} secondes to send the next email....")
    time.sleep(random_wait_time)

print("all Emails sent successfully!")
server.quit()      




