import smtplib, os
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import formatdate

from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders

sendingUser='dpotter@scips.com'
sendingPassword='xxxxxxxx'
toMail='linux@davidlpotter.com'
cc='hdroadglide@gmail.com'  # not used yet - not sure it ever has to be used
subject="Python Test Email"

file = 'Setting up git.docx'
fileDescription= file.split('.')[0]

msg = MIMEMultipart()
msg['From'] = sendingUser
msg['To'] = toMail
msg['Cc'] = cc
msg['Date'] = formatdate(localtime = True)
msg['Subject'] = 'Sending emails from Python'


def send_email():
    try:    
        # Create message container - the correct MIME type is multipart/alternative.

        fp = open(file, 'rb')
        part = MIMEBase('application','vnd.ms-excel')
        part.set_payload(fp.read())
        fp.close()

        encoders.encode_base64(part)
        part.add_header('Content-Disposition', 'attachment', filename=fileDescription)
        msg.attach(part)
        # see the code below to use template as body
        body_text = "Python says Hi this is the body for the text of email"
        body_html = "<p><h1>Python says Hi this is the body for the text of email</h1></p>"
        
        # Create the body of the message (a plain-text and an HTML version).
        # Record the MIME types of both parts - text/plain and text/html.
        part1 = MIMEText(body_text, 'plain')
        part2 = MIMEText(body_html, 'html')

        # Attach parts into message container.
        # According to RFC 2046, the last part of a multipart message, in this case
        # the HTML message, is best and preferred.
        # msg.attach(part)
        msg.attach(part1)
        msg.attach(part2)
        
        # Send the message via local SMTP server.
        
        
        mail = smtplib.SMTP("smtp.outlook.office365.com", 587, timeout=20)
        
        # if tls = True Microsoft does not support TLS this way - so leave it commented                 
        mail.starttls()        

        recepient = [toMail]
                    
        mail.login(sendingUser, sendingPassword)        
        mail.sendmail(sendingUser, recepient, msg.as_string())  
        
        mail.quit()
        
    except Exception as e:
        raise e

    
send_email()