import smtplib, os
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

sendingUser='dpotter@scips.com'
sendingPassword='Kathar!na1959'
toMail='linux@davidlpotter.com'
subject="Python Test Email"

def send_email():
    try:    
        # Create message container - the correct MIME type is multipart/alternative.
        msg = MIMEMultipart('alternative')
        
        msg['From'] = sendingUser
        msg['To'] = toMail        
        msg['Subject'] = subject

        # see the code below to use template as body
        body_text = "Python says Hi this is body text of email"
        body_html = "<p><h1>Python says Hi this is body text of email</h1></p>"
        
        # Create the body of the message (a plain-text and an HTML version).
        # Record the MIME types of both parts - text/plain and text/html.
        part1 = MIMEText(body_text, 'plain')
        part2 = MIMEText(body_html, 'html')

        # Attach parts into message container.
        # According to RFC 2046, the last part of a multipart message, in this case
        # the HTML message, is best and preferred.
        msg.attach(part1)
        msg.attach(part2)
        
        # Send the message via local SMTP server.
        
        mail = smtplib.SMTP("smtp.outlook.office365.com", 587, timeout=20)

        # if tls = True                
        mail.starttls()        

        recepient = [toMail]
                    
        mail.login(sendingUser, sendingPassword)        
        mail.sendmail(sendingUser, recepient, msg.as_string())        
        mail.quit()
        
    except Exception as e:
        raise e

    
send_email()