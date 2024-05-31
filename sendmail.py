import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders


def send_email(datastudent, emailpassword, from_addr):
    to_addr = f"{datastudent[3][1].lower()}.{datastudent[2][1].lower()}@satom.ch"
    subject = 'Votre bulletin du semestre'
    body = 'Bonjour, voici votre bulletin du semestre. \n\n Cordialement \n\n Alexandre REY'

    # Create a multipart message
    msg = MIMEMultipart()
    msg['From'] = from_addr
    msg['To'] = to_addr
    msg['Subject'] = subject

    # Attach the message body
    msg.attach(MIMEText(body, 'plain'))

    # Open and attach the file
    filename = f"{datastudent[2][1]}_{datastudent[3][1]}_bulletins.pdf"
    attachment = open(filename, 'rb')

    # Create a MIMEBase object
    part = MIMEBase('application', 'octet-stream')
    part.set_payload(attachment.read())

    # Encode the attachment
    encoders.encode_base64(part)

    # Add header
    part.add_header('Content-Disposition', f'attachment; filename= {filename}')

    # Attach the attachment to the message
    msg.attach(part)

    # Close the attachment file
    attachment.close()

    # Connect to the SMTP server (e.g.,Office 365)
    server = smtplib.SMTP('smtp.office365.com', 587)
    server.starttls()

    # Login to your email account
    server.login(from_addr, emailpassword)

    # Send the email
    server.sendmail(from_addr, to_addr, msg.as_string())

    # Close the connection
    server.quit()

    print('Email sent successfully.')
