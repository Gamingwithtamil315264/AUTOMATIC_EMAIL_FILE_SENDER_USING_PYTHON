import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import os
import pandas as pd


df = pd.read_csv('<CSV_FILE>')


email_send = "EMAIL_ID OF THE SENDER"
email_password = "PASSWORD OF THE SENDER"


body = """MESSAGE BODY CONTENT"""

def mail(path_to_file, recipient_email):

    message = MIMEMultipart()
    message['Subject'] = "MESSAGE SUBJECT"
    message['From'] = email_send
    message['To'] = recipient_email


    body_part = MIMEText(body, 'plain')
    message.attach(body_part)


    with open(path_to_file, 'rb') as file:
        
        message.attach(MIMEApplication(file.read(), Name="FILE NAME AND EXTENSION(EX:p1.docx)"))

    return message


certificate_folder = r"PATH TO THE FILE"
path_files = os.listdir(certificate_folder)
print(path_files)

for cert_file in path_files:
    
    file_path = os.path.join(certificate_folder, cert_file)

    
    try:
        file_name_without_ext = os.path.basename(cert_file).split('.')[0]
        email_index = int(file_name_without_ext) - 1  
    except ValueError:
        print(f"Skipping file '{cert_file}' due to invalid file name format.")
        continue

    
    if 0 <= email_index < len(df):
        recipient_email = df['EMAIL ID'].iloc[email_index]
        print(f"Sending email to: {recipient_email} with attachment {cert_file}")

        
        try:
            
            message = mail(file_path, recipient_email)

            
            with smtplib.SMTP('smtp.gmail.com', 587) as server:
                server.starttls()
                server.login(email_send, email_password)
                server.sendmail(email_send, recipient_email, message.as_string())
            print(f"Email sent to {recipient_email}")
        except Exception as e:
            print(f"Failed to send email to {recipient_email}: {e}")
    else:
        print(f"Index {email_index} is out of bounds for the DataFrame.")
