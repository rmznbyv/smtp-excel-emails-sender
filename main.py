import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
import pandas as pd
import time

def send_mail(to_address, subject, body, attachment_paths, school):
    smtp_server = 'smtp-mail.outlook.com'
    smtp_port = 587
    sender_email = 'yourmail'
    sender_password = 'yourpass'

    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = ", ".join(to_address)
    msg['Subject'] = subject

    html_body = f"""
    Dear {school} HR Team,<br><br>

    I trust this email finds you well. My name is Vali Aliyev, and I am writing to express my interest in potential teaching opportunities at your school. As an experienced <b>IGCSE and A, AS-Level chemistry teacher</b> with expertise in Python and Excel training, I am eager about the prospect of contributing to your esteemed institution.<br><br>

    I have recently updated my CV and crafted a comprehensive cover letter, both of which I have attached for your reference. In my endeavor to ensure the most effective application process, I am reaching out to you directly. <b>If this email is not directed to the Human Resources department</b>, I kindly request your assistance in forwarding my application to the appropriate contact or providing the correct HR contact information.<br><br>

    To facilitate a smooth application process, <b>I would greatly appreciate it if you could confirm</b> the receipt of this email and provide feedback on the status of my application. Your guidance on the next steps or any additional information required would be invaluable.<br><br>

    Thank you for considering my application. I look forward to the opportunity to contribute to the academic excellence at your school. I also trust that you will include my curriculum vitae in your database for future job openings.<br><br>

    Warm regards,<br><br>

    Mr. Vali Aliyev<br><br>
    """

    msg.attach(MIMEText(html_body, 'html'))

    for attachment_path in attachment_paths:
        try:
            with open(attachment_path, 'rb') as file:
                part = MIMEApplication(file.read())
                part.add_header('Content-Disposition', 'attachment', filename=attachment_path.split('/')[-1])
                msg.attach(part)
        except FileNotFoundError:
            print(f"Warning: File '{attachment_path}' not found. Skipping this attachment.")


    with smtplib.SMTP(smtp_server, smtp_port) as server:
        server.starttls()
        server.login(sender_email, sender_password)
        server.sendmail(sender_email, to_address, msg.as_string())

excel_file = "1.xlsx"
df = pd.read_excel(excel_file)

for index, row in df.iterrows():
    school_name = row['School Name']
    emails = [email.strip() for email in row['Emails'].split(';')]
    subject = "Job Application for Chemistry Teacher Position"
    body = "" 
    attachment_paths = ['CV_Vali_Aliyev_Chemistry_STEAM_teacher_Python_Trainer.pdf', 'Cover Letter.pdf']
    print(emails, subject, body, attachment_paths, school_name)
    send_mail(school_name + " " + emails)
    time.sleep(3)
