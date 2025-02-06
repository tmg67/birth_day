import pandas as pd
import datetime
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# Load DataFrame
try:
    df = pd.read_excel("birthday_tracker.xlsx")
    print(df)
except FileNotFoundError:
    print("Error: 'Birthday Tracker.xlsx' not found.")
    exit()

# Get today's date
today = datetime.datetime.now()
today_month = today.month
# today_day = today.day

# Iterate through rows in DataFrame
for index, row in df.iterrows():
    birthday = pd.to_datetime(row['BirthDay'])  # Ensure BirthDay is in datetime format
    birthday_month = birthday.month
    # birthday_day = birthday.day

    if birthday_month == today_month: #and birthday_day == today_day:
        name = row['Name']
        email = row['Email']  # Assuming 'Email' column exists
        print(f"Today is {name}'s birthday!")
        
        # Send Email
        def send_email(email, name):
            sender_email = "avipshashrestha75@gmail.com"
            password = "kldh uvgm jkbe czhr"  # Use environment variables for security
            receiver_email = email
            
            message = MIMEMultipart()
            message["From"] = sender_email
            message["To"] = receiver_email
            message["Subject"] = "Happy Birthday!"
            
            body = f"Dear {name},\n\nWishing you a wonderful birthday filled with love and joy!\n\nBest Regards,\nYour Python Classmate"
            message.attach(MIMEText(body, "plain"))

            try:
                with smtplib.SMTP("smtp.gmail.com", 587) as server:
                    server.starttls()
                    server.login(sender_email, password)
                    server.sendmail(sender_email, receiver_email, message.as_string())
                print(f"Birthday email sent successfully to {name}!")
            except smtplib.SMTPException as e:
                print(f"Error sending email to {email}: {e}")

        send_email(email, name)