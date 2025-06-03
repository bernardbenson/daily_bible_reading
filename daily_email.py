import pandas as pd
#from getpass4 import getpass
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from datetime import datetime, timedelta


# Read the Bible readings from the CSV file
bible_readings = pd.read_excel('/Users/cahabaheightschurch/bible_reading/bible_reading_2025.xlsx')

def format_reading(reading):
    # Split the reading by commas, remove any empty parts, and join each part with a new line
    parts = [part.strip() for part in reading.split(',') if part.strip()]
    formatted_reading = '\n'.join(parts)
    return formatted_reading


# Function to send the email
def send_email(subject, body, recipients, sender_name):
    from_email = "bernardvbenson@gmail.com"
    from_password = "" #put password here
    
    # Format the "From" field to include the sender's name
    formatted_from = f"{sender_name} <{from_email}>"
    
    msg = MIMEMultipart()
    msg['From'] = formatted_from
    msg['To'] = ", ".join(recipients)
    msg['Subject'] = subject
    
    msg.attach(MIMEText(body, 'plain'))
    
    try:
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(from_email, from_password)
        text = msg.as_string()
        server.sendmail(from_email, recipients, text)
        server.quit()
        print("Email sent successfully!")
    except Exception as e:
        print(f"Failed to send email: {e}")

# Get today's date and the day of the week
today = datetime.now()
day_of_week = today.strftime('%A')

# Only send email from Monday to Friday
if day_of_week in ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday']:
    # Get today's Bible reading
    reading = bible_readings[bible_readings['date'] == today.strftime('%Y-%m-%d')]
    # format reading to remove commas and separate on new line   
    if not reading.empty:
        subject = f"Daily Bible Reading for {today.strftime('%x')}"
        body = reading['reading'].values[0]
        body = format_reading(body)
        recipients = ['CahabaChurch@googlegroups.com','bernardbenson.dl@gmail.com']
        send_email(subject, body, recipients, "Bernard Benson")
    else:
        print("No reading for today.")
else:
    print("No email sent. It's the weekend!")
    

