import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import os
import pandas as pd

def send_qr_codes(csv_path, qr_dir, sender_email, sender_password, smtp_server, smtp_port):
    """
    Send QR codes via email.

    Args:
        csv_path (str): Path to the CSV file containing recipient details.
        qr_dir (str): Directory containing QR code images.
        sender_email (str): Sender's email address.
        sender_password (str): Sender's email password.
        smtp_server (str): SMTP server address.
        smtp_port (int): SMTP server port.
    """
    # Read the CSV file
    df = pd.read_csv(csv_path, dtype=str)
    
    # Set up the SMTP server
    try:
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        server.login(sender_email, sender_password)
        print("Logged in to the SMTP server successfully.")
    except Exception as e:
        print(f"Error connecting to SMTP server: {e}")
        return

    # Iterate through the rows in the CSV file
    for _, row in df.iterrows():
        recipient_email = row.get('mail', '').strip()
        mobile = row.get('mobile', '').strip()

        if not recipient_email or not mobile:
            print(f"Skipping row due to missing email or mobile: {row}")
            continue

        # Find the QR code file
        qr_file_path = os.path.join(qr_dir, f"{mobile}.png")
        if not os.path.exists(qr_file_path):
            print(f"QR code not found for mobile: {mobile}")
            continue

        # Create the email
        subject = "Your QR Code"
        body = f"Dear {row.get('isim', 'Participant')},\n\nPlease find your QR code attached.\n\nBest regards,\nEvent Team"
        msg = MIMEMultipart()
        msg['From'] = sender_email
        msg['To'] = recipient_email
        msg['Subject'] = subject
        msg.attach(MIMEText(body, 'plain'))

        # Attach the QR code
        with open(qr_file_path, 'rb') as attachment:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(attachment.read())
        encoders.encode_base64(part)
        part.add_header(
            'Content-Disposition',
            f'attachment; filename={os.path.basename(qr_file_path)}',
        )
        msg.attach(part)

        # Send the email
        try:
            server.sendmail(sender_email, recipient_email, msg.as_string())
            print(f"Email sent to {recipient_email}")
        except Exception as e:
            print(f"Error sending email to {recipient_email}: {e}")

    # Close the SMTP server connection
    server.quit()
    print("All emails have been sent.")

# Example usage
if __name__ == "__main__":
    CSV_PATH = 'output/csv/yanitlar.csv'
    QR_DIR = 'output/qr'
    SENDER_EMAIL = 'your_email@example.com'
    SENDER_PASSWORD = 'your_password'
    SMTP_SERVER = 'smtp.gmail.com'
    SMTP_PORT = 587

    send_qr_codes(CSV_PATH, QR_DIR, SENDER_EMAIL, SENDER_PASSWORD, SMTP_SERVER, SMTP_PORT)