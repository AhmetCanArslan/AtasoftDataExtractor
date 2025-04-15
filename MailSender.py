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
        qr_dir (str): Directory containing the designed QR code images to be sent.
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

        # Find the designed QR code file
        qr_file_path = os.path.join(qr_dir, f"{mobile}_designed.png")
        if not os.path.exists(qr_file_path):
            print(f"Designed QR code not found for mobile: {mobile} at path: {qr_file_path}")
            continue

        # Create the email
        subject = "Your A.I. Summit Erzurum E-Ticket"

        # Format participant name
        full_name = row.get('isim', 'Participant').strip()
        if full_name and full_name != 'Participant':
            name_parts = full_name.split()
            if len(name_parts) > 1:
                last_name = name_parts[-1].upper()
                first_middle_names = [name.capitalize() for name in name_parts[:-1]]
                formatted_name = " ".join(first_middle_names) + " " + last_name
            else:
                formatted_name = name_parts[0].capitalize()
        else:
            formatted_name = 'Participant'

        body = (
            f"Sevgili {formatted_name},\n\n"
            "Zirveye katılım için hazırladığımız e-biletiniz ekte yer almaktadır.\n\n"
            "Lütfen etkinlik alanında E-Biletinizi hazır bulundurunuz.❗❗\n\n"
            "Heyecan dolu bu deneyimin bir parçası olmaya hazır olun! Sizlerle buluşmak için sabırsızlanıyoruz.\n\n"
            "Etkinlik detayları ve güncellemeler için bizi Instagram’dan takip etmeyi unutmayın:\n\n"
            "https://www.instagram.com/atauniaisummiterzurum\n\n"
            "Görüşmek üzere!\n"
            "ATASOFT Ekibi"
        )
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
        # Use formatted participant's name for the attachment filename
        attachment_filename = f"{formatted_name}.png"
        part.add_header(
            'Content-Disposition',
            f'attachment; filename="{attachment_filename}"', # Enclose filename in quotes
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
