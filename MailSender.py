import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import os
import pandas as pd
from dotenv import load_dotenv # Import dotenv

# Load environment variables from .env file
load_dotenv()

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
    for index, row in df.iterrows(): # Use index for better logging
        recipient_email = row.get('mail', '').strip()
        mobile = row.get('mobile', '').strip() # Ensure mobile is stripped

        # Skip if email or mobile is missing/empty
        if not recipient_email or not mobile:
            print(f"Skipping row {index + 2} due to missing email ('{recipient_email}') or mobile ('{mobile}')")
            continue

        # Construct the expected designed QR code filename
        expected_filename = f"{mobile}_designed.png"
        qr_file_path = os.path.join(qr_dir, expected_filename)

        # Check if the designed QR code file exists
        if not os.path.exists(qr_file_path):
            print(f"Warning: Designed QR code not found for mobile '{mobile}'. Looked for '{expected_filename}' in '{qr_dir}'. Skipping email for row {index + 2}.")
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

        # Use formatted participant's name directly for the attachment filename
        attachment_filename = f"{formatted_name}.png"

        # Correctly encode the filename parameter for the header
        part.add_header(
            'Content-Disposition',
            'attachment',
            filename=attachment_filename
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

# --- Main block for debugging ---
if __name__ == "__main__":
    print("--- Running MailSender in Debug Mode ---")

    # --- Configuration (Load from .env) ---
    # Use the designed QR directory
    DESIGNED_QR_DIR = os.path.join('output', 'designed_qr')
    SENDER_EMAIL = os.getenv('SENDER_EMAIL')
    SENDER_PASSWORD = os.getenv('SENDER_PASSWORD')
    SMTP_SERVER = os.getenv('SMTP_SERVER')
    SMTP_PORT_STR = os.getenv('SMTP_PORT')

    # --- Default Test Recipient ---
    DEFAULT_RECIPIENT_EMAIL = "yildirimciomer237@gmail.com"
    DEFAULT_MOBILE = "5380718571" # Used to find the QR code
    DEFAULT_NAME = "Test User" # Placeholder name for the email

    # --- Basic Validation ---
    if not all([SENDER_EMAIL, SENDER_PASSWORD, SMTP_SERVER, SMTP_PORT_STR]):
        print("Error: Missing one or more email configuration variables in .env file.")
        print("Please ensure SENDER_EMAIL, SENDER_PASSWORD, SMTP_SERVER, SMTP_PORT are set.")
        exit(1)

    try:
        SMTP_PORT = int(SMTP_PORT_STR)
    except ValueError:
        print(f"Error: Invalid SMTP_PORT value '{SMTP_PORT_STR}'. Must be an integer.")
        exit(1)

    if not os.path.isdir(DESIGNED_QR_DIR):
        print(f"Error: Designed QR directory '{DESIGNED_QR_DIR}' not found.")
        exit(1)

    # --- Construct QR file path for the test user ---
    expected_filename = f"{DEFAULT_MOBILE}_designed.png"
    qr_file_path = os.path.join(DESIGNED_QR_DIR, expected_filename)

    if not os.path.exists(qr_file_path):
        print(f"Error: Designed QR code for test mobile '{DEFAULT_MOBILE}' not found.")
        print(f"Looked for: {qr_file_path}")
        exit(1)

    print(f"Attempting to send test email to: {DEFAULT_RECIPIENT_EMAIL}")
    print(f"Using QR code: {qr_file_path}")

    # --- Set up SMTP Connection ---
    try:
        server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
        server.starttls()
        server.login(SENDER_EMAIL, SENDER_PASSWORD)
        print("Logged in to SMTP server successfully for debug send.")
    except Exception as e:
        print(f"Error connecting or logging in to SMTP server: {e}")
        exit(1)

    # --- Create Email Message ---
    subject = "TEST - Your A.I. Summit Erzurum E-Ticket"
    formatted_name = DEFAULT_NAME # Use the default name for the test

    body = (
        f"Sevgili {formatted_name},\n\n"
        "--- THIS IS A TEST EMAIL ---\n\n"
        "Zirveye katılım için hazırladığımız e-biletiniz ekte yer almaktadır.\n\n"
        "Lütfen etkinlik alanında E-Biletinizi hazır bulundurunuz.❗❗\n\n"
        "Heyecan dolu bu deneyimin bir parçası olmaya hazır olun! Sizlerle buluşmak için sabırsızlanıyoruz.\n\n"
        "Etkinlik detayları ve güncellemeler için bizi Instagram’dan takip etmeyi unutmayın:\n\n"
        "https://www.instagram.com/atauniaisummiterzurum\n\n"
        "Görüşmek üzere!\n"
        "ATASOFT Ekibi"
    )
    msg = MIMEMultipart()
    msg['From'] = SENDER_EMAIL
    msg['To'] = DEFAULT_RECIPIENT_EMAIL
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'plain'))

    # --- Attach QR Code ---
    try:
        with open(qr_file_path, 'rb') as attachment:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(attachment.read())
        encoders.encode_base64(part)
        # Use the test name directly for the attachment filename
        attachment_filename = f"{formatted_name}_TestQR.png"

        # Apply the same header encoding method for consistency
        part.add_header(
            'Content-Disposition',
            'attachment',
            filename=attachment_filename
        )
        msg.attach(part)
        print(f"Successfully attached QR code: {qr_file_path}")
    except Exception as e:
        print(f"Error attaching QR code file '{qr_file_path}': {e}")
        server.quit()
        exit(1)

    # --- Send Email ---
    try:
        server.sendmail(SENDER_EMAIL, DEFAULT_RECIPIENT_EMAIL, msg.as_string())
        print(f"Test email successfully sent to {DEFAULT_RECIPIENT_EMAIL}")
    except Exception as e:
        print(f"Error sending test email to {DEFAULT_RECIPIENT_EMAIL}: {e}")
    finally:
        # --- Close Connection ---
        server.quit()
        print("SMTP server connection closed.")

    print("--- MailSender Debug Mode Finished ---")
