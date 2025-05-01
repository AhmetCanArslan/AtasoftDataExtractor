import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import os
import pandas as pd
from dotenv import load_dotenv
from datetime import datetime
import csv

# Load environment variables from .env file
load_dotenv()

# --- Constants ---
SENT_EMAIL_LOG_FILE = 'logs/sent_emails.csv'
EMAIL_ERROR_LOG_FILE = 'logs/email_errors.csv'

# --- Utility Functions ---

def _ensure_log_dir_exists(log_file):
    """Ensures the directory for the log file exists."""
    log_dir = os.path.dirname(log_file)
    if log_dir and not os.path.exists(log_dir):
        try:
            os.makedirs(log_dir)
            print(f"Created log directory: {log_dir}")
        except OSError as e:
            print(f"Warning: Could not create log directory '{log_dir}': {e}")
            return False
    return True

def get_sent_emails(log_file=SENT_EMAIL_LOG_FILE):
    """
    Get a set of email addresses that have already been sent QR codes.
    
    Args:
        log_file (str): Path to the CSV file tracking sent emails.
        
    Returns:
        set: A set of email addresses that have already received QR codes.
    """
    if not _ensure_log_dir_exists(log_file):
        return set()
    
    if not os.path.exists(log_file):
        try:
            with open(log_file, 'w', encoding='utf-8') as f:
                f.write("email,mobile,sent_date\n")
            return set()
        except IOError as e:
            print(f"Warning: Could not create sent emails log file '{log_file}': {e}")
            return set()
    
    try:
        sent_df = pd.read_csv(log_file, dtype=str)
        if 'email' in sent_df.columns:
            return set(sent_df['email'].str.lower().dropna())
        else:
            print(f"Warning: 'email' column not found in {log_file}. Assuming no emails sent.")
            return set()
    except Exception as e:
        print(f"Warning: Error reading sent emails log '{log_file}': {e}")
        return set()

def record_sent_email(email, mobile, log_file=SENT_EMAIL_LOG_FILE):
    """
    Record that an email has been sent to a specific address.
    
    Args:
        email (str): Email address that received the QR code.
        mobile (str): Mobile number associated with the email.
        log_file (str): Path to the CSV file tracking sent emails.
    """
    if not _ensure_log_dir_exists(log_file):
        return

    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    try:
        file_exists = os.path.exists(log_file)
        with open(log_file, 'a', encoding='utf-8', newline='') as f:
            if not file_exists or os.path.getsize(log_file) == 0:
                f.write("email,mobile,sent_date\n")
            f.write(f"{email},{mobile},{now}\n")
    except IOError as e:
        print(f"Warning: Could not record sent email to log '{log_file}': {e}")

def log_email_error(email, mobile, reason, log_file=EMAIL_ERROR_LOG_FILE):
    """
    Logs an error or skip reason for a specific email address using the csv module.

    Args:
        email (str): Email address that failed or was skipped.
        mobile (str): Mobile number associated with the email (can be None or 'N/A').
        reason (str): Description of the error or skip reason.
        log_file (str): Path to the CSV file for logging errors.
    """
    if not _ensure_log_dir_exists(log_file):
        print(f"Error: Cannot log email error because log directory cannot be created/accessed for '{log_file}'.")
        return

    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    email_str = str(email) if email else 'N/A'
    mobile_str = str(mobile) if mobile else 'N/A'
    reason_str = str(reason)

    try:
        file_exists = os.path.exists(log_file)
        with open(log_file, 'a', encoding='utf-8', newline='') as f:
            writer = csv.writer(f)
            if not file_exists or os.path.getsize(log_file) == 0:
                writer.writerow(["email", "mobile", "reason", "timestamp"])
            writer.writerow([email_str, mobile_str, reason_str, now])
    except (IOError, csv.Error) as e:
        print(f"CRITICAL ERROR: Could not write to email error log '{log_file}'. Error: {type(e).__name__} - {e}")
        print(f"Data attempted to log: Email='{email_str}', Mobile='{mobile_str}', Reason='{reason_str}', Timestamp='{now}'")

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
    try:
        df = pd.read_csv(csv_path, dtype=str)
    except FileNotFoundError:
        print(f"Error: CSV file not found at '{csv_path}'")
        log_email_error('N/A', 'N/A', f'Input CSV not found: {csv_path}')
        return
    except Exception as e:
        print(f"Error reading CSV file '{csv_path}': {e}")
        log_email_error('N/A', 'N/A', f'Error reading CSV: {e}')
        return

    sent_emails = get_sent_emails()
    print(f"Found {len(sent_emails)} previously sent emails in log file.")

    sent_count = 0
    skipped_already_sent = 0
    skipped_missing_data = 0
    skipped_missing_qr = 0
    failed_send_count = 0

    server = None
    try:
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        server.login(sender_email, sender_password)
        print("Logged in to the SMTP server successfully.")
    except Exception as e:
        print(f"Error connecting to SMTP server: {e}")
        log_email_error('N/A', 'N/A', f'SMTP Connection/Login Failed: {e}')
        return

    for index, row in df.iterrows():
        recipient_email = row.get('mail', '').strip()
        mobile = row.get('mobile', '').strip()

        if not recipient_email or not mobile:
            skipped_missing_data += 1
            log_email_error(recipient_email or 'N/A', mobile or 'N/A', 'Missing email or mobile in CSV row')
            continue

        if recipient_email.lower() in sent_emails:
            skipped_already_sent += 1
            continue

        expected_filename = f"{mobile}_designed.png"
        qr_file_path = os.path.join(qr_dir, expected_filename)

        if not os.path.exists(qr_file_path):
            skipped_missing_qr += 1
            log_email_error(recipient_email, mobile, f'Missing designed QR file: {expected_filename}')
            continue

        subject = "Your A.I. Summit Erzurum E-Ticket"
        full_name = row.get('isim', 'Participant').strip()
        formatted_name = 'Participant'
        if full_name and full_name != 'Participant':
            name_parts = full_name.split()
            if len(name_parts) > 1:
                last_name = name_parts[-1].upper()
                first_middle_names = [name.capitalize() for name in name_parts[:-1]]
                formatted_name = " ".join(first_middle_names) + " " + last_name
            elif len(name_parts) == 1:
                formatted_name = name_parts[0].capitalize()

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

        try:
            with open(qr_file_path, 'rb') as attachment:
                part = MIMEBase('application', 'octet-stream')
                part.set_payload(attachment.read())
            encoders.encode_base64(part)
            part.add_header(
                'Content-Disposition',
                f'attachment; filename={os.path.basename(qr_file_path)}',
            )
            msg.attach(part)
        except Exception as e:
            print(f"Error attaching QR code {qr_file_path} for {recipient_email}: {e}")
            failed_send_count += 1
            log_email_error(recipient_email, mobile, f'Error attaching QR file: {e}')
            continue

        try:
            server.sendmail(sender_email, recipient_email, msg.as_string())
            print(f"Email sent to {recipient_email}")
            record_sent_email(recipient_email, mobile)
            sent_count += 1
        except Exception as e:
            print(f"Error sending email to {recipient_email}: {e}")
            failed_send_count += 1
            log_email_error(recipient_email, mobile, f'Sending failed: {e}')

    if server:
        server.quit()
        print("SMTP server connection closed.")

    print("\nEmail Sending Summary:")
    print(f" - Successfully sent: {sent_count}")
    print(f" - Skipped (already sent): {skipped_already_sent}")
    print(f" - Skipped (missing data): {skipped_missing_data}")
    print(f" - Skipped (missing QR): {skipped_missing_qr}")
    print(f" - Failed to send/attach: {failed_send_count}")
    print(f" - Total processed: {len(df)}")
    print(f"Note: Errors/skips logged to '{EMAIL_ERROR_LOG_FILE}'")
    print("All emails have been processed.")

# --- Main block for debugging ---
if __name__ == "__main__":
    print("--- Running MailSender in Debug Mode ---")

    DESIGNED_QR_DIR = os.path.join('output', 'designed_qr')
    SENDER_EMAIL = os.getenv('SENDER_EMAIL')
    SENDER_PASSWORD = os.getenv('SENDER_PASSWORD')
    SMTP_SERVER = os.getenv('SMTP_SERVER')
    SMTP_PORT_STR = os.getenv('SMTP_PORT')

    DEFAULT_RECIPIENT_EMAIL = "yildirimciomer237@gmail.com"
    DEFAULT_MOBILE = "5380718571"
    DEFAULT_NAME = "Test User"

    if not all([SENDER_EMAIL, SENDER_PASSWORD, SMTP_SERVER, SMTP_PORT_STR]):
        print("Error: Missing one or more email configuration variables in .env file.")
        log_email_error(DEFAULT_RECIPIENT_EMAIL, DEFAULT_MOBILE, 'Missing .env configuration for debug send')
        exit(1)

    try:
        SMTP_PORT = int(SMTP_PORT_STR)
    except ValueError:
        print(f"Error: Invalid SMTP_PORT value '{SMTP_PORT_STR}'. Must be an integer.")
        log_email_error(DEFAULT_RECIPIENT_EMAIL, DEFAULT_MOBILE, f'Invalid SMTP_PORT in .env: {SMTP_PORT_STR}')
        exit(1)

    if not os.path.isdir(DESIGNED_QR_DIR):
        print(f"Error: Designed QR directory '{DESIGNED_QR_DIR}' not found.")
        log_email_error(DEFAULT_RECIPIENT_EMAIL, DEFAULT_MOBILE, f'Designed QR directory not found: {DESIGNED_QR_DIR}')
        exit(1)

    expected_filename = f"{DEFAULT_MOBILE}_designed.png"
    qr_file_path = os.path.join(DESIGNED_QR_DIR, expected_filename)

    if not os.path.exists(qr_file_path):
        print(f"Error: Designed QR code for test mobile '{DEFAULT_MOBILE}' not found.")
        log_email_error(DEFAULT_RECIPIENT_EMAIL, DEFAULT_MOBILE, f'Missing designed QR for test: {expected_filename}')
        exit(1)

    print(f"Attempting to send test email to: {DEFAULT_RECIPIENT_EMAIL}")
    print(f"Using QR code: {qr_file_path}")

    server = None
    try:
        server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
        server.starttls()
        server.login(SENDER_EMAIL, SENDER_PASSWORD)
        print("Logged in to SMTP server successfully for debug send.")
    except Exception as e:
        print(f"Error connecting or logging in to SMTP server: {e}")
        log_email_error(DEFAULT_RECIPIENT_EMAIL, DEFAULT_MOBILE, f'SMTP Connection/Login Failed (Debug): {e}')
        exit(1)

    subject = "TEST - Your A.I. Summit Erzurum E-Ticket"
    body = (
        f"Sevgili {DEFAULT_NAME},\n\n"
        "Bu bir test e-postasıdır.\n\n"
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

    try:
        with open(qr_file_path, 'rb') as attachment:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(attachment.read())
        encoders.encode_base64(part)
        part.add_header(
            'Content-Disposition',
            f'attachment; filename={os.path.basename(qr_file_path)}',
        )
        msg.attach(part)
        print("Test QR code attached successfully.")
    except Exception as e:
        print(f"Error attaching test QR code {qr_file_path}: {e}")
        log_email_error(DEFAULT_RECIPIENT_EMAIL, DEFAULT_MOBILE, f'Error attaching test QR: {e}')
        if server: server.quit()
        exit(1)

    try:
        server.sendmail(SENDER_EMAIL, DEFAULT_RECIPIENT_EMAIL, msg.as_string())
        print(f"Test email successfully sent to {DEFAULT_RECIPIENT_EMAIL}")
    except Exception as e:
        print(f"Error sending test email to {DEFAULT_RECIPIENT_EMAIL}: {e}")
        log_email_error(DEFAULT_RECIPIENT_EMAIL, DEFAULT_MOBILE, f'Test email sending failed: {e}')
    finally:
        if server:
            server.quit()
            print("SMTP server connection closed.")

    print(f"Note: Errors/skips logged to '{EMAIL_ERROR_LOG_FILE}'")
    print("--- MailSender Debug Mode Finished ---")
