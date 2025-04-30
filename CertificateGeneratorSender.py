import os
import smtplib
import sys
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from PIL import Image, ImageDraw, ImageFont
import firebase_admin
from firebase_admin import credentials, firestore
from dotenv import load_dotenv
from tqdm import tqdm

# --- Copied Utility Functions ---

def create_directory_if_not_exists(directory_path):
    """Create a directory if it does not already exist."""
    if not os.path.exists(directory_path):
        os.makedirs(directory_path)
        print(f"Created directory: {directory_path}")

# --- Copied Firebase Functions ---

SERVICE_ACCOUNT_KEY_PATH = 'qr-deneme.json' # Keep config here or load from env
COLLECTION_NAME = 'users'

def initialize_firebase():
    """Initializes the Firebase Admin SDK."""
    if not os.path.exists(SERVICE_ACCOUNT_KEY_PATH):
        print(f"Error: Service account key file not found at '{SERVICE_ACCOUNT_KEY_PATH}'")
        sys.exit(1)
    try:
        # Check if Firebase app is already initialized
        if not firebase_admin._apps:
             cred = credentials.Certificate(SERVICE_ACCOUNT_KEY_PATH)
             firebase_admin.initialize_app(cred)
             print("Firebase Admin SDK initialized successfully.")
        else:
             print("Firebase Admin SDK already initialized.")
        return firestore.client()
    except Exception as e:
        print(f"Error initializing Firebase: {e}")
        sys.exit(1)

def get_attendees_from_firebase(db):
    """Fetches users from Firestore where Counter > 0."""
    users_ref = db.collection(COLLECTION_NAME)
    attendees = []
    try:
        query = users_ref.where(filter=firestore.FieldFilter('Counter', '>', 0)).stream()
        print("Fetching attendees from Firebase (Counter > 0)...")
        docs = list(query)
        if not docs:
            print("No attendees found with Counter > 0.")
            return []

        for doc in tqdm(docs, desc="Processing attendees"):
            data = doc.to_dict()
            attendees.append({
                'isim': data.get('Ad-Soyad', 'Unknown Name'),
                'mail': data.get('Eposta', ''),
                'mobile': data.get('Telefon numaranız', '')
            })
        print(f"Found {len(attendees)} attendees.")
        return attendees
    except Exception as e:
        print(f"Error fetching attendees from Firestore: {e}")
        return []

# --- Copied Certificate Generation Function ---

def generate_certificate(name, mobile, template_path, output_dir, font_path="arial.ttf", font_size=100, text_color=(0, 0, 0)):
    """Generates a certificate by drawing a name onto a template image."""
    if not name or not mobile:
        print(f"Warning: Skipping certificate generation due to missing name ('{name}') or mobile ('{mobile}')")
        return False
    try:
        template = Image.open(template_path).convert("RGB")
        template_width, template_height = template.size
        try:
            font = ImageFont.truetype(font_path, font_size)
        except IOError:
            print(f"Error: Font file not found at '{font_path}'.")
            try:
                font = ImageFont.load_default()
                print("Warning: Using default PIL font.")
            except Exception as font_e:
                 print(f"Error: Could not load default font: {font_e}. Cannot generate certificate for {name}.")
                 return False

        draw = ImageDraw.Draw(template)
        formatted_name = ' '.join(part.capitalize() for part in name.split())
        Y_POSITION = template_height * 0.45
        try:
            bbox = draw.textbbox((0, 0), formatted_name, font=font)
            text_width = bbox[2] - bbox[0]
            x_position = (template_width - text_width) / 2
            y_position = Y_POSITION
        except AttributeError:
             text_width, text_height = draw.textsize(formatted_name, font=font)
             x_position = (template_width - text_width) / 2
             y_position = Y_POSITION

        draw.text((x_position, y_position), formatted_name, fill=text_color, font=font)
        output_path = os.path.join(output_dir, f"{mobile}.png")
        template.save(output_path)
        return True
    except FileNotFoundError:
        print(f"Error: Template image not found at '{template_path}'")
        return False
    except Exception as e:
        print(f"Error generating certificate for {name} ({mobile}): {e}")
        return False

# --- Copied Certificate Sending Function ---

def send_certificates(attendees, certificate_dir, sender_email, sender_password, smtp_server, smtp_port):
    """Sends attendance certificates via email to attendees."""
    if not attendees:
        print("No attendees provided to send certificates.")
        return
    try:
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        server.login(sender_email, sender_password)
        print("Logged in to the SMTP server successfully for sending certificates.")
    except Exception as e:
        print(f"Error connecting to SMTP server for certificates: {e}")
        return

    failed_email_count = 0
    sent_count = 0
    for attendee in tqdm(attendees, desc="Sending Certificates"):
        recipient_email = attendee.get('mail', '').strip()
        mobile = attendee.get('mobile', '').strip()
        full_name = attendee.get('isim', 'Participant').strip()

        if not recipient_email or not mobile:
            print(f"Skipping certificate email for '{full_name}' due to missing email ('{recipient_email}') or mobile ('{mobile}')")
            continue

        certificate_filename = f"{mobile}.png"
        certificate_file_path = os.path.join(certificate_dir, certificate_filename)

        if not os.path.exists(certificate_file_path):
            print(f"Warning: Certificate file not found for mobile '{mobile}'. Looked for '{certificate_filename}' in '{certificate_dir}'. Skipping email for '{full_name}'.")
            continue

        subject = "Your A.I. Summit Erzurum Attendance Certificate"
        if full_name and full_name != 'Participant':
             name_parts = full_name.split()
             if len(name_parts) > 1:
                 last_name = name_parts[-1].upper()
                 first_middle_names = [name.capitalize() for name in name_parts[:-1]]
                 formatted_name_body = " ".join(first_middle_names) + " " + last_name
             else:
                 formatted_name_body = name_parts[0].capitalize()
        else:
             formatted_name_body = 'Participant'

        body = (
            f"Sevgili {formatted_name_body},\n\n"
            "A.I. Summit Erzurum etkinliğimize katılımınız için teşekkür ederiz!\n\n"
            "Katılım belgeniz ekte yer almaktadır.\n\n"
            "Başka etkinliklerde tekrar görüşmek dileğiyle!\n\n"
            "Etkinlik detayları ve güncellemeler için bizi Instagram’dan takip etmeyi unutmayın:\n\n"
            "https://www.instagram.com/atauniaisummiterzurum\n\n"
            "Saygılarımızla,\n"
            "ATASOFT Ekibi"
        )
        msg = MIMEMultipart()
        msg['From'] = sender_email
        msg['To'] = recipient_email
        msg['Subject'] = subject
        msg.attach(MIMEText(body, 'plain', 'utf-8'))

        try:
            with open(certificate_file_path, 'rb') as attachment:
                part = MIMEBase('application', 'octet-stream')
                part.set_payload(attachment.read())
            encoders.encode_base64(part)
            attachment_filename = f"AI_Summit_Erzurum_Certificate_{formatted_name_body.replace(' ', '_')}.png"
            part.add_header('Content-Disposition', 'attachment', filename=attachment_filename)
            msg.attach(part)
        except Exception as attach_e:
             print(f"Error attaching certificate file '{certificate_file_path}' for {recipient_email}: {attach_e}")
             failed_email_count += 1
             continue

        try:
            server.send_message(msg)
            sent_count += 1
        except Exception as e:
            print(f"Error sending certificate email to {recipient_email}: {type(e).__name__} - {str(e)}")
            failed_email_count += 1

    print(f"\nCertificate Email Sending Summary:")
    print(f" - Successfully sent: {sent_count}")
    print(f" - Failed attempts: {failed_email_count}")
    print(f" - Total attendees processed: {len(attendees)}")
    server.quit()
    print("Finished sending certificate emails.")

# --- Main Execution Block ---
if __name__ == "__main__":
    print("--- Starting Certificate Generation and Sending Process ---")
    load_dotenv()

    # --- Configuration ---
    CERTIFICATES_OUTPUT_DIR = os.path.join('output', 'certificates')
    TEMPLATE_IMAGE_PATH = "tasarim.jpg"
    FONT_FILE = "arial.ttf" # Ensure this font file is in the root directory

    SENDER_EMAIL = os.getenv('SENDER_EMAIL')
    SENDER_PASSWORD = os.getenv('SENDER_PASSWORD')
    SMTP_SERVER = os.getenv('SMTP_SERVER')
    SMTP_PORT_STR = os.getenv('SMTP_PORT')

    # --- Basic Validation ---
    if not all([SENDER_EMAIL, SENDER_PASSWORD, SMTP_SERVER, SMTP_PORT_STR]):
        print("Error: Missing one or more email configuration variables in .env file.")
        print("Please ensure SENDER_EMAIL, SENDER_PASSWORD, SMTP_SERVER, SMTP_PORT are set.")
        sys.exit(1)
    try:
        SMTP_PORT = int(SMTP_PORT_STR)
    except ValueError:
        print(f"Error: Invalid SMTP_PORT value '{SMTP_PORT_STR}'. Must be an integer.")
        sys.exit(1)

    if not os.path.exists(TEMPLATE_IMAGE_PATH):
        print(f"Error: Template image '{TEMPLATE_IMAGE_PATH}' not found. Cannot generate certificates.")
        sys.exit(1)

    if not os.path.exists(FONT_FILE):
        print(f"WARNING: Font file '{FONT_FILE}' not found in project root. Certificate generation might fail or use default font.")
        # Allow continuing with default font, generate_certificate handles fallback

    # --- Create Output Directory ---
    create_directory_if_not_exists(CERTIFICATES_OUTPUT_DIR)

    # --- Initialize Firebase ---
    db_client = initialize_firebase()
    if not db_client:
        sys.exit(1) # Exit if Firebase initialization failed

    # --- Fetch Attendees ---
    attendees = get_attendees_from_firebase(db_client)

    # --- Generate Certificates ---
    if attendees:
        print(f"\nGenerating {len(attendees)} certificates...")
        generated_count = 0
        for attendee in tqdm(attendees, desc="Generating Certificates"):
            success = generate_certificate(
                name=attendee['isim'],
                mobile=attendee['mobile'],
                template_path=TEMPLATE_IMAGE_PATH,
                output_dir=CERTIFICATES_OUTPUT_DIR,
                font_path=FONT_FILE
                # font_size and text_color use defaults from function definition
            )
            if success:
                generated_count += 1
        print(f"Finished generating certificates. {generated_count} successfully created.")

        # --- Ask before Sending Certificates ---
        if generated_count > 0:
            send_choice = input(f"Successfully generated {generated_count} certificates. Do you want to send them via email? (yes/no): ").strip().lower()
            if send_choice == 'yes':
                print("\n--- Starting Certificate Email Sending Process ---")
                send_certificates(
                    attendees, CERTIFICATES_OUTPUT_DIR, SENDER_EMAIL,
                    SENDER_PASSWORD, SMTP_SERVER, SMTP_PORT
                )
                # Message "Finished sending certificate emails." is printed inside send_certificates
            else:
                print("Email sending skipped by user.")
        else:
            print("No certificates were generated, skipping email sending step.")
    else:
        # Message already printed by get_attendees_from_firebase if no attendees found
        pass

    print("\n--- Certificate Generation and Sending Process Finished ---")

