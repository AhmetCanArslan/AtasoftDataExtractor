import os
import pandas as pd
import uuid
import qrcode
import re
import openpyxl
from openpyxl import Workbook
from openpyxl.drawing.image import Image as OpenpyxlImage
import subprocess # Added for running the sync script
import sys # Added for getting python executable path
from FileOperations import create_directory_if_not_exists, clean_phone_number # Import clean_phone_number
from QRDesign import overlay_qr_on_template # Removed generate_certificate
from MailSender import send_qr_codes # Removed send_certificates
# Removed FirebaseSync imports related to fetching attendees
from dotenv import load_dotenv
# Removed tqdm import if only used for certificates

# Load environment variables from .env file
load_dotenv()

# --- Configuration ---
# EXCEL_FILE_PATH = os.getenv('EXCEL_FILE_PATH') # Removed: Path will be determined dynamically
INPUT_DIR = 'input' # Define the input directory
PHONE_COLUMN_NAME = os.getenv('PHONE_COLUMN_NAME')
UUID_COLUMN_NAME = os.getenv('UUID_COLUMN_NAME')
COUNTER_COLUMN_NAME = os.getenv('COUNTER_COLUMN_NAME')
CSV_OUTPUT_DIR = os.getenv('CSV_OUTPUT_DIR')
QR_OUTPUT_DIR = os.getenv('QR_OUTPUT_DIR')
# EXCEL_OUTPUT_PATH = os.getenv('EXCEL_OUTPUT_PATH') # Removed: Path will be generated dynamically
FIREBASE_SYNC_SCRIPT = os.getenv('FIREBASE_SYNC_SCRIPT')

# Email configuration
SENDER_EMAIL = os.getenv('SENDER_EMAIL')
SENDER_PASSWORD = os.getenv('SENDER_PASSWORD')
SMTP_SERVER = os.getenv('SMTP_SERVER')
SMTP_PORT = int(os.getenv('SMTP_PORT'))

NAMESPACE = uuid.NAMESPACE_DNS  

# --- Utility Functions ---

def generate_uuid_from_phone(phone_number_str):
    if pd.isna(phone_number_str) or not str(phone_number_str).strip():
        return None
    return str(uuid.uuid5(NAMESPACE, str(phone_number_str)))

# --- Excel Processing & CSV Generation ---
def process_excel(file_path, phone_col, uuid_col, counter_col):
    try:
        if not os.path.exists(file_path):
            print(f"Error: File not found at '{file_path}'")
            return None

        # Load the Excel file, reading the phone column as string
        df = pd.read_excel(file_path, engine='openpyxl', dtype={phone_col: str})
        print(f"Successfully read {len(df)} rows from '{file_path}'.")
        print(f"Columns found: {list(df.columns)}")

        if phone_col not in df.columns:
            print(f"Error: Column '{phone_col}' not found.")
            return None
            
        # Convert phone column to string
        df[phone_col] = df[phone_col].astype(str)

        # Clean phone numbers
        print("Cleaning phone numbers...")
        df['cleaned_phone'] = df[phone_col].apply(clean_phone_number)
        
        # Generate UUID based on cleaned phone numbers
        df[uuid_col] = df['cleaned_phone'].apply(lambda x: generate_uuid_from_phone(x) if x else None)
        print("Generated UUIDs based on cleaned phone numbers.")
        
        # Report the count of rows with invalid/empty phone numbers
        invalid_phones = df[df['cleaned_phone'] == ""].shape[0]
        if invalid_phones > 0:
            print(f"WARNING: {invalid_phones} rows have invalid or empty phone numbers after cleaning.")
        
        # Add Counter column with initial value zero
        df[counter_col] = 0
        print(f"Added '{counter_col}' column with initial value 0.")

        # Remove the original phone column and use the cleaned number
        df = df.drop(columns=[phone_col])
        df.rename(columns={'cleaned_phone': 'mobile'}, inplace=True)
        print("Replaced original phone column with cleaned phone numbers.")

        # --- Remove unnecessary columns ---
        # Define the exact column names to remove, including multi-line ones
        tc_kimlik_col_name = "TC Kimlik Numarası \n(Bu alanda alınan veriler, ÜNİDES proje kapsamında Gençlik ve Spor Bakanlığı tarafından talep edilmektedir.)"
        kvkk_col_name = "KVKK AYDINLATMA METNİ" # Note: The original Excel might have slightly different spacing/casing
        okul_col_name = "Öğrenim gördüğünüz / mezun olduğunuz öğretim kurumu" # Added for removal
        bolum_col_name = "Öğrenim gördüğünüz / mezun olduğunuz bölüm/dal" # Added for removal

        columns_to_remove = [
            'Üniversiteniz',
            'Cinsiyet',
            'Bölümünüz',
            'Kaçıncı sınıftasınız? ',
            'Etkinliği nereden duydunuz? ',
            'Etkinliğimizden beklentileriniz nelerdir? ',
            'Eklemek istediğiniz bir şey var mı? ',
            # 'KVKK AYDINLATMA METNİ', # Removed old version if different
            kvkk_col_name, # Add the exact KVKK column name
            tc_kimlik_col_name, # Add the exact TC Kimlik column name
            okul_col_name, # Add Okul original name
            bolum_col_name, # Add Bolum original name
            'Zaman damgası',
            # Add other potential variations or related columns if needed
            '13. sütun', # Based on CSV header
            '12. sütun'  # Based on CSV header
        ]
        removed_cols_count = 0
        cols_before_removal = df.columns.tolist() # Get columns before removal for accurate checking

        print("Attempting to remove specified columns...")
        # Iterate through a copy of the list to avoid modification issues
        for col_name_to_check in list(columns_to_remove):
            # Check against stripped column names in the DataFrame
            found_col = None
            for df_col in cols_before_removal:
                # Strip both the check name and the DataFrame column name for comparison
                if str(df_col).strip() == str(col_name_to_check).strip():
                    found_col = df_col # Use the actual column name from the DataFrame
                    break

            if found_col and found_col in df.columns: # Check if it still exists in df
                try:
                    df = df.drop(columns=[found_col])
                    print(f" - Removed column: '{found_col}' (Matched based on '{col_name_to_check}')")
                    removed_cols_count += 1
                except KeyError:
                    # Should not happen with the check, but as a safeguard
                    print(f" - INFO: Column '{found_col}' was already removed or not found during drop.")
            else:
                # Only print warning if the column wasn't found by stripping/matching
                if not found_col:
                     print(f" - WARNING: Column matching '{col_name_to_check}' not found. Skipping removal.")


        if removed_cols_count > 0:
            print(f"Successfully removed {removed_cols_count} columns.")
        # else: # Keep this less verbose unless debugging
            # print("WARNING: None of the specified columns were found or removed.")

        # --- Reorder columns (UUID and Counter come first) ---
        cols = df.columns.tolist()
        if uuid_col in cols: cols.remove(uuid_col)
        if counter_col in cols: cols.remove(counter_col)
        df = df[[uuid_col, counter_col] + cols]
        print(f"Moved '{uuid_col}' and '{counter_col}' columns to the beginning.")
        
        # Rename column names (ensure these target names exist after removals)
        rename_map = {
            "Ad-Soyad": "isim",
            "E-posta adresiniz": "mail",
            # "Telefon numaranız": "mobile" # This is handled by cleaned_phone rename earlier
        }
        actual_rename_map = {}
        print("Attempting to rename columns...")
        for original_name, new_name in rename_map.items():
            # Check against stripped column names
            found_original_col = None
            for df_col in df.columns:
                 if str(df_col).strip() == str(original_name).strip():
                     found_original_col = df_col
                     break
            if found_original_col:
                 actual_rename_map[found_original_col] = new_name
                 print(f" - Will rename '{found_original_col}' to '{new_name}'")
            else:
                 print(f" - WARNING: Column matching '{original_name}' not found for renaming.")

        df.rename(columns=actual_rename_map, inplace=True)
        if actual_rename_map:
            print("Finished renaming columns.")
        else:
            print("No columns were renamed.")

        # Convert email column to lowercase using Turkish-specific handling
        if 'mail' in df.columns:
            def turkish_lowercase(text):
                if pd.isna(text):
                    return text
                text = str(text)
                # Replace Turkish 'İ' (U+0130) with 'i' (U+0069)
                text = text.replace('\u0130', '\u0069')
                # Apply standard lower() for other characters (e.g., 'I' -> 'ı')
                return text.lower()

            df['mail'] = df['mail'].apply(turkish_lowercase)
            print("Applied Turkish-specific lowercase conversion to 'mail' column.")
        else:
             print("Warning: 'mail' column not found after renaming. Cannot convert to lowercase.")

        # --- CSV output ---
        if not os.path.exists(CSV_OUTPUT_DIR):
            os.makedirs(CSV_OUTPUT_DIR)
            print(f"Created CSV output directory: '{CSV_OUTPUT_DIR}'")
        output_csv_path = os.path.join(CSV_OUTPUT_DIR, os.path.splitext(os.path.basename(file_path))[0] + '.csv')
        df.to_csv(output_csv_path, index=False, encoding='utf-8-sig')
        print(f"Successfully saved CSV to '{output_csv_path}'.")
        return output_csv_path
    except Exception as e:
        print(f"An unexpected error occurred during Excel processing: {e}")
        import traceback
        traceback.print_exc() # Print full traceback for debugging
        return None

# --- QR Code Generation ---
from QRGenerator import generate_qr_codes_from_csv

# --- Excel with QR Code Generation ---
def generate_excel_with_qr(csv_path, qr_dir, excel_output_path):
    df = pd.read_csv(csv_path, dtype=str)
    wb = Workbook()
    ws = wb.active
    ws.title = "QR Kodlar"
    ws.column_dimensions['A'].width = 20  # Column width for QR
    # Create header row
    headers = ['qr kodlar', 'ad soyad', 'posta', 'numara']
    ws.append(headers)
    row_number = 2
    for _, row in df.iterrows():
        ws.row_dimensions[row_number].height = 100  # Adjust row height according to QR size
        mobile = row.get('mobile', '').strip() # Ensure mobile is stripped

        img_path = None
        if mobile: # Only proceed if mobile is not empty
            img_path = os.path.join(qr_dir, f"{mobile}.png") # Use mobile directly for filename

        # Other columns
        ws.cell(row=row_number, column=2, value=row.get('isim', ''))
        ws.cell(row=row_number, column=3, value=row.get('mail', ''))
        ws.cell(row=row_number, column=4, value=row.get('mobile', ''))
        # Add QR image (if exists)
        if img_path and os.path.exists(img_path):
            img = OpenpyxlImage(img_path)
            img.width = 145
            img.height = 145
            ws.add_image(img, f"A{row_number}")
        else:
            ws.cell(row=row_number, column=1, value="No QR")
            if mobile:
                 print(f"Warning: QR image not found for mobile '{mobile}' at expected path: {img_path}")
        row_number += 1

    # Adjust other column widths
    for col in ['B', 'C', 'D']:
        max_length = 0
        for cell in ws[col]:
            if cell.value:
                max_length = max(max_length, len(str(cell.value)))
        ws.column_dimensions[col].width = max_length + 5  # add padding

    # Ensure the directory exists before saving
    excel_dir = os.path.dirname(excel_output_path)
    if not os.path.exists(excel_dir):
        os.makedirs(excel_dir)
    wb.save(excel_output_path)
    # Add print statement for clarity
    print(f"Successfully generated Excel with QR codes at: '{excel_output_path}'")


# --- File Operations ---
# from FileOperations import create_directory_if_not_exists # This line is now handled above

# --- Main Execution ---
if __name__ == "__main__":
    # --- Find Input Excel File ---
    if not os.path.exists(INPUT_DIR):
        os.makedirs(INPUT_DIR)
        print(f"Created directory '{INPUT_DIR}'. Please place your single Excel (.xlsx) file inside it and rerun the script.")
        sys.exit(1)

    excel_files = [f for f in os.listdir(INPUT_DIR) if f.endswith('.xlsx') and not f.startswith('~')] # Ignore temp files

    if len(excel_files) == 0:
        print(f"Error: No Excel (.xlsx) file found in the '{INPUT_DIR}' directory.")
        print("Please place the input Excel file there.")
        sys.exit(1)
    elif len(excel_files) > 1:
        print(f"Error: Multiple Excel (.xlsx) files found in the '{INPUT_DIR}' directory:")
        for f in excel_files:
            print(f" - {f}")
        print("Please ensure only one Excel file is present in the input directory.")
        sys.exit(1)
    else:
        excel_file_path = os.path.join(INPUT_DIR, excel_files[0])
        # Encode path for safe printing
        safe_excel_file_path_repr = repr(excel_file_path.encode(sys.stdout.encoding, errors='replace').decode(sys.stdout.encoding, errors='replace'))
        print(f"Using input file: {safe_excel_file_path_repr}") # Use encoded path representation

    # Define output directory paths
    designed_qr_output_dir = os.path.join('output', 'designed_qr')
    excel_output_dir = os.path.join('output', 'excel') # Define Excel output dir

    # Pass the dynamically found excel_file_path
    csv_file = process_excel(excel_file_path, PHONE_COLUMN_NAME, UUID_COLUMN_NAME, COUNTER_COLUMN_NAME)
    if csv_file:
        # Ensure output directories exist
        create_directory_if_not_exists(QR_OUTPUT_DIR)
        # create_directory_if_not_exists(os.path.dirname(EXCEL_OUTPUT_PATH)) # Removed old static path check
        create_directory_if_not_exists(excel_output_dir) # Ensure dynamic Excel output dir exists
        create_directory_if_not_exists(designed_qr_output_dir)

        # --- QR code generation and Excel ---
        generate_qr_codes_from_csv(csv_file, UUID_COLUMN_NAME, 'mobile', QR_OUTPUT_DIR)
        print("QR code generation completed.")

        # Dynamically generate output Excel path
        base_name = os.path.splitext(os.path.basename(excel_file_path))[0]
        dynamic_excel_output_path = os.path.join(excel_output_dir, f"{base_name}_modified.xlsx")

        generate_excel_with_qr(csv_file, QR_OUTPUT_DIR, dynamic_excel_output_path) # Use dynamic path
        # print("Excel with QR codes generation completed.") # Removed, moved inside function
        print("CSV generation completed.") # This seems misplaced, maybe move after CSV save? Or keep here.

        # --- QR Design ---
        print("\n--- Starting QR Design Process ---")
        template_image_path = "tasarim.jpg"
        if not os.path.exists(template_image_path):
             print(f"Warning: Template image '{template_image_path}' not found. Skipping QR design.")
             can_design = False # Still needed for QR emails
        else:
            overlay_qr_on_template(QR_OUTPUT_DIR, template_image_path, designed_qr_output_dir, csv_path=csv_file)
            print("QR Design process completed.")
            can_design = True

        # --- Firebase Sync ---
        firebase_choice = input("Do you want to sync with Firebase? (yes/no): ").strip().lower()
        if firebase_choice == 'yes':
            print("\n--- Starting Firebase Synchronization ---")
            try:
                if not os.path.exists(FIREBASE_SYNC_SCRIPT):
                    print(f"Error: Firebase sync script '{FIREBASE_SYNC_SCRIPT}' not found.")
                else:
                    python_executable = sys.executable
                    result = subprocess.run(
                        [python_executable, FIREBASE_SYNC_SCRIPT, csv_file],
                        capture_output=True, text=True, check=True
                    )
                    print("Firebase sync script output:")
                    print(result.stdout)
                    if result.stderr: print("Firebase sync script errors:", result.stderr)
                    print("--- Firebase Synchronization Finished ---")
            except FileNotFoundError:
                print(f"Error: Python executable not found at '{sys.executable}'. Cannot run sync script.")
            except subprocess.CalledProcessError as e:
                print(f"Error: Firebase sync script failed with exit code {e.returncode}.")
                print("Output:")
                print(e.stdout)
                print("Errors:")
                print(e.stderr)
            except Exception as e:
                print(f"An unexpected error occurred during Firebase sync: {e}")

        # --- Send QR Emails ---
        email_choice = input("Do you want to send emails with designed QR codes? (yes/no): ").strip().lower()
        if email_choice == 'yes':
            if can_design and os.path.exists(csv_file) and os.path.exists(designed_qr_output_dir):
                print("\n--- Starting QR Email Sending Process ---")
                send_qr_codes(
                    csv_file, designed_qr_output_dir, SENDER_EMAIL,
                    SENDER_PASSWORD, SMTP_SERVER, SMTP_PORT
                )
                print("--- QR Email Sending Process Finished ---")
            else:
                print(f"Prerequisites not met (Template exists: {can_design}, CSV exists: {os.path.exists(csv_file)}, Designed QRs exist: {os.path.exists(designed_qr_output_dir)}). QR Emails cannot be sent.")

        print("\nAll processes have been completed.")
        print("Note: To generate and send certificates, run 'python CertificateGeneratorSender.py' separately.") # Added note
    else:
        print("CSV file generation failed. Skipping subsequent steps.")
