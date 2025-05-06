import os
import sys
import pandas as pd
import firebase_admin
from firebase_admin import credentials, firestore
import glob
from dotenv import load_dotenv

# --- Configuration ---
load_dotenv() # Load environment variables if needed, though not strictly used here

SERVICE_ACCOUNT_KEY_PATH = 'qr-deneme.json'
COLLECTION_NAME = 'users'
CSV_INPUT_DIR = os.path.join('output', 'csv')
EXCEL_OUTPUT_DIR = os.path.join('output', 'excel')
OUTPUT_FILENAME = 'katilimcilar.xlsx'

# --- Column Names ---
# Firebase field names
FIREBASE_NAME_COL = 'Ad-Soyad'
FIREBASE_EMAIL_COL = 'Eposta'
FIREBASE_MOBILE_COL = 'Telefon numaranız'
FIREBASE_COUNTER_COL = 'Counter'

# CSV column names (from *_form.csv, matching original Excel)
# Using 'mobile' as the key after cleaning in DataExtractor.py
CSV_MOBILE_COL = 'mobile'
# Exact names from the original form/Excel needed for lookup in *_form.csv
CSV_TCKN_COL = "TC Kimlik Numarası \n(Bu alanda alınan veriler, ÜNİDES proje kapsamında Gençlik ve Spor Bakanlığı tarafından talep edilmektedir.)"
CSV_BIRTHDATE_COL = "Doğum tarihiniz\n(Bu alanda alınan veriler, ÜNİDES proje kapsamında Gençlik ve Spor Bakanlığı tarafından talep edilmektedir.)"

# Output Excel column names
OUT_NAME_COL = 'Name'
OUT_EMAIL_COL = 'Email'
OUT_MOBILE_COL = 'Mobile'
OUT_TCKN_COL = 'TCKN'
OUT_BIRTHDATE_COL = 'Birth Date'

# --- Utility Functions ---

def create_directory_if_not_exists(directory_path):
    """Create a directory if it does not already exist."""
    if not os.path.exists(directory_path):
        os.makedirs(directory_path)
        print(f"Created directory: {directory_path}")

def initialize_firebase():
    """Initializes the Firebase Admin SDK."""
    if not os.path.exists(SERVICE_ACCOUNT_KEY_PATH):
        print(f"Error: Service account key file not found at '{SERVICE_ACCOUNT_KEY_PATH}'")
        sys.exit(1)
    try:
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

def find_form_csv(csv_dir):
    """Finds the '*_form.csv' file in the specified directory."""
    search_pattern = os.path.join(csv_dir, '*_form.csv')
    form_csv_files = glob.glob(search_pattern)

    if not form_csv_files:
        print(f"Error: No '*_form.csv' file found in '{csv_dir}'.")
        print("Please ensure DataExtractor.py has been run successfully to generate the intermediate CSV.")
        return None
    elif len(form_csv_files) > 1:
        print(f"Warning: Multiple '*_form.csv' files found in '{csv_dir}'. Using the first one found:")
        print(f" - {os.path.basename(form_csv_files[0])}")
        return form_csv_files[0]
    else:
        print(f"Found form CSV file: {os.path.basename(form_csv_files[0])}")
        return form_csv_files[0]

def get_attendees_from_firebase(db):
    """Fetches users from Firestore where Counter > 1."""
    users_ref = db.collection(COLLECTION_NAME)
    attendees_dict = {}
    try:
        # Query for documents where Counter is greater than 1
        query = users_ref.where(filter=firestore.FieldFilter(FIREBASE_COUNTER_COL, '>', 1)).stream()
        print(f"Fetching attendees from Firebase (where {FIREBASE_COUNTER_COL} > 1)...")

        attendee_count = 0
        for doc in query:
            data = doc.to_dict()
            mobile = data.get(FIREBASE_MOBILE_COL)
            if mobile: # Use mobile number as the key
                attendees_dict[str(mobile).strip()] = {
                    OUT_NAME_COL: data.get(FIREBASE_NAME_COL),
                    OUT_EMAIL_COL: data.get(FIREBASE_EMAIL_COL),
                    OUT_MOBILE_COL: str(mobile).strip(), # Store mobile for the final output
                    # Initialize TCKN and Birth Date as None, to be filled from CSV
                    OUT_TCKN_COL: None,
                    OUT_BIRTHDATE_COL: None
                }
                attendee_count += 1
            else:
                print(f"Warning: Found attendee document (ID: {doc.id}) with missing mobile number. Skipping.")

        print(f"Found {attendee_count} attendees with {FIREBASE_COUNTER_COL} > 1 in Firebase.")
        if attendee_count == 0:
             print("No attendees met the criteria.")
        return attendees_dict
    except Exception as e:
        print(f"Error fetching attendees from Firestore: {e}")
        return {}

def read_form_csv_data(csv_path):
    """Reads the form CSV and extracts relevant columns, indexed by mobile."""
    try:
        # Read mobile as string, handle potential missing columns gracefully
        df = pd.read_csv(csv_path, dtype={CSV_MOBILE_COL: str})
        print(f"Read {len(df)} rows from {os.path.basename(csv_path)}.")

        # Select necessary columns, checking if they exist
        required_cols = [CSV_MOBILE_COL]
        optional_cols = [CSV_TCKN_COL, CSV_BIRTHDATE_COL]
        cols_to_read = []
        missing_cols = []

        for col in required_cols + optional_cols:
             # Check if column exists, potentially with different whitespace/casing
             found_col = None
             for df_col in df.columns:
                  if str(df_col).strip() == str(col).strip():
                       found_col = df_col # Use the actual column name from the DataFrame
                       break
             if found_col:
                  cols_to_read.append(found_col)
             elif col in optional_cols:
                  print(f"Warning: Optional column '{col}' not found in {os.path.basename(csv_path)}. It will be omitted.")
             else: # Required column missing
                  missing_cols.append(col)

        if missing_cols:
             print(f"Error: Required column(s) not found in {os.path.basename(csv_path)}: {', '.join(missing_cols)}")
             return None

        df_subset = df[cols_to_read].copy()

        # Rename columns to standardized names used internally, handling missing optional cols
        rename_map = {CSV_MOBILE_COL: CSV_MOBILE_COL} # Mobile is key, keep its name for now
        if CSV_TCKN_COL in df_subset.columns:
             rename_map[CSV_TCKN_COL] = OUT_TCKN_COL
        if CSV_BIRTHDATE_COL in df_subset.columns:
             rename_map[CSV_BIRTHDATE_COL] = OUT_BIRTHDATE_COL

        # Find the actual column name matching CSV_MOBILE_COL for renaming/indexing
        actual_mobile_col_name = None
        for df_col in df_subset.columns:
             if str(df_col).strip() == str(CSV_MOBILE_COL).strip():
                  actual_mobile_col_name = df_col
                  break

        if not actual_mobile_col_name:
             print(f"Critical Error: Could not find the mobile column ('{CSV_MOBILE_COL}') in the CSV subset.")
             return None

        # Perform renaming using actual found column names
        actual_rename_map = {}
        for csv_col_pattern, out_col in rename_map.items():
             found_actual_col = None
             for df_col in df_subset.columns:
                  if str(df_col).strip() == str(csv_col_pattern).strip():
                       found_actual_col = df_col
                       break
             if found_actual_col:
                  actual_rename_map[found_actual_col] = out_col

        df_subset.rename(columns=actual_rename_map, inplace=True)

        # Set index to the actual mobile column name found
        df_subset.set_index(actual_mobile_col_name, inplace=True)
        return df_subset

    except FileNotFoundError:
        print(f"Error: CSV file not found at '{csv_path}'")
        return None
    except Exception as e:
        print(f"Error reading or processing CSV file '{os.path.basename(csv_path)}': {e}")
        return None

# --- Main Execution ---
if __name__ == "__main__":
    print("--- Starting Attendee Data Export ---")

    # Initialize Firebase
    db_client = initialize_firebase()
    if not db_client:
        sys.exit(1)

    # Fetch attendees from Firebase
    firebase_attendees = get_attendees_from_firebase(db_client)

    if not firebase_attendees:
        print("No attendees found in Firebase matching criteria. Exiting.")
        sys.exit(0)

    # Find and read the form CSV data
    form_csv_path = find_form_csv(CSV_INPUT_DIR)
    if not form_csv_path:
        sys.exit(1)

    csv_data = read_form_csv_data(form_csv_path)
    if csv_data is None:
        print("Failed to load or process data from form CSV. Exiting.")
        sys.exit(1)

    # Combine Firebase data with CSV data
    final_attendee_list = []
    print("\nMatching Firebase attendees with CSV data...")
    matched_count = 0
    unmatched_count = 0

    for mobile, fb_data in firebase_attendees.items():
        if mobile in csv_data.index:
            csv_row = csv_data.loc[mobile]
            # Update TCKN and Birth Date if columns exist in csv_data
            if OUT_TCKN_COL in csv_row.index:
                 fb_data[OUT_TCKN_COL] = csv_row[OUT_TCKN_COL]
            if OUT_BIRTHDATE_COL in csv_row.index:
                 # Attempt to format birth date if it's read as datetime
                 try:
                      birth_date_val = pd.to_datetime(csv_row[OUT_BIRTHDATE_COL], errors='coerce')
                      if pd.notna(birth_date_val):
                           fb_data[OUT_BIRTHDATE_COL] = birth_date_val.strftime('%Y-%m-%d') # Format as YYYY-MM-DD
                      else:
                           fb_data[OUT_BIRTHDATE_COL] = csv_row[OUT_BIRTHDATE_COL] # Keep original if not date
                 except Exception:
                      fb_data[OUT_BIRTHDATE_COL] = csv_row[OUT_BIRTHDATE_COL] # Keep original on error

            final_attendee_list.append(fb_data)
            matched_count += 1
        else:
            # Attendee found in Firebase but not in CSV (should be rare if CSV is source)
            print(f"Warning: Attendee with mobile {mobile} found in Firebase but not in form CSV. Including basic data.")
            final_attendee_list.append(fb_data) # Add anyway, TCKN/BirthDate will be None
            unmatched_count += 1

    print(f"Finished matching. Matched: {matched_count}, Unmatched in CSV: {unmatched_count}")

    if not final_attendee_list:
        print("No attendee data to write to Excel. Exiting.")
        sys.exit(0)

    # Create DataFrame and save to Excel
    output_df = pd.DataFrame(final_attendee_list)

    # Reorder columns for the final Excel file
    final_columns_order = [OUT_NAME_COL, OUT_EMAIL_COL, OUT_MOBILE_COL, OUT_TCKN_COL, OUT_BIRTHDATE_COL]
    # Ensure only existing columns are included in the final order
    final_columns_order = [col for col in final_columns_order if col in output_df.columns]
    output_df = output_df[final_columns_order]

    # Ensure output directory exists
    create_directory_if_not_exists(EXCEL_OUTPUT_DIR)
    output_excel_path = os.path.join(EXCEL_OUTPUT_DIR, OUTPUT_FILENAME)

    try:
        output_df.to_excel(output_excel_path, index=False, engine='openpyxl')
        print(f"\nSuccessfully exported {len(output_df)} attendees to:")
        print(f"'{output_excel_path}'")
    except Exception as e:
        print(f"\nError saving data to Excel file '{output_excel_path}': {e}")

    print("\n--- Attendee Data Export Finished ---")
