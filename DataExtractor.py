import os
import pandas as pd
import uuid
import qrcode
import re
import openpyxl
from openpyxl import Workbook
from openpyxl.drawing.image import Image as OpenpyxlImage

# --- Configuration ---
EXCEL_FILE_PATH = 'yanitlar.xlsx'
PHONE_COLUMN_NAME = 'Telefon numaranız'
UUID_COLUMN_NAME = 'UUID'
COUNTER_COLUMN_NAME = 'Counter'
NAMESPACE = uuid.NAMESPACE_DNS
CSV_OUTPUT_DIR = os.path.join('output', 'csv')
QR_OUTPUT_DIR = os.path.join('output', 'qr')

# --- Utility Functions ---
def clean_phone_number(phone_str):
    if not isinstance(phone_str, str):
        phone_str = str(phone_str)
    cleaned = re.sub(r'\D', '', phone_str)
    if cleaned.startswith('90'):
        cleaned = cleaned[2:]
    elif cleaned.startswith('0'):
        cleaned = cleaned[1:]
    return cleaned

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

        df = pd.read_excel(file_path, engine='openpyxl')
        print(f"Successfully read {len(df)} rows from '{file_path}'.")
        print(f"Columns found: {list(df.columns)}")

        if phone_col not in df.columns:
            print(f"Error: Column '{phone_col}' not found.")
            return None

        df[uuid_col] = df[phone_col].apply(generate_uuid_from_phone)
        print(f"Generated UUIDs based on '{phone_col}' column.")
        df[counter_col] = 0
        print(f"Added '{counter_col}' column with initial value 0.")

        # --- Remove unnecessary columns ---
        columns_to_remove = [
            'Üniversiteniz',
            'Cinsiyet',
            'Bölümünüz',
            'Kaçıncı sınıftasınız? ',
            'Etkinliği nereden duydunuz? ',
            'Etkinliğimizden beklentileriniz nelerdir? ',
            'Eklemek istediğiniz bir şey var mı? ',
            'KVKK AYDINLATMA METNİ',
            'Zaman damgası'
            # Do not include e-postanı çıkaran bir satır here
        ]
        removed_cols_count = 0
        print("Attempting to remove specified columns...")
        for col_name in columns_to_remove:
            if col_name in df.columns:
                df = df.drop(columns=[col_name])
                print(f" - Removed column: '{col_name}'")
                removed_cols_count += 1
            else:
                print(f" - WARNING: Column '{col_name}' not found. Skipping removal.")

        if removed_cols_count > 0:
            print(f"Successfully removed {removed_cols_count} columns.")
        else:
            print("WARNING: None of the specified columns were found.")

        # --- Reorder columns (UUID and Counter first) ---
        cols = df.columns.tolist()
        if uuid_col in cols: cols.remove(uuid_col)
        if counter_col in cols: cols.remove(counter_col)
        df = df[[uuid_col, counter_col] + cols]
        print(f"Moved '{uuid_col}' and '{counter_col}' columns to the beginning.")
        
        # Rename output columns as specified
        df.rename(columns={
            "Ad-Soyad": "isim",
            "E-posta adresiniz": "mail",
            "Telefon numaranız": "mobile"
        }, inplace=True)
        print("Renamed columns: Ad-Soyad -> isim, E-posta adresiniz -> mail, Telefon numaranız -> mobile")

        # --- Save CSV to output/csv ---
        if not os.path.exists(CSV_OUTPUT_DIR):
            os.makedirs(CSV_OUTPUT_DIR)
            print(f"Created CSV output directory: '{CSV_OUTPUT_DIR}'")
        output_csv_path = os.path.join(CSV_OUTPUT_DIR, os.path.splitext(os.path.basename(file_path))[0] + '.csv')
        df.to_csv(output_csv_path, index=False, encoding='utf-8-sig')
        print(f"Successfully saved CSV to '{output_csv_path}'.")
        return output_csv_path
    except Exception as e:
        print(f"An unexpected error occurred during Excel processing: {e}")
        return None

# --- QR Code Generation ---
def generate_qr_codes_from_csv(csv_path, uuid_col, phone_col, output_dir):
    try:
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)
            print(f"Created QR output directory: '{output_dir}'")
        df = pd.read_csv(csv_path, dtype={uuid_col: str, phone_col: str})
        print(f"CSV file read from '{csv_path}' with {len(df)} rows.")

        generated_count = 0
        skipped_count = 0
        for index, row in df.iterrows():
            uuid_value = row.get(uuid_col)
            phone_value = row.get(phone_col)
            if pd.isna(uuid_value) or not str(uuid_value).strip():
                print(f"Skipping row {index+2}: empty UUID.")
                skipped_count += 1
                continue
            uuid_value = str(uuid_value).strip()
            # Use cleaned phone number as filename if available; fallback to UUID
            safe_filename = clean_phone_number(str(phone_value)) if phone_value and str(phone_value).strip() else uuid_value

            qr = qrcode.QRCode(
                version=1,
                error_correction=qrcode.constants.ERROR_CORRECT_L,
                box_size=3,  # updated from 2 to 3 for slightly larger QR code
                border=4,
            )
            qr.add_data(uuid_value)
            qr.make(fit=True)
            img = qr.make_image(fill_color="black", back_color="white")
            output_path = os.path.join(output_dir, f"{safe_filename}.png")
            try:
                if os.path.exists(output_path):
                    print(f"Row {index+2}: '{output_path}' exists. Overwriting.")
                img.save(output_path)
                generated_count += 1
            except Exception as e:
                print(f"Error saving QR for row {index+2} with UUID '{uuid_value}': {e}")
                skipped_count += 1

        print("-" * 20)
        print(f"QR Code Generation Summary:")
        print(f"  Successfully generated: {generated_count}")
        print(f"  Skipped: {skipped_count}")
    except Exception as e:
        print(f"An unexpected error occurred during QR code generation: {e}")

# --- Excel with QR Code Generation ---
def generate_excel_with_qr(csv_path, qr_dir, excel_output_path):
    df = pd.read_csv(csv_path, dtype=str)
    wb = Workbook()
    ws = wb.active
    ws.title = "QR Kodlar"
    ws.column_dimensions['A'].width = 12  # set QR column width to fit QR images
    # Write header row
    headers = ['qr kodlar', 'ad soyad', 'posta', 'numara']
    ws.append(headers)
    row_number = 2
    for _, row in df.iterrows():
        ws.row_dimensions[row_number].height = 60  # adjust row height to fit QR images
        mobile = row.get('mobile', '')
        uuid_value = row.get(f"{UUID_COLUMN_NAME}", '')
        safe_filename = clean_phone_number(mobile) if mobile and str(mobile).strip() else str(uuid_value).strip()
        img_path = os.path.join(qr_dir, f"{safe_filename}.png")
        # Write text columns
        ws.cell(row=row_number, column=2, value=row.get('isim', ''))
        ws.cell(row=row_number, column=3, value=row.get('mail', ''))
        ws.cell(row=row_number, column=4, value=row.get('mobile', ''))
        # Insert QR image in column A if exists
        if os.path.exists(img_path):
            img = OpenpyxlImage(img_path)
            ws.add_image(img, f"A{row_number}")
        else:
            ws.cell(row=row_number, column=1, value="No QR")
        row_number += 1
    # Adjust widths for columns B, C, and D based on the longest content
    for col in ['B', 'C', 'D']:
        max_length = 0
        for cell in ws[col]:
            if cell.value:
                max_length = max(max_length, len(str(cell.value)))
        ws.column_dimensions[col].width = max_length + 5  # add padding
    # Ensure output directory exists
    excel_dir = os.path.dirname(excel_output_path)
    if not os.path.exists(excel_dir):
        os.makedirs(excel_dir)
    wb.save(excel_output_path)

# --- Main Execution ---
if __name__ == "__main__":
    csv_file = process_excel(EXCEL_FILE_PATH, PHONE_COLUMN_NAME, UUID_COLUMN_NAME, COUNTER_COLUMN_NAME)
    if csv_file:
        # Update phone column name to 'mobile' for QR code generation
        generate_qr_codes_from_csv(csv_file, UUID_COLUMN_NAME, 'mobile', QR_OUTPUT_DIR)
        print("QR code generation completed.")
        excel_output_path = os.path.join('output', 'excel', 'qrKodlar.xlsx')
        generate_excel_with_qr(csv_file, QR_OUTPUT_DIR, excel_output_path)
        print("Excel with QR codes generation completed.")
        # Final summary logs
        print("CSV generation completed.")
        print("All processes have been completed.")
