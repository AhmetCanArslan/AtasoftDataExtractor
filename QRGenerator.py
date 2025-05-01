import os
import qrcode
import pandas as pd

def generate_qr_codes_from_csv(csv_path, uuid_column, phone_column, output_dir):
    """
    Generate QR codes from a CSV file, using phone numbers for filenames.
    Skips generation if a QR code file for the phone number already exists.

    Args:
        csv_path (str): Path to the CSV file.
        uuid_column (str): Column name for UUIDs (data for QR code).
        phone_column (str): Column name for phone numbers (used for filename).
        output_dir (str): Directory to save the QR codes.
    """
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    df = pd.read_csv(csv_path, dtype=str)
    generated_count = 0
    skipped_uuid = 0
    skipped_phone = 0
    skipped_existing = 0 # Counter for existing QR codes
    for index, row in df.iterrows():
        uuid_value = row.get(uuid_column, '').strip()
        if not uuid_value:
            skipped_uuid += 1
            continue

        phone_value = row.get(phone_column, '').strip()
        if not phone_value:
            skipped_phone += 1
            continue

        # Construct the expected image path
        img_path = os.path.join(output_dir, f"{phone_value}.png")

        # Check if the QR code file already exists
        if os.path.exists(img_path):
            skipped_existing += 1
            continue

        # Proceed with QR generation if file doesn't exist
        qr = qrcode.QRCode(
            version=1,
            error_correction=qrcode.constants.ERROR_CORRECT_L,
            box_size=10,
            border=4,
        )
        qr.add_data(uuid_value)
        qr.make(fit=True)

        img = qr.make_image(fill_color="black", back_color="white")
        img_path = os.path.join(output_dir, f"{phone_value}.png")
        try:
            img.save(img_path)
            generated_count += 1
        except Exception as e:
            print(f"Error saving QR code for phone {phone_value}: {e}")

    print(f"\nQR Code Generation Summary:")
    print(f" - Successfully generated: {generated_count}")
    print(f" - Skipped (missing UUID): {skipped_uuid}")
    print(f" - Skipped (missing Phone): {skipped_phone}")
    print(f" - Skipped (already exists): {skipped_existing}") # Added existing count
    print(f" - Total processed: {len(df)}")