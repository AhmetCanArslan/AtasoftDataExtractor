import os
import qrcode
import pandas as pd

def generate_qr_codes_from_csv(csv_path, uuid_column, phone_column, output_dir):
    """
    Generate QR codes from a CSV file.

    Args:
        csv_path (str): Path to the CSV file.
        uuid_column (str): Column name for UUIDs.
        phone_column (str): Column name for phone numbers (not used).
        output_dir (str): Directory to save the QR codes.
    """
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    df = pd.read_csv(csv_path, dtype=str)
    for _, row in df.iterrows():
        uuid_value = row.get(uuid_column, '').strip()

        if not uuid_value:
            print("Skipping row with no valid UUID.")
            continue

        qr = qrcode.QRCode(
            version=1,
            error_correction=qrcode.constants.ERROR_CORRECT_L,
            box_size=10,
            border=4,
        )
        qr.add_data(uuid_value)
        qr.make(fit=True)

        img = qr.make_image(fill_color="black", back_color="white")
        img_path = os.path.join(output_dir, f"{uuid_value}.png")
        img.save(img_path)
        print(f"Generated QR code for UUID: {uuid_value}")