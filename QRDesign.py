import os
from PIL import Image # Removed ImageDraw, ImageFont
import pandas as pd
from FileOperations import create_directory_if_not_exists
from tqdm import tqdm

def overlay_qr_on_template(qr_dir, template_path, output_dir, uuid_column=None, csv_path=None):
    """
    Overlay QR codes on a template image. Skips if the designed QR already exists.
    
    Args:
        qr_dir (str): Directory containing generated QR codes
        template_path (str): Path to the template image
        output_dir (str): Directory to save the designed QR codes
        uuid_column (str, optional): Column name for UUIDs in CSV
        csv_path (str, optional): Path to the CSV file containing data
    """
    create_directory_if_not_exists(output_dir)
    
    # Load the template image
    try:
        template = Image.open(template_path)
        template_width, template_height = template.size
        print(f"Loaded template image: {template_path} ({template_width}x{template_height})")
    except Exception as e:
        print(f"Error loading template image: {e}")
        return False
    
    # Check how to process the QR codes
    processed_count = 0
    skipped_existing = 0 # Counter for existing designed QRs
    skipped_not_found = 0 # Counter for missing basic QRs

    if csv_path and os.path.exists(csv_path):
        # Process based on CSV data
        df = pd.read_csv(csv_path, dtype=str)
        print(f"Processing {len(df)} records from CSV for QR design...")
        for index, row in tqdm(df.iterrows(), total=len(df), desc="Designing QR codes (CSV)"):
            mobile = row.get('mobile', '').strip()
            if not mobile:
                continue # Skip rows with no mobile number

            # Construct the expected output path first
            output_path = os.path.join(output_dir, f"{mobile}_designed.png")

            # Check if the designed QR code file already exists
            if os.path.exists(output_path):
                # print(f"Skipping row {index + 2}: Designed QR already exists for phone {mobile} at '{output_path}'") # Optional: Can be verbose
                skipped_existing += 1
                continue

            # Find the basic QR code file
            qr_file_path = os.path.join(qr_dir, f"{mobile}.png")
            if not os.path.exists(qr_file_path):
                print(f"Warning: Basic QR code not found for phone: {mobile} (Row {index + 2}). Cannot design.")
                skipped_not_found += 1
                continue
            
            # Create a copy of the template for each QR code
            new_img = template.copy()
            
            # Load QR code
            try:
                qr_img = Image.open(qr_file_path)
            except Exception as img_e:
                print(f"Error opening basic QR image {qr_file_path} for phone {mobile}: {img_e}")
                continue # Skip this QR if it can't be opened

            # Calculate position to center the QR code on the blue area
            # Assuming the blue area is centered and 1500x1500px as shown in the image
            x_position = (template_width - 1500) // 2
            y_position = (template_height - 1500) // 2
            
            try:
                qr_img_resized = qr_img.resize((1500, 1500))
            except Exception as resize_e:
                print(f"Error resizing QR image {qr_file_path} for phone {mobile}: {resize_e}")
                continue # Skip if resizing fails

            # Paste the QR code onto the template
            new_img.paste(qr_img_resized, (x_position, y_position))
            
            # Save the result
            try:
                new_img.save(output_path)
                processed_count += 1
            except Exception as save_e:
                 print(f"Error saving designed QR image {output_path} for phone {mobile}: {save_e}")

    else:
        # Process all QR code files in the directory (Fallback if no CSV)
        print("Processing QR files directly from directory (CSV not provided)...")
        qr_files = [f for f in os.listdir(qr_dir) if f.endswith('.png')]
        for qr_file in tqdm(qr_files, desc="Designing QR codes (Dir)"):
            qr_file_path = os.path.join(qr_dir, qr_file)
            base_name = os.path.splitext(qr_file)[0] # Usually the mobile number

            # Construct the expected output path
            output_filename = f"{base_name}_designed.png"
            output_path = os.path.join(output_dir, output_filename)

            # Check if the designed QR code file already exists
            if os.path.exists(output_path):
                # print(f"Skipping: Designed QR already exists for {base_name} at '{output_path}'") # Optional
                skipped_existing += 1
                continue

            # Create a copy of the template for each QR code
            new_img = template.copy()
            
            # Load QR code
            try:
                qr_img = Image.open(qr_file_path)
            except Exception as img_e:
                print(f"Error opening basic QR image {qr_file_path}: {img_e}")
                continue

            # Calculate position to center the QR code on the blue area
            x_position = (template_width - 1500) // 2
            y_position = (template_height - 1500) // 2
            
            # Resize QR code to fit the blue area (1500x1500)
            try:
                qr_img_resized = qr_img.resize((1500, 1500))
            except Exception as resize_e:
                print(f"Error resizing QR image {qr_file_path}: {resize_e}")
                continue

            # Paste the QR code onto the template
            new_img.paste(qr_img_resized, (x_position, y_position))
            
            # Save the result
            try:
                new_img.save(output_path)
                processed_count += 1
            except Exception as save_e:
                 print(f"Error saving designed QR image {output_path}: {save_e}")
    
    print(f"\nQR Design Summary:")
    print(f" - Successfully created/overlaid: {processed_count}")
    print(f" - Skipped (designed QR already exists): {skipped_existing}")
    if csv_path: # Only relevant if processing via CSV
        print(f" - Skipped (basic QR not found): {skipped_not_found}")
    print(f" - Output directory: '{output_dir}'")
    return True
