import os
from PIL import Image
import pandas as pd
from FileOperations import create_directory_if_not_exists
from tqdm import tqdm

def overlay_qr_on_template(qr_dir, template_path, output_dir, uuid_column=None, csv_path=None):
    """
    Overlay QR codes on a template image.
    
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
    if csv_path and os.path.exists(csv_path):
        # Process based on CSV data
        df = pd.read_csv(csv_path, dtype=str)
        for _, row in tqdm(df.iterrows(), total=len(df), desc="Processing QR codes"):
            mobile = row.get('mobile', '').strip()
            if not mobile:
                continue
            
            # Find the QR code file
            qr_file_path = os.path.join(qr_dir, f"{mobile}.png")
            if not os.path.exists(qr_file_path):
                print(f"QR code not found for phone: {mobile}")
                continue
            
            # Create a copy of the template for each QR code
            new_img = template.copy()
            
            # Load QR code
            qr_img = Image.open(qr_file_path)
            
            # Calculate position to center the QR code on the blue area
            # Assuming the blue area is centered and 1500x1500px as shown in the image
            x_position = (template_width - 1500) // 2
            y_position = (template_height - 1500) // 2
            
            qr_img_resized = qr_img.resize((1500, 1500))
            
            # Paste the QR code onto the template
            new_img.paste(qr_img_resized, (x_position, y_position))
            
            # Save the result
            output_path = os.path.join(output_dir, f"{mobile}_designed.png")
            new_img.save(output_path)
            processed_count += 1
    else:
        # Process all QR code files in the directory
        qr_files = [f for f in os.listdir(qr_dir) if f.endswith('.png')]
        for qr_file in tqdm(qr_files, desc="Processing QR codes"):
            qr_file_path = os.path.join(qr_dir, qr_file)
            
            # Create a copy of the template for each QR code
            new_img = template.copy()
            
            # Load QR code
            qr_img = Image.open(qr_file_path)
            
            # Calculate position to center the QR code on the blue area
            # Assuming the blue area is centered and 1500x1500px as shown in the image
            x_position = (template_width - 1500) // 2
            y_position = (template_height - 1500) // 2
            
            # Resize QR code to fit the blue area (1500x1500)
            qr_img_resized = qr_img.resize((1500, 1500))
            
            # Paste the QR code onto the template
            new_img.paste(qr_img_resized, (x_position, y_position))
            
            # Save the result
            output_filename = os.path.splitext(qr_file)[0] + "_designed.png"
            output_path = os.path.join(output_dir, output_filename)
            new_img.save(output_path)
            processed_count += 1
    
    print(f"Successfully created {processed_count} designed QR codes in '{output_dir}'")
    return True
