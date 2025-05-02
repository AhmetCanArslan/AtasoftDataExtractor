import os
from PIL import Image, ImageDraw, ImageFont # Added ImageDraw, ImageFont
import pandas as pd
from FileOperations import create_directory_if_not_exists
from tqdm import tqdm

# --- Helper function for Turkish character capitalization ---
def turkish_capitalize_name(name):
    """Capitalizes a name string respecting Turkish characters."""
    if not isinstance(name, str):
        return "" # Return empty string if input is not a string

    def turkish_upper(char):
        if char == 'i':
            return 'İ'
        elif char == 'ı':
            return 'I'
        else:
            # Use standard upper() for other characters, handling potential errors
            try:
                return char.upper()
            except AttributeError:
                return char # Return original if not a string-like char

    def turkish_lower(char):
        if char == 'İ':
            return 'i'
        elif char == 'I':
            return 'ı'
        else:
            # Use standard lower() for other characters, handling potential errors
            try:
                return char.lower()
            except AttributeError:
                return char # Return original if not a string-like char

    parts = name.split()
    capitalized_parts = []
    for part in parts:
        if not part:
            continue
        first_char = turkish_upper(part[0])
        rest_chars = "".join(turkish_lower(c) for c in part[1:])
        capitalized_parts.append(first_char + rest_chars)
    return " ".join(capitalized_parts)
# --- End of helper function ---

def overlay_qr_on_template(qr_dir, template_path, output_dir, uuid_column=None, csv_path=None):
    """
    Overlay QR codes and participant names on a template image. Skips if the designed QR already exists.
    
    Args:
        qr_dir (str): Directory containing generated QR codes
        template_path (str): Path to the template image
        output_dir (str): Directory to save the designed QR codes
        uuid_column (str, optional): Column name for UUIDs in CSV (Not directly used here but kept for signature consistency)
        csv_path (str, optional): Path to the CSV file containing data (mobile, isim)
    """
    create_directory_if_not_exists(output_dir)
    
    # --- Font Configuration ---
    try:
        # Ensure a suitable font file (e.g., Arial) is available
        # You might need to adjust the path or install the font
        font_path = "arial.ttf" # Or specify a full path if needed
        font_size = 100 # Adjust as needed
        font = ImageFont.truetype(font_path, font_size)
        print(f"Loaded font: {font_path}")
    except IOError:
        print(f"Warning: Font file not found at '{font_path}'. Using default PIL font.")
        try:
            font = ImageFont.load_default() # Fallback to default font
        except Exception as font_e:
            print(f"Error: Could not load default font: {font_e}. Cannot add names to images.")
            font = None # Set font to None if default also fails
    text_color = (0, 0, 0) # Black text color

    # Load the template image
    try:
        template = Image.open(template_path).convert("RGB") # Ensure RGB for drawing color text
        template_width, template_height = template.size
        print(f"Loaded template image: {template_path} ({template_width}x{template_height})")
    except Exception as e:
        print(f"Error loading template image: {e}")
        return False
    
    # Check how to process the QR codes
    processed_count = 0
    skipped_existing = 0 # Counter for existing designed QRs
    skipped_not_found = 0 # Counter for missing basic QRs
    skipped_missing_name = 0 # Counter for missing names in CSV

    if csv_path and os.path.exists(csv_path):
        # Process based on CSV data
        df = pd.read_csv(csv_path, dtype=str)
        print(f"Processing {len(df)} records from CSV for QR design...")
        for index, row in tqdm(df.iterrows(), total=len(df), desc="Designing QR codes (CSV)"):
            mobile = row.get('mobile', '').strip()
            participant_name = row.get('isim', '').strip() # Get participant name

            if not mobile:
                continue # Skip rows with no mobile number

            # Construct the expected output path first
            output_path = os.path.join(output_dir, f"{mobile}_designed.png")

            # Check if the designed QR code file already exists
            if os.path.exists(output_path):
                skipped_existing += 1
                continue

            # Find the basic QR code file
            qr_file_path = os.path.join(qr_dir, f"{mobile}.png")
            if not os.path.exists(qr_file_path):
                skipped_not_found += 1
                continue
            
            # Create a copy of the template for each QR code
            new_img = template.copy()
            draw = ImageDraw.Draw(new_img) # Create Draw object
            
            # Load QR code
            try:
                qr_img = Image.open(qr_file_path)
            except Exception as img_e:
                print(f"Error opening basic QR image {qr_file_path} for phone {mobile}: {img_e}")
                continue # Skip this QR if it can't be opened

            # Calculate position to center the QR code on the blue area
            qr_size = 1500 # Size of the QR code to be pasted
            qr_x_position = (template_width - qr_size) // 2
            qr_y_position = (template_height - qr_size) // 2
            
            try:
                qr_img_resized = qr_img.resize((qr_size, qr_size))
            except Exception as resize_e:
                print(f"Error resizing QR image {qr_file_path} for phone {mobile}: {resize_e}")
                continue # Skip if resizing fails

            # Paste the QR code onto the template
            new_img.paste(qr_img_resized, (qr_x_position, qr_y_position))

            # --- Add Participant Name ---
            if participant_name and font:
                # Format name using Turkish-aware capitalization
                formatted_name = turkish_capitalize_name(participant_name)

                # Calculate text position
                try:
                    # Use textbbox for more accurate width calculation if possible (newer PIL)
                    # Ensure text is treated as string for bbox calculation
                    bbox = draw.textbbox((0, 0), str(formatted_name), font=font)
                    text_width = bbox[2] - bbox[0]
                    # text_height = bbox[3] - bbox[1] # Not needed for centering x
                except AttributeError:
                    # Fallback for older PIL versions
                    # Ensure text is treated as string for textsize calculation
                    text_width, _ = draw.textsize(str(formatted_name), font=font)
                except Exception as bbox_e:
                     print(f"Warning: Could not calculate text dimensions for '{formatted_name}' ({mobile}): {bbox_e}. Skipping text placement.")
                     text_width = None # Indicate failure

                if text_width is not None: # Proceed only if width calculation was successful
                    text_x_position = (template_width - text_width) / 2
                    # Position text below the QR code area
                    # QR bottom edge is at qr_y_position + qr_size
                    text_y_position = qr_y_position + qr_size + 550 # Adjust vertical spacing if needed

                    # Draw the text
                    try:
                        # Ensure text is treated as string for drawing
                        draw.text((text_x_position, text_y_position), str(formatted_name), fill=text_color, font=font)
                    except Exception as draw_e:
                        print(f"Error drawing text for '{formatted_name}' ({mobile}): {draw_e}")
            elif not participant_name:
                skipped_missing_name += 1
            # No 'else' needed if font failed to load, as 'font' would be None

            # Save the result
            try:
                new_img.save(output_path)
                processed_count += 1
            except Exception as save_e:
                 print(f"Error saving designed QR image {output_path} for phone {mobile}: {save_e}")

    else:
        # Process all QR code files in the directory (Fallback if no CSV)
        # Note: Participant names cannot be added in this mode
        print("Processing QR files directly from directory (CSV not provided)...")
        print("Warning: Participant names will not be added to images in this mode.")
        qr_files = [f for f in os.listdir(qr_dir) if f.endswith('.png')]
        for qr_file in tqdm(qr_files, desc="Designing QR codes (Dir)"):
            qr_file_path = os.path.join(qr_dir, qr_file)
            base_name = os.path.splitext(qr_file)[0] # Usually the mobile number

            # Construct the expected output path
            output_filename = f"{base_name}_designed.png"
            output_path = os.path.join(output_dir, output_filename)

            # Check if the designed QR code file already exists
            if os.path.exists(output_path):
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
            qr_size = 1500
            x_position = (template_width - qr_size) // 2
            y_position = (template_height - qr_size) // 2
            
            # Resize QR code to fit the blue area (1500x1500)
            try:
                qr_img_resized = qr_img.resize((qr_size, qr_size))
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
        print(f" - Skipped adding name (missing in CSV): {skipped_missing_name}")
    print(f" - Output directory: '{output_dir}'")
    return True
