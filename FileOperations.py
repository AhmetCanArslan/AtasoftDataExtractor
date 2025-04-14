import os
import pandas as pd
from openpyxl import Workbook

def create_directory_if_not_exists(directory_path):
    """
    Create a directory if it does not already exist.

    Args:
        directory_path (str): Path to the directory to create.
    """
    if not os.path.exists(directory_path):
        os.makedirs(directory_path)
        print(f"Created directory: {directory_path}")

def read_excel(file_path, phone_col):
    """
    Read an Excel file and return a DataFrame with the specified phone column as string.

    Args:
        file_path (str): Path to the Excel file.
        phone_col (str): Column name for phone numbers.

    Returns:
        pd.DataFrame: DataFrame with the Excel data.
    """
    if not os.path.exists(file_path):
        print(f"Error: File not found at '{file_path}'")
        return None

    df = pd.read_excel(file_path, engine='openpyxl', dtype={phone_col: str})
    print(f"Successfully read {len(df)} rows from '{file_path}'.")
    return df

def save_csv(df, output_dir, file_name):
    """
    Save a DataFrame to a CSV file.

    Args:
        df (pd.DataFrame): DataFrame to save.
        output_dir (str): Directory to save the CSV file.
        file_name (str): Name of the CSV file.
    """
    create_directory_if_not_exists(output_dir)
    output_csv_path = os.path.join(output_dir, file_name)
    df.to_csv(output_csv_path, index=False, encoding='utf-8-sig')
    print(f"Successfully saved CSV to '{output_csv_path}'.")

def save_excel_with_qr(df, qr_dir, excel_output_path, uuid_column):
    """
    Save a DataFrame to an Excel file with QR codes.

    Args:
        df (pd.DataFrame): DataFrame to save.
        qr_dir (str): Directory containing QR code images.
        excel_output_path (str): Path to save the Excel file.
        uuid_column (str): Column name for UUIDs.
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "QR Kodlar"
    ws.column_dimensions['A'].width = 20  # QR için sütun genişliği
    # Başlık satırı oluştur
    headers = ['qr kodlar', 'ad soyad', 'posta', 'numara']
    ws.append(headers)
    row_number = 2
    for _, row in df.iterrows():
        ws.row_dimensions[row_number].height = 100  # Satır yüksekliğini QR boyutuna göre ayarla
        uuid_value = row.get(uuid_column, '')
        img_path = os.path.join(qr_dir, f"{uuid_value}.png")
        # Diğer sütunlar
        ws.cell(row=row_number, column=2, value=row.get('isim', ''))
        ws.cell(row=row_number, column=3, value=row.get('mail', ''))
        ws.cell(row=row_number, column=4, value=row.get('mobile', ''))
        # QR görselini ekle (varsa)
        if os.path.exists(img_path):
            from openpyxl.drawing.image import Image as OpenpyxlImage
            img = OpenpyxlImage(img_path)
            img.width = 145
            img.height = 145
            ws.add_image(img, f"A{row_number}")
        else:
            ws.cell(row=row_number, column=1, value="No QR")
        row_number += 1

    # Diğer sütun genişliklerini ayarla
    for col in ['B', 'C', 'D']:
        max_length = 0
        for cell in ws[col]:
            if cell.value:
                max_length = max(max_length, len(str(cell.value)))
        ws.column_dimensions[col].width = max_length + 5  # padding ekle

    excel_dir = os.path.dirname(excel_output_path)
    create_directory_if_not_exists(excel_dir)
    wb.save(excel_output_path)
    print(f"Successfully saved Excel to '{excel_output_path}'.")