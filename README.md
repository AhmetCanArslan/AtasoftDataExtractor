# DataExtractor Project

**Note:** This project was developed in response to receiving an overwhelming number of applications through the AI Summit Erzurum application forms, and as a solution for QR code–based entry.

## About the Project
DataExtractor processes data from an Excel file and generates QR codes. This project performs the following operations:
- Reading data from an Excel file
- Generating UUID based on the phone number
- Removing unnecessary columns and reordering columns
- Saving the processed data as a CSV file
- Generating QR codes from the CSV data
- Creating an Excel file containing the QR codes
- **Uploading processed data to Firebase**

## Technologies Used
- **Python**: Programming language
- **pandas**: Data processing and analysis
- **openpyxl**: Reading/writing Excel files and creating Excel files with QR codes
- **qrcode**: QR code generation
- **Pillow (PIL)**: Image processing
- **Firebase**: Cloud database for storing processed data

## Project Structure
```
DataExtractor/
├── DataExtractor.py      # Main application file
├── requirements.txt      # List of required packages
├── .gitignore            # Git ignore configuration
├── .gitattributes        # Git attributes configuration
└── README.md             # (This) Project documentation
```

## Requirements
To install the required Python packages for the project, run the following command:
```bash
pip install -r requirements.txt
```

## Firebase Setup
1. **Create a Firebase Project**:  
   Go to the [Firebase Console](https://console.firebase.google.com/) and create a new project.

2. **Enable Firestore Database**:  
   In the Firebase Console, navigate to the Firestore Database section and enable it in "Production Mode."

3. **Generate Service Account Key**:  
   - Go to the "Project Settings" in the Firebase Console.
   - Under the "Service Accounts" tab, generate a new private key and download the JSON file.
   - Save the JSON file in the project root directory as `firebase_credentials.json`.

4. **Install Firebase Admin SDK**:  
   Add the Firebase Admin SDK to the `requirements.txt` file and install it:
   ```bash
   pip install firebase-admin
   ```

## Usage Instructions
1. **Prepare the Excel File**: An Excel file named `yanitlar.xlsx` should be located in the project root. This file must contain at least the columns "Telefon numaranız", "Ad-Soyad", and "E-posta adresiniz."
2. **Run the Python Script**: Execute the following command from the project directory using a terminal or CMD:
   ```bash
   python DataExtractor.py
   ```
3. **Check the Outputs**:
   - **CSV File**: Generated under the `output/csv/` directory.
   - **QR Codes**: PNG files created in the `output/qr/` directory.
   - **QR Code Excel File**: The Excel file `output/excel/qrKodlar.xlsx` contains the data along with QR codes.
   - **Firebase Database**: Processed data is uploaded to the Firestore database.

## Functions and Workflow
- **process_excel()**:  
  Reads the Excel file, generates a UUID based on the phone number column, removes unnecessary columns, and saves the data as a CSV file.

- **generate_qr_codes_from_csv()**:  
  Reads the CSV file, cleans the phone number to create dynamic filenames, and generates a QR code for each record.

- **generate_excel_with_qr()**:  
  Uses the CSV data to create a new Excel file. This file embeds the corresponding QR codes in the relevant cells.

- **upload_to_firebase()**:  
  Uploads the processed data to the Firestore database using the Firebase Admin SDK.

## Important Notes
- The `clean_phone_number()` function is used to obtain a clean phone number.
- UUIDs are generated using `uuid5` based on the phone number to ensure uniqueness.
- Warnings and error messages are displayed if any of the unnecessary columns are missing.
- Output directories are created automatically when the script is executed.
- Ensure the `firebase_credentials.json` file is correctly configured for Firebase integration.

## Debugging and Logs
- Functions log descriptive error messages during failures.
- Detailed logs are provided during QR code generation, file writing errors, or Firebase upload issues.

## Contribution
This project is open source. Contributions are welcome via pull requests or by opening issues.

## License
This project is licensed under the MIT License.
