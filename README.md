# DataExtractor Project

**Note:** This project was developed in response to receiving an overwhelming number of applications through the AI Summit Erzurum application forms. It serves as a solution for QR code–based entry, design, and distribution, automating the process and enabling the functionality of the associated mobile application by synchronizing participant data.

## About the Project
DataExtractor processes data from an Excel file, generates unique identifiers and QR codes, designs the QR codes by overlaying them onto a template, optionally syncs data with Firebase Firestore, and sends the designed QR codes via email to participants. This project performs the following operations:
- Reading data from an Excel file (`.xlsx`).
- Cleaning phone numbers and generating unique UUIDs based on them.
- Removing unnecessary columns and reordering remaining columns.
- Saving the processed data as a CSV file.
- Generating QR codes based on the UUIDs.
- Designing QR codes by overlaying them onto a template image.
- Creating an Excel file containing participant data and basic QR codes.
- **Optionally:** Uploading processed data (UUID, Name, Email, Phone, Counter) to Firebase Firestore, replacing existing data.
- **Optionally:** Sending personalized emails with the designed QR codes attached.

## Technologies Used
- **Python**: Programming language
- **pandas**: Data processing and analysis
- **openpyxl**: Reading/writing Excel files
- **qrcode**: QR code generation
- **Pillow (PIL)**: Image processing (QR generation, overlaying on template)
- **firebase-admin**: Interacting with Firebase Firestore
- **python-dotenv**: Managing environment variables (API keys, file paths)
- **smtplib, email**: Sending emails with attachments

## Project Structure
```
DataExtractor/
├── DataExtractor.py        # Main script orchestrating the process
├── FileOperations.py       # Handles file reading (Excel), saving (CSV), directory creation, phone cleaning
├── QRGenerator.py          # Generates basic QR code images from CSV data
├── QRDesign.py             # Overlays QR codes onto a template image
├── FirebaseSync.py         # Handles synchronization with Firebase Firestore
├── DeleteFirebaseCollection.py # Utility to clear the Firestore collection
├── MailSender.py           # Sends emails with designed QR codes
├── CertificateGeneratorSender.py # Generates and sends attendance certificates
├── requirements.txt        # List of required Python packages
├── .env                    # Environment variables (file paths, credentials) - **DO NOT COMMIT**
├── .gitignore              # Git ignore configuration
├── .gitattributes          # Git attributes configuration
├── LICENSE                 # Project license file
├── qr-deneme.json          # Firebase service account key - **DO NOT COMMIT**
├── tasarim.jpg             # Template image for QR code design (example name)
├── input/                  # Directory for input files
│   └── (Place your .xlsx file here)
├── output/                 # Directory for generated files
│   ├── csv/                # Output CSV files (e.g., input_file_form.csv, input_file_clean.csv)
│   ├── qr/                 # Output basic QR code images (.png)
│   ├── designed_qr/        # Output designed QR code images (.png)
│   ├── excel/              # Output Excel file with basic QR codes (e.g., input_file_modified.xlsx)
│   └── certificates/       # Output certificate images (.png)
├── logs/                   # Directory for log files
│   ├── sent_emails.csv     # Tracks successfully sent QR code emails
│   └── email_errors.csv    # Tracks email sending errors and skips
└── README.md               # (This) Project documentation
```

## Requirements
To install the required Python packages for the project, run the following command:
```bash
pip install -r requirements.txt
```

## Configuration (`.env` file)
Create a file named `.env` in the root directory of the project and add the following variables. Replace the example values with your actual configuration. **Do not commit this file to version control.**

```dotenv
# File Paths (Input Excel path is now handled automatically from the 'input' directory)
# EXCEL_FILE_PATH=yanitlar.xlsx # No longer needed here
PHONE_COLUMN_NAME=Telefon numaranız # Column name in Excel for phone numbers
UUID_COLUMN_NAME=UUID
COUNTER_COLUMN_NAME=Counter
CSV_OUTPUT_DIR=output/csv
QR_OUTPUT_DIR=output/qr
# EXCEL_OUTPUT_PATH=output/excel/qrKodlar.xlsx # Removed: Filename is now dynamic (e.g., input_file_modified.xlsx)
FIREBASE_SYNC_SCRIPT=FirebaseSync.py
# Add path for the delete script if needed, though it's typically run manually
# FIREBASE_DELETE_SCRIPT=DeleteFirebaseCollection.py

# Email Configuration (for Gmail example)
SENDER_EMAIL=your_email@gmail.com
SENDER_PASSWORD=your_gmail_app_password # Use an App Password if 2FA is enabled
SMTP_SERVER=smtp.gmail.com
SMTP_PORT=587

# QR Code Design Image (used by QRDesign.py) 
# Ensure the template image (e.g., tasarim.jpg) exists in the root directory
```
Special thanks to [a0pia](https://github.com/a0pia) for qr design.

## Firebase Setup
1.  **Create a Firebase Project**:
    Go to the [Firebase Console](https://console.firebase.google.com/) and create a new project.

2.  **Enable Firestore Database**:
    In the Firebase Console, navigate to the Firestore Database section and enable it (choose "Start in production mode" or "Start in test mode" as needed). Note the security rules.
3.  **Generate Service Account Key**:
    - Go to "Project settings" > "Service accounts" in the Firebase Console.
    - Select "Python" and click "Generate new private key".
    - Download the JSON file.
    - **Rename the downloaded file to `qr-deneme.json`** and place it in the project's root directory. **Ensure this file is listed in your `.gitignore` file.**
4.  **Firebase Admin SDK**: The SDK is listed in `requirements.txt` and will be installed with `pip install -r requirements.txt`.

## Usage Instructions
1.  **Prepare Input Files**:
    *   Place the **single** Excel file (`.xlsx`) containing participant data into the `input/` directory. Ensure it contains the necessary columns (like the phone number column specified in `.env`).
    *   Ensure the QR design template image (e.g., `tasarim.jpg`) is in the project root directory.
    *   Ensure the Firebase service account key (`qr-deneme.json`) is in the project root directory.
    *   For certificate generation: Ensure the certificate template image (e.g., `tasarim.jpg` or another specified file) and the font file (e.g., `arial.ttf`) are in the project root directory.
2.  **Configure Environment**: Create and populate the `.env` file as described in the "Configuration" section. Make sure `PHONE_COLUMN_NAME` matches the column header in your Excel file.

3.  **Run the Main Script**: Execute the following command from the project directory:

    ```bash
    python DataExtractor.py
    ```
4.  **Follow Prompts**: The script will:
    *   Process the Excel file, generate intermediate and final CSV files, basic QR codes, and the Excel output.
    *   Attempt to design the QR codes using the template.
    *   Ask if you want to sync the data with Firebase (updates/adds records). Enter `yes` or `no`.
    *   Ask if you want to send emails with the designed QR codes. Enter `yes` or `no`.
5.  **Check the Outputs**:
    *   **Intermediate CSV File**: An intermediate CSV (`*_form.csv`) is saved in the directory specified by `CSV_OUTPUT_DIR` before duplicate removal.
    *   **Final CSV File**: The final processed CSV file (`*_clean.csv`) is generated in the directory specified by `CSV_OUTPUT_DIR` (e.g., `output/csv/your_input_file_clean.csv`) after cleaning, deduplication, and column removal. This file is used for subsequent steps (QR generation, Excel, Firebase sync, Email sending).
    *   **Basic QR Codes**: PNG files created in the directory specified by `QR_OUTPUT_DIR`.
    *   **Designed QR Codes**: PNG files created in `output/designed_qr/`.
    *   **QR Code Excel File**: An Excel file generated in `output/excel/` with a name based on your input file (e.g., `your_input_file_modified.xlsx`).
    *   **Firebase Database**: If sync was chosen, data from `*_clean.csv` is uploaded/updated (merged) in the Firestore `users` collection.
    *   **Emails**: If chosen, emails are sent to the addresses in `*_clean.csv`, with designed QR codes attached. Check the console output and `logs/` files for success/error messages.
5.  **(Optional) Delete Firebase Collection**: If you need to clear the Firestore `users` collection completely, run:
    ```bash
    python DeleteFirebaseCollection.py
    ```
    You will be asked for confirmation before deletion occurs.
6.  **(Optional) Generate and Send Certificates**: To generate certificates for attendees (users with `Counter > 0` in Firestore) and email them, run:
    ```bash
    python CertificateGeneratorSender.py
    ```
    Check the `output/certificates/` directory and console output.

## Functions and Workflow
- **`DataExtractor.py`**: Orchestrates the main QR generation and distribution workflow.
    - Finds the input Excel file in the `input/` directory.
    - Calls `process_excel` to read Excel, clean data, generate UUIDs, save an intermediate `*_form.csv`, remove duplicates and specified columns (including "Doğum tarihiniz..."), and save the final `*_clean.csv`.
    - Calls `generate_qr_codes_from_csv` (from `QRGenerator.py`) using `*_clean.csv` to create basic QR images.
    - Constructs the output Excel filename dynamically (e.g., `input_file_modified.xlsx`).
    - Calls `generate_excel_with_qr` using `*_clean.csv` to create the output Excel file in `output/excel/`.
    - Calls `overlay_qr_on_template` (from `QRDesign.py`) using `*_clean.csv` to create designed QR images.
    - Prompts the user and conditionally runs the `FirebaseSync.py` script via `subprocess`, passing `*_clean.csv`.
    - Prompts the user and conditionally calls `send_qr_codes` (from `MailSender.py`) using `*_clean.csv`.
- **`FileOperations.py`**:
    - `read_excel()`: Reads the input Excel.
    - `clean_phone_number()`: Standardizes phone numbers.
    - `save_csv()`: Saves the processed DataFrame to CSV (Note: `DataExtractor.py` now handles the specific naming logic).
    - `create_directory_if_not_exists()`: Utility for creating output folders.
    - `save_excel_with_qr()`: Saves the DataFrame with embedded QR images to the specified Excel path.
- **`QRGenerator.py`**:
    - `generate_qr_codes_from_csv()`: Creates individual QR code PNG files from CSV data (UUID).
- **`QRDesign.py`**:
    - `overlay_qr_on_template()`: Loads a template image, resizes QR codes, and pastes them onto the template, saving the results.
- **`FirebaseSync.py`**:
    - Initializes Firebase Admin SDK.
    - `initialize_firebase_sync()`: Handles SDK initialization.
    - `delete_collection()`: (Used by `DeleteFirebaseCollection.py`) Clears the target Firestore collection.
    - `sync_csv_to_firestore()`: Reads the CSV and uploads/updates data to Firestore in batches using `merge=True`.
- **`DeleteFirebaseCollection.py`**: Standalone script to clear the Firestore collection after confirmation.
- **`CertificateGeneratorSender.py`**: Standalone script to fetch attendees (Counter > 0) from Firestore, generate certificates, and send them via email.
- **`MailSender.py`**:
    - `send_qr_codes()`: Reads the CSV, connects to the SMTP server, formats emails, attaches the corresponding *designed* QR code, and sends emails individually.

## Important Notes
- The script now automatically looks for a single `.xlsx` file in the `input/` directory. Ensure only one Excel file is present there.
- The `FirebaseSync.py` script (run via `DataExtractor.py`) now **updates or adds** records to Firestore based on UUID, rather than overwriting the collection. Use `DeleteFirebaseCollection.py` if you need to clear the collection first.
- The `clean_phone_number()` function in `FileOperations.py` standardizes phone numbers before UUID generation. Invalid numbers result in skipped rows.
- UUIDs are generated using `uuid5` with the `NAMESPACE_DNS` and the cleaned phone number to ensure consistency.
- Configure all paths and credentials securely in the `.env` file. **Never commit `.env` or your `qr-deneme.json` file to Git.**
- Ensure the template image (`tasarim.jpg` or your specified name) exists for the QR design step.
- For Gmail, you might need to enable "Less secure app access" or preferably generate an "App Password" if you have 2-Factor Authentication enabled.
- Output directories (`input/`, `output/csv`, `output/qr`, etc.) are created automatically if they don't exist.
- The `CertificateGeneratorSender.py` script runs independently and relies on data already present in Firestore (specifically users with `Counter > 0`).

## Debugging and Logs
- The scripts print status messages, warnings, and errors to the console during execution.
- Check console output for details on file processing, QR generation, Firebase sync status (updates/additions), email sending results, and certificate processing.
- **`logs/sent_emails.csv`**: Records the email, mobile number, and timestamp for each successfully sent QR code email to prevent duplicates.
- **`logs/email_errors.csv`**: Records the email, mobile number (if available), reason, and timestamp for emails that failed to send or were skipped (e.g., missing data, missing QR file, SMTP error). This helps diagnose issues with email delivery.

## Contribution
This project is open source. Contributions are welcome via pull requests or by opening issues. 

## License
This project is licensed under the MIT License. See the `LICENSE` file for details.
