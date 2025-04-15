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
├── MailSender.py           # Sends emails with designed QR codes
├── requirements.txt        # List of required Python packages
├── .env                    # Environment variables (file paths, credentials) - **DO NOT COMMIT**
├── .gitignore              # Git ignore configuration
├── .gitattributes          # Git attributes configuration
├── LICENSE                 # Project license file
├── qr-deneme.json          # Firebase service account key - **DO NOT COMMIT**
├── tasarim.jpg             # Template image for QR code design (example name)
├── yanitlar.xlsx           # Input Excel file (example name)
├── output/                 # Directory for generated files
│   ├── csv/                # Output CSV files
│   ├── qr/                 # Output basic QR code images (.png)
│   ├── designed_qr/        # Output designed QR code images (.png)
│   └── excel/              # Output Excel file with basic QR codes
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
# File Paths
EXCEL_FILE_PATH=yanitlar.xlsx
PHONE_COLUMN_NAME=Telefon numaranız # Column name in Excel for phone numbers
UUID_COLUMN_NAME=UUID
COUNTER_COLUMN_NAME=Counter
CSV_OUTPUT_DIR=output/csv
QR_OUTPUT_DIR=output/qr
EXCEL_OUTPUT_PATH=output/excel/qrKodlar.xlsx
FIREBASE_SYNC_SCRIPT=FirebaseSync.py

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
    *   Ensure the Excel file (e.g., `yanitlar.xlsx`) is in the project root directory and contains the necessary columns (specified in `.env`).
    *   Ensure the QR design template image (e.g., `tasarim.jpg`) is in the project root directory.
    *   Ensure the Firebase service account key (`qr-deneme.json`) is in the project root directory.
2.  **Configure Environment**: Create and populate the `.env` file as described in the "Configuration" section.
3.  **Run the Main Script**: Execute the following command from the project directory:

    ```bash
    python DataExtractor.py
    ```
4.  **Follow Prompts**: The script will:
    *   Process the Excel file, generate CSV, basic QR codes, and the Excel output.
    *   Attempt to design the QR codes using the template.
    *   Ask if you want to sync the data with Firebase. Enter `yes` or `no`.
    *   Ask if you want to send emails with the designed QR codes. Enter `yes` or `no`.
5.  **Check the Outputs**:
    *   **CSV File**: Generated in the directory specified by `CSV_OUTPUT_DIR` in `.env`.
    *   **Basic QR Codes**: PNG files created in the directory specified by `QR_OUTPUT_DIR`.
    *   **Designed QR Codes**: PNG files created in `output/designed_qr/`.
    *   **QR Code Excel File**: The Excel file generated at the path specified by `EXCEL_OUTPUT_PATH`.
    *   **Firebase Database**: If sync was chosen, data is uploaded/updated in the Firestore `users` collection.
    *   **Emails**: If chosen, emails are sent to the addresses in the CSV, with designed QR codes attached. Check the console output for success/error messages.

## Functions and Workflow
- **`DataExtractor.py`**: Orchestrates the entire workflow.
    - Calls `process_excel` to read Excel, clean data, generate UUIDs, and save CSV.
    - Calls `generate_qr_codes_from_csv` (from `QRGenerator.py`) to create basic QR images.
    - Calls `generate_excel_with_qr` to create the output Excel file.
    - Calls `overlay_qr_on_template` (from `QRDesign.py`) to create designed QR images.
    - Prompts the user and conditionally runs the `FirebaseSync.py` script via `subprocess`.
    - Prompts the user and conditionally calls `send_qr_codes` (from `MailSender.py`).
- **`FileOperations.py`**:
    - `read_excel()`: Reads the input Excel.
    - `clean_phone_number()`: Standardizes phone numbers.
    - `save_csv()`: Saves the processed DataFrame to CSV.
    - `create_directory_if_not_exists()`: Utility for creating output folders.
- **`QRGenerator.py`**:
    - `generate_qr_codes_from_csv()`: Creates individual QR code PNG files from CSV data (UUID).
- **`QRDesign.py`**:
    - `overlay_qr_on_template()`: Loads a template image, resizes QR codes, and pastes them onto the template, saving the results.
- **`FirebaseSync.py`**:
    - Initializes Firebase Admin SDK.
    - `delete_collection()`: Clears the target Firestore collection before uploading.
    - `sync_csv_to_firestore()`: Reads the CSV and uploads data to Firestore in batches.
- **`MailSender.py`**:
    - `send_qr_codes()`: Reads the CSV, connects to the SMTP server, formats emails, attaches the corresponding *designed* QR code, and sends emails individually.

## Important Notes
- The `clean_phone_number()` function in `FileOperations.py` standardizes phone numbers to 10 digits (removing country codes like +90, 90, or leading 0) before UUID generation. Invalid numbers result in skipped rows.
- UUIDs are generated using `uuid5` with the `NAMESPACE_DNS` and the cleaned phone number to ensure consistency.
- Configure all paths and credentials securely in the `.env` file. **Never commit `.env` or your `qr-deneme.json` file to Git.**
- Ensure the template image (`tasarim.jpg` or your specified name) exists for the QR design step.
- For Gmail, you might need to enable "Less secure app access" or preferably generate an "App Password" if you have 2-Factor Authentication enabled.
- Output directories (`output/csv`, `output/qr`, etc.) are created automatically if they don't exist.

## Debugging and Logs
- The scripts print status messages, warnings, and errors to the console during execution.
- Check console output for details on file processing, QR generation, Firebase sync status, and email sending results.

## Contribution
This project is open source. Contributions are welcome via pull requests or by opening issues. 

## License
This project is licensed under the MIT License. See the `LICENSE` file for details.
