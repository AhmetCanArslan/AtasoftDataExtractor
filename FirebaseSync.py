import firebase_admin
from firebase_admin import credentials, firestore
import pandas as pd
import sys
import os

# --- Configuration ---
# IMPORTANT: Replace with the actual path to your Firebase service account key file
SERVICE_ACCOUNT_KEY_PATH = 'qr-deneme.json'
COLLECTION_NAME = 'users'
CSV_UUID_COL = 'UUID'
CSV_COUNTER_COL = 'Counter'
CSV_NAME_COL = 'isim'
CSV_EMAIL_COL = 'mail'
CSV_PHONE_COL = 'mobile'

# --- Firebase Initialization ---
def initialize_firebase():
    """Initializes the Firebase Admin SDK."""
    if not os.path.exists(SERVICE_ACCOUNT_KEY_PATH):
        print(f"Error: Service account key file not found at '{SERVICE_ACCOUNT_KEY_PATH}'")
        print("Please update SERVICE_ACCOUNT_KEY_PATH in FirebaseSync.py")
        sys.exit(1)
    try:
        cred = credentials.Certificate(SERVICE_ACCOUNT_KEY_PATH)
        firebase_admin.initialize_app(cred)
        print("Firebase Admin SDK initialized successfully.")
        return firestore.client()
    except Exception as e:
        print(f"Error initializing Firebase: {e}")
        sys.exit(1)

# --- Firestore Operations ---
def delete_collection(db, collection_ref, batch_size=500):
    """Deletes all documents in a collection in batches."""
    try:
        docs = collection_ref.limit(batch_size).stream()
        deleted = 0
        while True:
            batch = db.batch()
            doc_count = 0
            for doc in docs:
                batch.delete(doc.reference)
                doc_count += 1
            if doc_count == 0:
                break # No more documents to delete
            batch.commit()
            deleted += doc_count
            print(f"Deleted {deleted} documents...")
            # Get the next batch
            docs = collection_ref.limit(batch_size).stream()

        print(f"Successfully deleted all documents in '{collection_ref.id}' collection.")
    except Exception as e:
        print(f"Error deleting collection '{collection_ref.id}': {e}")
        sys.exit(1)

def sync_csv_to_firestore(db, csv_path):
    """Reads CSV and uploads data to Firestore, replacing existing data."""
    if not os.path.exists(csv_path):
        print(f"Error: CSV file not found at '{csv_path}'")
        sys.exit(1)

    try:
        df = pd.read_csv(csv_path, dtype={CSV_UUID_COL: str, CSV_COUNTER_COL: str, CSV_PHONE_COL: str})
        print(f"Read {len(df)} rows from '{csv_path}'.")
    except Exception as e:
        print(f"Error reading CSV file '{csv_path}': {e}")
        sys.exit(1)

    users_ref = db.collection(COLLECTION_NAME)

    # 1. Delete existing data
    print(f"Attempting to delete all existing documents in '{COLLECTION_NAME}' collection...")
    delete_collection(db, users_ref)

    # 2. Upload new data
    print(f"Uploading {len(df)} records to '{COLLECTION_NAME}' collection...")
    upload_count = 0
    batch = db.batch()
    batch_count = 0

    for index, row in df.iterrows():
        doc_id = row[CSV_UUID_COL]
        if not doc_id or pd.isna(doc_id):
            print(f"Warning: Skipping row {index + 2} due to missing or invalid UUID.")
            continue

        try:
            # Ensure Counter is an integer, default to 0 if conversion fails or is NaN/None
            counter_val = row[CSV_COUNTER_COL]
            if pd.isna(counter_val):
                counter = 0
            else:
                try:
                    counter = int(float(counter_val)) # Use float first for potential decimals like '0.0'
                except (ValueError, TypeError):
                    print(f"Warning: Could not convert Counter '{counter_val}' to int for UUID {doc_id}. Defaulting to 0.")
                    counter = 0

            data = {
                "Counter": counter,
                "Ad-Soyad": row[CSV_NAME_COL],
                "Eposta": row[CSV_EMAIL_COL],
                "Telefon numaranız": row[CSV_PHONE_COL]
            }
            doc_ref = users_ref.document(doc_id)
            batch.set(doc_ref, data)
            batch_count += 1
            upload_count += 1

            # Commit batch periodically to avoid exceeding limits (Firestore batch limit is 500 operations)
            if batch_count >= 499:
                print(f"Committing batch of {batch_count} documents...")
                batch.commit()
                print("Batch committed.")
                batch = db.batch() # Start a new batch
                batch_count = 0

        except Exception as e:
            print(f"Error processing row {index + 2} (UUID: {doc_id}): {e}")

    # Commit any remaining documents in the last batch
    if batch_count > 0:
        print(f"Committing final batch of {batch_count} documents...")
        batch.commit()
        print("Final batch committed.")

    print(f"Successfully uploaded {upload_count} records to Firestore collection '{COLLECTION_NAME}'.")


# --- Main Execution ---
if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python FirebaseSync.py <path_to_csv_file>")
        sys.exit(1)

    csv_file_path = sys.argv[1]

    print("--- Starting Firebase Synchronization ---")
    db_client = initialize_firebase()
    sync_csv_to_firestore(db_client, csv_file_path)
    print("--- Firebase Synchronization Complete ---")
