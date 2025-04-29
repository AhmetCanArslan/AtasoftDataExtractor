import sys
import os
import firebase_admin
from firebase_admin import firestore

# Import necessary functions from FirebaseSync.py
# Ensure FirebaseSync.py is in the same directory or Python path
try:
    from FirebaseSync import initialize_firebase_sync, delete_collection, COLLECTION_NAME
except ImportError:
    print("Error: Could not import functions from FirebaseSync.py.")
    print("Ensure FirebaseSync.py is in the same directory.")
    sys.exit(1)

if __name__ == "__main__":
    print("--- Firebase Collection Deletion Utility ---")

    # Initialize Firebase
    db_client = initialize_firebase_sync()
    if not db_client:
        print("Exiting due to Firebase initialization failure.")
        sys.exit(1)

    collection_ref = db_client.collection(COLLECTION_NAME)

    print(f"\nWARNING: This script will permanently delete all documents")
    print(f"within the '{COLLECTION_NAME}' collection in your Firestore database.")
    print("This action cannot be undone.")

    confirm = input(f"Are you sure you want to delete the entire '{COLLECTION_NAME}' collection? (yes/no): ").strip().lower()

    if confirm == 'yes':
        print(f"\nProceeding with deletion of collection '{COLLECTION_NAME}'...")
        delete_collection(db_client, collection_ref)
        print(f"--- Deletion process for '{COLLECTION_NAME}' finished. ---")
    else:
        print("\nDeletion cancelled by user.")
        print("--- Exiting without deleting collection. ---")

