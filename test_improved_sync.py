#!/usr/bin/env python3
"""
Test script for improved Firebase Storage sync functionality.
This script demonstrates the improved timestamp handling and metadata storage.
"""

import os
import time
import json
from datetime import datetime
from google.cloud import storage
from google.oauth2 import service_account

# Configuration
SERVICE_ACCOUNT_PATH = "MacroMouse_Data/spendingcache-personal-firebase-adminsdk-fbsvc-148f467967.json"
BUCKET_NAME = "spendingcache-personal.appspot.com"
TEST_FILE = "test_sync_data.json"
REMOTE_PATH = "macro-data/test_sync_data.json"

def initialize_firebase():
    """Initialize Firebase Storage client."""
    try:
        if not os.path.exists(SERVICE_ACCOUNT_PATH):
            print(f"❌ Service account file not found: {SERVICE_ACCOUNT_PATH}")
            return None
        
        # Set up credentials
        os.environ["GOOGLE_APPLICATION_CREDENTIALS"] = SERVICE_ACCOUNT_PATH
        
        # Create client
        client = storage.Client()
        bucket = client.bucket(BUCKET_NAME)
        
        # Test connection
        blobs = list(bucket.list_blobs(max_results=1))
        print(f"✅ Successfully connected to Firebase bucket: {BUCKET_NAME}")
        return bucket
        
    except Exception as e:
        print(f"❌ Failed to initialize Firebase: {str(e)}")
        return None

def get_local_timestamp(local_path):
    """Get local file modification timestamp."""
    if not os.path.exists(local_path):
        return 0
    return int(os.path.getmtime(local_path))

def get_remote_timestamp(bucket, firebase_path):
    """Get remote file timestamp from metadata or blob updated time."""
    try:
        blob = bucket.blob(firebase_path)
        
        # Try to get custom timestamp from metadata first
        if blob.exists():
            metadata = blob.metadata or {}
            custom_timestamp = metadata.get('last_modified')
            if custom_timestamp:
                return int(float(custom_timestamp))
            
            # Fall back to blob updated time
            return int(blob.updated.replace(tzinfo=None).timestamp())
        else:
            return 0
    except Exception as e:
        print(f"Error getting remote timestamp for {firebase_path}: {str(e)}")
        return 0

def upload_file_with_metadata(bucket, local_path, firebase_path):
    """Upload file with custom timestamp metadata."""
    try:
        blob = bucket.blob(firebase_path)
        
        # Set custom metadata with current timestamp
        current_timestamp = str(time.time())
        metadata = {
            'last_modified': current_timestamp,
            'uploaded_at': datetime.now().isoformat(),
            'file_size': str(os.path.getsize(local_path)),
            'test_file': 'true'
        }
        
        blob.metadata = metadata
        blob.upload_from_filename(local_path)
        
        print(f"✅ Uploaded {os.path.basename(local_path)} with timestamp {current_timestamp}")
        return True
    except Exception as e:
        print(f"❌ Upload failed for {os.path.basename(local_path)}: {str(e)}")
        return False

def download_file_with_metadata(bucket, firebase_path, local_path):
    """Download file and preserve metadata."""
    try:
        blob = bucket.blob(firebase_path)
        
        # Ensure directory exists
        os.makedirs(os.path.dirname(local_path), exist_ok=True)
        
        # Download the file
        blob.download_to_filename(local_path)
        
        # Update local file timestamp to match remote if possible
        if blob.metadata and 'last_modified' in blob.metadata:
            remote_timestamp = float(blob.metadata['last_modified'])
            os.utime(local_path, (remote_timestamp, remote_timestamp))
        
        print(f"✅ Downloaded {os.path.basename(local_path)}")
        return True
    except Exception as e:
        print(f"❌ Download failed for {os.path.basename(local_path)}: {str(e)}")
        return False

def sync_file(bucket, local_path, remote_path):
    """Sync a single file with improved timestamp handling."""
    print(f"\n🔄 Syncing {local_path} with gs://{bucket.name}/{remote_path}...")
    
    # Get timestamps
    local_timestamp = get_local_timestamp(local_path)
    remote_timestamp = get_remote_timestamp(bucket, remote_path)
    
    # Format timestamps for display
    local_time_str = datetime.fromtimestamp(local_timestamp).strftime('%Y-%m-%d %H:%M:%S') if local_timestamp > 0 else "N/A"
    remote_time_str = datetime.fromtimestamp(remote_timestamp).strftime('%Y-%m-%d %H:%M:%S') if remote_timestamp > 0 else "N/A"
    
    print(f"   Local timestamp:  {local_time_str}")
    print(f"   Remote timestamp: {remote_time_str}")
    
    # Compare timestamps and sync
    if local_timestamp > remote_timestamp:
        print("   📤 Local file is newer. Uploading...")
        if upload_file_with_metadata(bucket, local_path, remote_path):
            print("   ✅ Upload complete.")
        else:
            print("   ❌ Upload failed.")
            
    elif remote_timestamp > local_timestamp:
        print("   📥 Remote file is newer. Downloading...")
        if download_file_with_metadata(bucket, remote_path, local_path):
            print("   ✅ Download complete.")
        else:
            print("   ❌ Download failed.")
            
    elif local_timestamp == remote_timestamp and local_timestamp > 0:
        print("   ✅ Files are in sync. No action needed.")
        
    else:
        # One or both files don't exist
        if local_timestamp == 0 and remote_timestamp == 0:
            print("   ⚠️  File doesn't exist locally or remotely.")
        elif local_timestamp == 0:
            print("   📥 Downloading remote file (no local copy)...")
            if download_file_with_metadata(bucket, remote_path, local_path):
                print("   ✅ Download complete.")
            else:
                print("   ❌ Download failed.")
        else:
            print("   📤 Uploading local file (no remote copy)...")
            if upload_file_with_metadata(bucket, local_path, remote_path):
                print("   ✅ Upload complete.")
            else:
                print("   ❌ Upload failed.")

def create_test_file(filename, data):
    """Create a test file with sample data."""
    with open(filename, 'w') as f:
        json.dump(data, f, indent=2)
    print(f"📝 Created test file: {filename}")

def main():
    """Main test function."""
    print("🧪 Testing Improved Firebase Storage Sync")
    print("=" * 50)
    
    # Initialize Firebase
    bucket = initialize_firebase()
    if not bucket:
        return
    
    # Create test data
    test_data = {
        "test_id": f"test_{int(time.time())}",
        "created_at": datetime.now().isoformat(),
        "message": "This is a test file for improved sync functionality",
        "version": "1.0"
    }
    
    # Create local test file
    create_test_file(TEST_FILE, test_data)
    
    # Test 1: Initial upload (no remote file)
    print("\n🔍 Test 1: Initial upload")
    sync_file(bucket, TEST_FILE, REMOTE_PATH)
    
    # Wait a moment
    time.sleep(2)
    
    # Test 2: Modify local file and sync
    print("\n🔍 Test 2: Local file modification")
    test_data["message"] = "This file has been modified locally"
    test_data["modified_at"] = datetime.now().isoformat()
    create_test_file(TEST_FILE, test_data)
    sync_file(bucket, TEST_FILE, REMOTE_PATH)
    
    # Wait a moment
    time.sleep(2)
    
    # Test 3: Check sync status (should be in sync)
    print("\n🔍 Test 3: Sync status check")
    sync_file(bucket, TEST_FILE, REMOTE_PATH)
    
    # Test 4: Simulate remote file being newer
    print("\n🔍 Test 4: Simulating remote file being newer")
    # Create a different test file with newer timestamp
    newer_data = {
        "test_id": f"remote_test_{int(time.time())}",
        "created_at": datetime.now().isoformat(),
        "message": "This is a newer remote version",
        "version": "2.0"
    }
    
    # Upload this as the remote file
    temp_file = "temp_remote.json"
    create_test_file(temp_file, newer_data)
    upload_file_with_metadata(bucket, temp_file, REMOTE_PATH)
    
    # Now sync - should download the newer remote version
    sync_file(bucket, TEST_FILE, REMOTE_PATH)
    
    # Cleanup
    print("\n🧹 Cleaning up test files...")
    if os.path.exists(TEST_FILE):
        os.remove(TEST_FILE)
    if os.path.exists(temp_file):
        os.remove(temp_file)
    
    print("\n✅ Test completed!")

if __name__ == "__main__":
    main()
