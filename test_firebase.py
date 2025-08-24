#!/usr/bin/env python3
import json
from google.cloud import storage
from google.oauth2 import service_account

def test_firebase_connection():
    print("Testing Firebase Storage connection...")
    
    try:
        # Load service account
        creds = service_account.Credentials.from_service_account_file('service_account.json')
        print("✅ Service account loaded successfully")
        
        # Create client
        client = storage.Client(credentials=creds)
        print("✅ Storage client created successfully")
        
        # Test different bucket names
        bucket_names = [
            'spendingcache-personal',
            'spendingcache-personal.appspot.com',
            'spendingcache-personal.firebaseapp.com'
        ]
        
        for bucket_name in bucket_names:
            try:
                print(f"\nTesting bucket: {bucket_name}")
                bucket = client.bucket(bucket_name)
                
                # Try to list blobs
                blobs = list(bucket.list_blobs(max_results=1))
                print(f"✅ Bucket '{bucket_name}' exists and is accessible!")
                print(f"   Found {len(blobs)} blobs")
                
                if blobs:
                    print(f"   First blob: {blobs[0].name}")
                
                return bucket_name  # Found working bucket
                
            except Exception as e:
                print(f"❌ Bucket '{bucket_name}' failed: {str(e)}")
        
        print("\n❌ No buckets found. You may need to:")
        print("   1. Create a Firebase Storage bucket")
        print("   2. Give the service account Storage Object User permissions")
        print("   3. Check the bucket name in Firebase Console")
        
    except Exception as e:
        print(f"❌ Connection failed: {str(e)}")
        print("\nPossible issues:")
        print("   1. Service account key is invalid")
        print("   2. Service account doesn't have Storage permissions")
        print("   3. Firebase project doesn't exist")

if __name__ == "__main__":
    test_firebase_connection()
