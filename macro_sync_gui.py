import os
import time
import tkinter as tk
from tkinter import messagebox
import requests
from google.cloud import storage
from datetime import datetime
from email.utils import parsedate_to_datetime

# === CONFIG ===
LOCAL_DIR = r'C:\Users\chris\OneDrive\Desktop\scripts\MacroMouse\MacroMouse_Data'
FILES = {
    'macros.xml': {
        'url': 'https://firebasestorage.googleapis.com/v0/b/spendingcache-personal.appspot.com/o/macro-data%2Fmacros.xml?alt=media&token=9b66f288-0df6-420c-95a8-816d2dd81bad',
        'firebase_path': 'macro-data/macros.xml'
    },
    'config.json': {
        'url': 'https://firebasestorage.googleapis.com/v0/b/spendingcache-personal.appspot.com/o/macro-data%2Fconfig.json?alt=media&token=d8990581-5a15-4922-87b5-013c9a5a9ed8',
        'firebase_path': 'macro-data/config.json'
    },
    'MacroMouse.log': {
        'url': 'https://firebasestorage.googleapis.com/v0/b/spendingcache-personal.appspot.com/o/macro-data%2FMacroMouse.log?alt=media&token=f26e4225-4dbb-45d5-bffb-97ad2dbd74eb',
        'firebase_path': 'macro-data/MacroMouse.log'
    }
}

# === GOOGLE AUTH ===
os.environ["GOOGLE_APPLICATION_CREDENTIALS"] = r"path\to\your\service-account.json"
bucket_name = 'spendingcache-personal.appspot.com'

client = storage.Client()
bucket = client.bucket(bucket_name)

def get_remote_modified_time(file_url):
    metadata_url = file_url.replace("?alt=media", "")  # Strip ?alt=media
    r = requests.get(metadata_url)
    if r.status_code != 200:
        raise Exception(f"Failed to get metadata: {r.text}")
    updated_str = r.json().get("updated")  # ISO 8601 format
    return datetime.strptime(updated_str, "%Y-%m-%dT%H:%M:%S.%fZ")

def sync_files():
    synced = []

    for filename, info in FILES.items():
        local_path = os.path.join(LOCAL_DIR, filename)
        file_url = info['url']
        firebase_path = info['firebase_path']

        try:
            remote_time = get_remote_modified_time(file_url)
        except Exception as e:
            messagebox.showerror("Sync Error", f"Could not get remote time for {filename}: {e}")
            continue

        if os.path.exists(local_path):
            local_time = datetime.utcfromtimestamp(os.path.getmtime(local_path))

            if remote_time > local_time:
                # DOWNLOAD FROM FIREBASE
                r = requests.get(file_url)
                with open(local_path, 'wb') as f:
                    f.write(r.content)
                synced.append(f"⬇️ Downloaded newer version of {filename}")
            elif remote_time < local_time:
                # UPLOAD TO FIREBASE
                blob = bucket.blob(firebase_path)
                blob.upload_from_filename(local_path)
                synced.append(f"⬆️ Uploaded newer local version of {filename}")
            else:
                synced.append(f"✅ {filename} is up to date")
        else:
            # FILE DOESN’T EXIST LOCALLY → DOWNLOAD IT
            r = requests.get(file_url)
            with open(local_path, 'wb') as f:
                f.write(r.content)
            synced.append(f"📥 Downloaded {filename} (no local copy)")
    
    messagebox.showinfo("Sync Complete", "\n".join(synced))

# === UI ===
root = tk.Tk()
root.title("MacroMouse Cloud Sync")

btn = tk.Button(root, text="🔁 Sync Now", font=('Arial', 16), width=20, command=sync_files)
btn.pack(padx=40, pady=40)

root.mainloop()
