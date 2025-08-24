import os
import time
import tkinter as tk
from tkinter import messagebox
import customtkinter as ctk
import requests
from google.cloud import storage
from datetime import datetime
import json
import threading

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
    },
    'usage_counts.json': {
        'url': 'https://firebasestorage.googleapis.com/v0/b/spendingcache-personal.appspot.com/o/macro-data%2Fusage_counts.json?alt=media&token=a1234567-89ab-cdef-0123-456789abcdef',
        'firebase_path': 'macro-data/usage_counts.json'
    }
}

# === GOOGLE AUTH ===
os.environ["GOOGLE_APPLICATION_CREDENTIALS"] = r"service_account.json"
bucket_name = 'spendingcache-personal.appspot.com'

try:
    client = storage.Client()
    bucket = client.bucket(bucket_name)
except Exception as e:
    print(f"Failed to initialize Firebase: {e}")
    bucket = None

def get_local_timestamp(local_path):
    """Get local file modification timestamp."""
    if not os.path.exists(local_path):
        return 0
    return int(os.path.getmtime(local_path))

def get_remote_timestamp(firebase_path):
    """Get remote file timestamp from metadata or blob updated time."""
    if not bucket:
        raise Exception("Firebase not initialized")
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

def download_file_with_metadata(firebase_path, local_path):
    """Download file and preserve metadata."""
    if not bucket:
        raise Exception("Firebase not initialized")
    try:
        blob = bucket.blob(firebase_path)
        
        # Ensure directory exists
        os.makedirs(os.path.dirname(local_path), exist_ok=True)
        
        # Create backup before overwriting
        if os.path.exists(local_path):
            backup_path = f"{local_path}.backup"
            import shutil
            shutil.copy2(local_path, backup_path)
        
        # Download the file
        blob.download_to_filename(local_path)
        
        # Update local file timestamp to match remote if possible
        if blob.metadata and 'last_modified' in blob.metadata:
            remote_timestamp = float(blob.metadata['last_modified'])
            os.utime(local_path, (remote_timestamp, remote_timestamp))
        
        return True
    except Exception as e:
        print(f"Download error: {e}")
        return False

def upload_file_with_metadata(local_path, firebase_path):
    """Upload file with custom timestamp metadata."""
    if not bucket:
        raise Exception("Firebase not initialized")
    try:
        blob = bucket.blob(firebase_path)
        
        # Set custom metadata with current timestamp
        import time
        current_timestamp = str(time.time())
        metadata = {
            'last_modified': current_timestamp,
            'uploaded_at': datetime.now().isoformat(),
            'file_size': str(os.path.getsize(local_path))
        }
        
        blob.metadata = metadata
        blob.upload_from_filename(local_path)
        
        return True
    except Exception as e:
        print(f"Upload error: {e}")
        return False

def format_time(dt):
    """Format datetime for display."""
    if dt is None:
        return "Not Found"
    return dt.strftime("%Y-%m-%d %H:%M:%S UTC")

class SyncDialog(ctk.CTkToplevel):
    def __init__(self, parent=None):
        super().__init__(parent)
        
        self.title("MacroMouse Cloud Sync")
        self.geometry("800x600")
        self.resizable(True, True)
        
        # Configure grid
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)
        
        # Header
        header_frame = ctk.CTkFrame(self, fg_color="#181C22", height=60, corner_radius=0)
        header_frame.grid(row=0, column=0, sticky="ew")
        
        title_label = ctk.CTkLabel(
            header_frame,
            text="🔄 Cloud Sync Manager",
            font=("Segoe UI", 20, "bold"),
            text_color="white"
        )
        title_label.pack(pady=15)
        
        # Main content area
        self.content_frame = ctk.CTkScrollableFrame(self)
        self.content_frame.grid(row=1, column=0, sticky="nsew", padx=20, pady=20)
        
        # Status label
        self.status_label = ctk.CTkLabel(
            self.content_frame,
            text="Checking file status...",
            font=("Segoe UI", 14)
        )
        self.status_label.pack(pady=(0, 20))
        
        # File comparison cards
        self.file_frames = {}
        
        # Action buttons
        button_frame = ctk.CTkFrame(self)
        button_frame.grid(row=2, column=0, sticky="ew", padx=20, pady=(0, 20))
        
        self.auto_sync_btn = ctk.CTkButton(
            button_frame,
            text="Auto Sync All",
            font=("Segoe UI", 14),
            width=150,
            height=40,
            command=self.auto_sync_all
        )
        self.auto_sync_btn.pack(side="left", padx=10)
        
        self.manual_sync_btn = ctk.CTkButton(
            button_frame,
            text="Manual Sync Selected",
            font=("Segoe UI", 14),
            width=180,
            height=40,
            fg_color="#28a745",
            command=self.manual_sync
        )
        self.manual_sync_btn.pack(side="left", padx=10)
        
        self.refresh_btn = ctk.CTkButton(
            button_frame,
            text="Refresh",
            font=("Segoe UI", 14),
            width=100,
            height=40,
            fg_color="#6c757d",
            command=self.check_files
        )
        self.refresh_btn.pack(side="left", padx=10)
        
        self.close_btn = ctk.CTkButton(
            button_frame,
            text="Close",
            font=("Segoe UI", 14),
            width=100,
            height=40,
            fg_color="#dc3545",
            command=self.destroy
        )
        self.close_btn.pack(side="right", padx=10)
        
        # File sync choices
        self.sync_choices = {}
        
        # Start checking files
        self.check_files()
    
    def check_files(self):
        """Check all files and display their status."""
        self.status_label.configure(text="Checking file status...")
        
        # Clear existing frames
        for frame in self.file_frames.values():
            frame.destroy()
        self.file_frames.clear()
        self.sync_choices.clear()
        
        for filename, info in FILES.items():
            self.create_file_card(filename, info)
        
        self.status_label.configure(text="File status check complete. Choose sync action for each file.")
    
    def create_file_card(self, filename, info):
        """Create a card showing file comparison."""
        # Main card frame
        card_frame = ctk.CTkFrame(self.content_frame, corner_radius=10)
        card_frame.pack(fill="x", pady=10)
        
        # File name header
        header = ctk.CTkLabel(
            card_frame,
            text=filename,
            font=("Segoe UI", 16, "bold")
        )
        header.pack(pady=(10, 5))
        
        # Get file times
        local_path = os.path.join(LOCAL_DIR, filename)
        local_timestamp = get_local_timestamp(local_path)
        local_time = datetime.fromtimestamp(local_timestamp) if local_timestamp > 0 else None
        
        try:
            remote_timestamp = get_remote_timestamp(info['firebase_path'])
            remote_time = datetime.fromtimestamp(remote_timestamp) if remote_timestamp > 0 else None
        except Exception as e:
            remote_time = None
            error_label = ctk.CTkLabel(
                card_frame,
                text=f"Error checking remote: {e}",
                text_color="red"
            )
            error_label.pack()
        
        # Comparison frame
        comp_frame = ctk.CTkFrame(card_frame, fg_color="transparent")
        comp_frame.pack(fill="x", padx=20, pady=10)
        
        # Local info
        local_frame = ctk.CTkFrame(comp_frame)
        local_frame.pack(side="left", fill="x", expand=True, padx=10)
        
        ctk.CTkLabel(
            local_frame,
            text="📁 Local",
            font=("Segoe UI", 14, "bold")
        ).pack()
        
        ctk.CTkLabel(
            local_frame,
            text=format_time(local_time),
            font=("Segoe UI", 12)
        ).pack()
        
        # Status/Action
        status_frame = ctk.CTkFrame(comp_frame)
        status_frame.pack(side="left", padx=20)
        
        # Determine status
        if local_time is None and remote_time is None:
            status_text = "❌ Not Found"
            status_color = "red"
            recommendation = "Skip"
        elif local_time is None:
            status_text = "⬇️ Remote Only"
            status_color = "blue"
            recommendation = "Download"
        elif remote_time is None:
            status_text = "⬆️ Local Only"
            status_color = "green"
            recommendation = "Upload"
        elif local_time > remote_time:
            diff = (local_time - remote_time).total_seconds()
            status_text = f"⬆️ Local Newer ({int(diff)}s)"
            status_color = "green"
            recommendation = "Upload"
        elif remote_time > local_time:
            diff = (remote_time - local_time).total_seconds()
            status_text = f"⬇️ Remote Newer ({int(diff)}s)"
            status_color = "blue"
            recommendation = "Download"
        else:
            status_text = "✅ In Sync"
            status_color = "#28a745"
            recommendation = "Skip"
        
        status_label = ctk.CTkLabel(
            status_frame,
            text=status_text,
            font=("Segoe UI", 12, "bold"),
            text_color=status_color
        )
        status_label.pack(pady=5)
        
        # Sync choice radio buttons
        sync_var = tk.StringVar(value=recommendation)
        self.sync_choices[filename] = sync_var
        
        radio_frame = ctk.CTkFrame(status_frame, fg_color="transparent")
        radio_frame.pack()
        
        choices = [
            ("Upload to Cloud", "Upload"),
            ("Download from Cloud", "Download"),
            ("Skip", "Skip")
        ]
        
        for text, value in choices:
            radio = ctk.CTkRadioButton(
                radio_frame,
                text=text,
                variable=sync_var,
                value=value,
                font=("Segoe UI", 11)
            )
            radio.pack(anchor="w", pady=2)
        
        # Remote info
        remote_frame = ctk.CTkFrame(comp_frame)
        remote_frame.pack(side="right", fill="x", expand=True, padx=10)
        
        ctk.CTkLabel(
            remote_frame,
            text="☁️ Cloud",
            font=("Segoe UI", 14, "bold")
        ).pack()
        
        ctk.CTkLabel(
            remote_frame,
            text=format_time(remote_time),
            font=("Segoe UI", 12)
        ).pack()
        
        self.file_frames[filename] = card_frame
    
    def auto_sync_all(self):
        """Automatically sync all files based on timestamps."""
        synced = []
        errors = []
        
        self.status_label.configure(text="Auto-syncing files...")
        self.update()
        
        for filename, info in FILES.items():
            local_path = os.path.join(LOCAL_DIR, filename)
            firebase_path = info['firebase_path']
            
            try:
                local_timestamp = get_local_timestamp(local_path)
                remote_timestamp = get_remote_timestamp(firebase_path)
                
                # Format timestamps for logging
                local_time_str = datetime.fromtimestamp(local_timestamp).strftime('%Y-%m-%d %H:%M:%S') if local_timestamp > 0 else "N/A"
                remote_time_str = datetime.fromtimestamp(remote_timestamp).strftime('%Y-%m-%d %H:%M:%S') if remote_timestamp > 0 else "N/A"
                
                print(f"Auto sync: {filename} - Local: {local_time_str}, Remote: {remote_time_str}")
                
                # Compare timestamps and sync
                if local_timestamp > remote_timestamp:
                    # Local file is newer
                    if upload_file_with_metadata(local_path, firebase_path):
                        synced.append(f"⬆️ Uploaded newer local version of {filename}")
                    else:
                        errors.append(f"Failed to upload {filename}")
                        
                elif remote_timestamp > local_timestamp:
                    # Remote file is newer
                    if download_file_with_metadata(firebase_path, local_path):
                        synced.append(f"⬇️ Downloaded newer remote version of {filename}")
                    else:
                        errors.append(f"Failed to download {filename}")
                        
                elif local_timestamp == remote_timestamp and local_timestamp > 0:
                    # Files are in sync
                    synced.append(f"✅ {filename} is up to date")
                    
                else:
                    # One or both files don't exist
                    if local_timestamp == 0 and remote_timestamp == 0:
                        errors.append(f"{filename} doesn't exist locally or remotely")
                    elif local_timestamp == 0:
                        # Download remote file
                        if download_file_with_metadata(firebase_path, local_path):
                            synced.append(f"📥 Downloaded {filename} (no local copy)")
                        else:
                            errors.append(f"Failed to download {filename}")
                    else:
                        # Upload local file
                        if upload_file_with_metadata(local_path, firebase_path):
                            synced.append(f"📤 Uploaded {filename} (no remote copy)")
                        else:
                            errors.append(f"Failed to upload {filename}")
                        
            except Exception as e:
                errors.append(f"Error syncing {filename}: {e}")
        
        # Show results
        result_msg = "Auto Sync Complete!\n\n"
        if synced:
            result_msg += "Synced:\n" + "\n".join(synced)
        if errors:
            result_msg += "\n\nErrors:\n" + "\n".join(errors)
        
        self.show_result_dialog("Auto Sync Results", result_msg)
        self.check_files()  # Refresh display
    
    def manual_sync(self):
        """Manually sync files based on user selections."""
        synced = []
        errors = []
        
        self.status_label.configure(text="Syncing selected files...")
        self.update()
        
        for filename, sync_var in self.sync_choices.items():
            choice = sync_var.get()
            if choice == "Skip":
                continue
                
            local_path = os.path.join(LOCAL_DIR, filename)
            info = FILES[filename]
            firebase_path = info['firebase_path']
            
            try:
                if choice == "Upload":
                    if os.path.exists(local_path):
                        if upload_file_with_metadata(local_path, firebase_path):
                            synced.append(f"⬆️ Uploaded {filename}")
                        else:
                            errors.append(f"Failed to upload {filename}")
                    else:
                        errors.append(f"Cannot upload {filename}: File not found locally")
                        
                elif choice == "Download":
                    if download_file_with_metadata(firebase_path, local_path):
                        synced.append(f"⬇️ Downloaded {filename}")
                    else:
                        errors.append(f"Failed to download {filename}")
                        
            except Exception as e:
                errors.append(f"Error syncing {filename}: {e}")
        
        # Show results
        result_msg = "Manual Sync Complete!\n\n"
        if synced:
            result_msg += "Synced:\n" + "\n".join(synced)
        else:
            result_msg += "No files were synced."
        if errors:
            result_msg += "\n\nErrors:\n" + "\n".join(errors)
        
        self.show_result_dialog("Manual Sync Results", result_msg)
        self.check_files()  # Refresh display
    
    def show_result_dialog(self, title, message):
        """Show a result dialog."""
        dialog = ctk.CTkToplevel(self)
        dialog.title(title)
        dialog.geometry("500x400")
        dialog.grab_set()
        
        # Header
        header = ctk.CTkLabel(
            dialog,
            text=title,
            font=("Segoe UI", 16, "bold")
        )
        header.pack(pady=20)
        
        # Message
        msg_frame = ctk.CTkScrollableFrame(dialog)
        msg_frame.pack(fill="both", expand=True, padx=20, pady=(0, 20))
        
        msg_label = ctk.CTkLabel(
            msg_frame,
            text=message,
            font=("Segoe UI", 12),
            justify="left"
        )
        msg_label.pack(anchor="w")
        
        # OK button
        ok_btn = ctk.CTkButton(
            dialog,
            text="OK",
            command=dialog.destroy,
            width=100
        )
        ok_btn.pack(pady=(0, 20))
        
        # Center dialog
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (dialog.winfo_width() // 2)
        y = (dialog.winfo_screenheight() // 2) - (dialog.winfo_height() // 2)
        dialog.geometry(f"+{x}+{y}")

# === MAIN ===
if __name__ == "__main__":
    ctk.set_appearance_mode("dark")
    ctk.set_default_color_theme("dark-blue")
    
    root = ctk.CTk()
    root.withdraw()  # Hide main window
    
    sync_dialog = SyncDialog()
    sync_dialog.mainloop()
