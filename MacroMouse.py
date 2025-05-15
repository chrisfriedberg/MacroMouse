import customtkinter as ctk
import tkinter as tk
from tkinter import messagebox, filedialog
import pyperclip
import subprocess
import os
import sys
import json
from datetime import datetime
import uuid
import xml.etree.ElementTree as ET
import re
import tkinter.simpledialog as sd

# --- GLOBALS ---
log_file_path = None
config_file_path = None
selected_macro_name = None
macro_list_items = []
selected_category = "All"
macro_data_file_path = None
reference_file_path = None  # Path to user's reference file
macros_dict = {}  # In-memory macro storage
macro_usage_counts = {}  # Dictionary to track macro usage counts
window = None  # Global reference to the main window
update_list_func = None  # Global reference to the update_list function
# Dictionary to store the 'Leave Raw' preference for each macro
global macro_leave_raw_preferences
macro_leave_raw_preferences = {}

# --- LOGGING ---
def get_log_timestamp():
    """Return a timestamp string for logging."""
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

def log_message(message):
    """Append a log message to the log file."""
    global log_file_path
    if log_file_path:
        try:
            log_dir = os.path.dirname(log_file_path)
            if log_dir and not os.path.exists(log_dir):
                os.makedirs(log_dir, exist_ok=True)
            with open(log_file_path, "a", encoding="utf-8") as logf:
                logf.write(f"[{get_log_timestamp()}] {message}\n")
        except Exception as e:
            print(f"Warning: could not write log file '{log_file_path}': {e}")

# --- CONFIG MANAGEMENT ---
def load_config():
    """Load configuration from config.json"""
    global config_file_path
    if not config_file_path or not os.path.exists(config_file_path):
        return {}
    try:
        with open(config_file_path, 'r') as f:
            return json.load(f)
    except Exception as e:
        log_message(f"Error loading config: {e}")
        return {}

def save_config(config_data):
    """Save configuration to config.json"""
    global config_file_path
    if not config_file_path:
        return False
    try:
        with open(config_file_path, 'w') as f:
            json.dump(config_data, f, indent=4)
        return True
    except Exception as e:
        log_message(f"Error saving config: {e}")
        return False

def change_app_icon(window):
    """Handle changing the application icon"""
    config = load_config()
    
    icon_path = filedialog.askopenfilename(
        title="Select Application Icon",
        filetypes=[("Icon files", "*.ico")],
        initialdir=os.path.dirname(config.get('icon_path', '')) or os.path.expanduser("~")
    )
    
    if not icon_path:
        return
        
    try:
        window.iconbitmap(icon_path)
        config['icon_path'] = icon_path
        if save_config(config):
            log_message(f"Changed application icon to: {icon_path}")
        else:
            messagebox.showerror("Error", "Failed to save icon configuration")
    except Exception as e:
        messagebox.showerror("Error", f"Invalid icon file: {e}")
        if 'icon_path' in config:
            try:
                window.iconbitmap(config['icon_path'])
            except:
                pass

# --- DATA MANAGEMENT ---
def generate_unique_id(prefix="MACRO"):
    """Generate a unique ID for macros or categories."""
    return f"{prefix}_{uuid.uuid4().hex[:8].upper()}"

def load_macro_data():
    """Load the structured macro data from XML file."""
    global macro_data_file_path
    
    # Debug log
    log_message(f"Attempting to load macro data from: {macro_data_file_path}")
    
    if not macro_data_file_path or not os.path.exists(macro_data_file_path):
        log_message(f"Macro data file does not exist at: {macro_data_file_path}")
        return {
            "version": "1.0",
            "categories": {},
            "macros": {},
            "category_order": []
        }
        
    try:
        tree = ET.parse(macro_data_file_path)
        root = tree.getroot()
        
        data = {
            "version": root.find("version").text if root.find("version") is not None else "1.0",
            "categories": {},
            "macros": {},
            "category_order": []
        }
        
        # Load categories
        categories_elem = root.find("categories")
        if categories_elem is not None:
            for cat_elem in categories_elem.findall("category"):
                cat_id = cat_elem.get("id")
                data["categories"][cat_id] = {
                    "name": cat_elem.find("name").text,
                    "created": cat_elem.find("created").text,
                    "modified": cat_elem.find("modified").text,
                    "description": cat_elem.find("description").text if cat_elem.find("description") is not None else ""
                }
        
        # Load macros
        macros_elem = root.find("macros")
        if macros_elem is not None:
            for macro_elem in macros_elem.findall("macro"):
                macro_id = macro_elem.get("id")
                data["macros"][macro_id] = {
                    "name": macro_elem.find("name").text,
                    "category_id": macro_elem.find("category_id").text,
                    "content": macro_elem.find("content").text,
                    "created": macro_elem.find("created").text,
                    "modified": macro_elem.find("modified").text,
                    "version": int(macro_elem.find("version").text)
                }
        
        # Load category order
        order_elem = root.find("category_order")
        if order_elem is not None and order_elem.text:
            data["category_order"] = order_elem.text.split(",")
        else:
            data["category_order"] = list(data["categories"].keys())
        
        log_message(f"Successfully loaded macro data with {len(data['categories'])} categories and {len(data['macros'])} macros")
        return data
    except Exception as e:
        log_message(f"Error loading macro data: {e}")
        return {
            "version": "1.0",
            "categories": {},
            "macros": {},
            "category_order": []
        }

def save_macro_data(data, category_order=None):
    """Save the structured macro data to XML file."""
    global macro_data_file_path
    
    if not macro_data_file_path:
        log_message("Cannot save macro data: File path is not set")
        return False
        
    # Ensure directory exists
    data_dir = os.path.dirname(macro_data_file_path)
    if not os.path.exists(data_dir):
        try:
            os.makedirs(data_dir, exist_ok=True)
            log_message(f"Created data directory: {data_dir}")
        except Exception as e:
            log_message(f"Error creating data directory: {e}")
            return False
            
    try:
        root = ET.Element("macro_data")
        
        version_elem = ET.SubElement(root, "version")
        version_elem.text = data.get("version", "1.0")
        
        if category_order is None:
            category_order = list(data["categories"].keys())
        order_elem = ET.SubElement(root, "category_order")
        order_elem.text = ",".join(category_order)
        
        categories_elem = ET.SubElement(root, "categories")
        for cat_id, cat_data in data["categories"].items():
            cat_elem = ET.SubElement(categories_elem, "category", id=cat_id)
            name_elem = ET.SubElement(cat_elem, "name")
            name_elem.text = cat_data["name"]
            created_elem = ET.SubElement(cat_elem, "created")
            created_elem.text = cat_data["created"]
            modified_elem = ET.SubElement(cat_elem, "modified")
            modified_elem.text = cat_data["modified"]
            desc_elem = ET.SubElement(cat_elem, "description")
            desc_elem.text = cat_data.get("description", "")
        
        macros_elem = ET.SubElement(root, "macros")
        for macro_id, macro_data in data["macros"].items():
            macro_elem = ET.SubElement(macros_elem, "macro", id=macro_id)
            name_elem = ET.SubElement(macro_elem, "name")
            name_elem.text = macro_data["name"]
            cat_id_elem = ET.SubElement(macro_elem, "category_id")
            cat_id_elem.text = macro_data["category_id"]
            content_elem = ET.SubElement(macro_elem, "content")
            content_elem.text = macro_data["content"]
            created_elem = ET.SubElement(macro_elem, "created")
            created_elem.text = macro_data["created"]
            modified_elem = ET.SubElement(macro_elem, "modified")
            modified_elem.text = macro_data["modified"]
            version_elem = ET.SubElement(macro_elem, "version")
            version_elem.text = str(macro_data["version"])
        
        tree = ET.ElementTree(root)
        
        # Create a backup before saving
        if os.path.exists(macro_data_file_path):
            backup_path = f"{macro_data_file_path}.bak"
            import shutil
            shutil.copy2(macro_data_file_path, backup_path)
            log_message(f"Created backup of macro data file at: {backup_path}")
            
        tree.write(macro_data_file_path, encoding="utf-8", xml_declaration=True)
        log_message(f"Successfully saved macro data to: {macro_data_file_path}")
        return True
    except Exception as e:
        log_message(f"Error saving macro data: {e}")
        return False

def create_new_category(name, description=""):
    """Create a new category in the macro data."""
    data = load_macro_data()
    cat_id = generate_unique_id("CAT")
    data["categories"][cat_id] = {
        "name": name,
        "created": datetime.now().isoformat(),
        "modified": datetime.now().isoformat(),
        "description": description
    }
    if save_macro_data(data):
        return cat_id
    return None

def add_macro_to_data(category_id, name, content):
    """Add a new macro to the structured data."""
    data = load_macro_data()
    macro_id = generate_unique_id()
    data["macros"][macro_id] = {
        "name": name,
        "category_id": category_id,
        "content": content,
        "created": datetime.now().isoformat(),
        "modified": datetime.now().isoformat(),
        "version": 1
    }
    if save_macro_data(data):
        return macro_id
    return None

def update_macro_in_data(macro_id, category_id, name, content):
    """Update an existing macro in the structured data."""
    data = load_macro_data()
    if macro_id in data["macros"]:
        macro = data["macros"][macro_id]
        macro.update({
            "name": name,
            "category_id": category_id,
            "content": content,
            "modified": datetime.now().isoformat(),
            "version": macro.get("version", 1) + 1
        })
        if save_macro_data(data):
            return True
    return False

def delete_macro_from_data(macro_id):
    """Delete a macro from the structured data."""
    data = load_macro_data()
    if macro_id in data["macros"]:
        del data["macros"][macro_id]
        return save_macro_data(data)
    return False

def get_category_by_name(name):
    """Get category ID by name."""
    data = load_macro_data()
    for cat_id, cat_data in data["categories"].items():
        if cat_data["name"] == name:
            return cat_id
    return None

def get_macro_by_name(name, category_id=None):
    """Get macro ID by name and optionally category."""
    data = load_macro_data()
    for macro_id, macro_data in data["macros"].items():
        if macro_data["name"] == name:
            if category_id is None or macro_data["category_id"] == category_id:
                return macro_id
    return None

def get_macros_for_ui(selected_category="All", search_term=""):
    """Returns a list of (category, macro_name, content) for UI display."""
    data = load_macro_data()
    macros = []
    search_term = search_term.lower().strip()
    for macro_id, macro in data["macros"].items():
        cat_id = macro["category_id"]
        cat_name = data["categories"].get(cat_id, {}).get("name", "Uncategorized")
        name = macro["name"]
        content = macro["content"]
        if selected_category != "All" and cat_name != selected_category:
            continue
        if search_term and (search_term not in name.lower() and search_term not in content.lower()):
            continue
        macros.append((cat_name, name, content))
    
    # First sort all macros alphabetically
    macros.sort(key=lambda x: x[1].lower())
    
    # Then identify and move the top 5 most used macros to the beginning
    if macro_usage_counts:
        # Filter usage counts to only include macros from the current category or search
        filtered_usage_counts = {}
        for (cat, name), count in macro_usage_counts.items():
            # For 'All' category, include all macros, otherwise filter by the selected category
            if selected_category == "All" or cat == selected_category:
                # Check if this macro is in our current list (might be filtered by search)
                if any(m[1] == name and (selected_category == "All" or m[0] == selected_category) for m in macros):
                    filtered_usage_counts[(cat, name)] = count
        
        # Sort macros by their usage count
        macros.sort(key=lambda x: (-filtered_usage_counts.get((x[0], x[1]), 0), x[1].lower()))
        
        # Optional: You can still separate top 5 from the rest if you want that visual distinction
        # Extract top 5 most used macros
        top_macros = [m for m in macros[:5] if filtered_usage_counts.get((m[0], m[1]), 0) > 0]
        rest_macros = [m for m in macros if m not in top_macros]
        
        # Combine them back
        macros = top_macros + rest_macros
    
    return macros

def get_categories():
    """Return a sorted list of categories from the data."""
    data = load_macro_data()
    order = data.get("category_order", list(data["categories"].keys()))
    names = [data["categories"][cid]["name"] for cid in order if cid in data["categories"]]
    return ["All"] + names

# --- FILE OPERATIONS ---
def show_copied_popup(parent, macro_name):
    popup = ctk.CTkToplevel(parent)
    popup.title("Copied")
    popup.geometry("320x180")
    popup.resizable(False, False)
    popup.attributes("-topmost", True)
    popup.grab_set()
    
    # Apply header styling
    header_frame = ctk.CTkFrame(popup, fg_color="#181C22", height=44, corner_radius=0)
    header_frame.pack(fill="x", side="top")
    
    title_label = ctk.CTkLabel(
        header_frame,
        text="Copied to Clipboard",
        font=("Segoe UI", 15, "bold"),
        text_color="white",
        anchor="w"
    )
    title_label.pack(side="left", padx=(15, 0), pady=6)
    
    # Content
    content_frame = ctk.CTkFrame(popup)
    content_frame.pack(fill="both", expand=True, padx=10, pady=10)
    
    message = f"Macro '{macro_name}' copied to clipboard."
    message_label = ctk.CTkLabel(
        content_frame, 
        text=message,
        font=("Segoe UI", 12),
        wraplength=280
    )
    message_label.pack(pady=(20, 15))
    
    ok_btn = ctk.CTkButton(content_frame, text="OK", command=popup.destroy, width=100)
    ok_btn.pack(pady=(0, 15))
    
    # Center the popup over the parent
    popup.update_idletasks()
    x = parent.winfo_rootx() + (parent.winfo_width() // 2) - (popup.winfo_width() // 2)
    y = parent.winfo_rooty() + (parent.winfo_height() // 2) - (popup.winfo_height() // 2)
    popup.geometry(f"+{x}+{y}")

def copy_macro(macro_key):
    """Copies the content of the specified macro to the clipboard."""
    if macro_key and macro_key in macros_dict:
        try:
            content = macros_dict[macro_key]
            
            # Check for {{tag}} placeholders and prompt for input
            pattern = r'\{\{(.*?)\}\}' 
            placeholders = set(re.findall(pattern, content))
            
            # If there are placeholders, show the dialog
            if placeholders:
                # Create a custom dialog for all placeholders
                inputs = show_placeholder_dialog(macro_key[1], list(placeholders))
                
                # If dialog was canceled, return immediately without copying
                if inputs is None:
                    return
                    
                # Replace placeholders with user input only if provided
                for ph, value in inputs.items():
                    if value:  # Only replace if a value was provided
                        content = content.replace(f"{{{{{ph}}}}}", value)
                
                # Do not show warning for unresolved tags if any are due to 'Leave Raw'
                # We skip the unresolved tags dialog entirely since 'Leave Raw' handles it
            
            # At this point, if we've made it here, user has confirmed all dialogs
            # Ensure we're copying plain text
            pyperclip.copy(content)
            log_message(f"Copied macro '{macro_key[1]}' to clipboard.")
            print(f"Copied to clipboard: {macro_key[1]}")
            
            # Update usage count
            macro_usage_counts[macro_key] = macro_usage_counts.get(macro_key, 0) + 1
            save_usage_counts()
            
            # Show copied success popup
            show_copied_popup(window, macro_key[1])
            
        except Exception as e:
            print(f"Error: {e}")
            messagebox.showerror("Clipboard Error", f"Could not copy to clipboard: {e}")
            log_message(f"Clipboard error for macro '{macro_key[1]}': {e}")
    elif macro_key:
        log_message(f"Attempted to copy non-existent macro: {macro_key}")
        messagebox.showerror("Error", f"Macro '{macro_key[1]}' not found in current data.")

def show_placeholder_dialog(macro_name, placeholders):
    """
    Creates a custom dialog to input values for all placeholders at once.
    Returns a dictionary of placeholder:value pairs or None if canceled.
    Includes a 'Leave Raw' checkbox for each placeholder to keep original text.
    """
    global macro_leave_raw_preferences
    dialog = ctk.CTkToplevel()
    dialog.title(f"Fill Placeholders for '{macro_name}'")
    # Keep window large but with more reasonable proportions
    dialog.geometry("900x650")  
    dialog.minsize(800, 600)
    dialog.grab_set()
    
    # Header with app-matching style
    header_frame = ctk.CTkFrame(dialog, fg_color="#181C22", height=60, corner_radius=0)
    header_frame.pack(fill="x", side="top")
    
    title_label = ctk.CTkLabel(
        header_frame,
        text=f"Fill Placeholder Values",
        font=("Segoe UI", 18, "bold"),
        text_color="white",
        anchor="w"
    )
    title_label.pack(side="left", padx=(20, 0), pady=10)
    
    # Content area with scrollable frame
    content_frame = ctk.CTkScrollableFrame(dialog)
    content_frame.pack(fill="both", expand=True, padx=25, pady=25)
    
    # Description label
    desc_label = ctk.CTkLabel(
        content_frame,
        text=f"Please enter values for the following placeholders in '{macro_name}':",
        font=("Segoe UI", 14),
        wraplength=700,
        justify="left"
    )
    desc_label.pack(anchor="w", pady=(0, 20))
    
    # Create entry fields and checkboxes for each placeholder
    entries = {}
    leave_raw_vars = {}
    sorted_placeholders = sorted(placeholders)
    
    # Initialize preferences for this macro if not already present
    if macro_name not in macro_leave_raw_preferences:
        macro_leave_raw_preferences[macro_name] = {}
    
    for i, placeholder in enumerate(sorted_placeholders):
        # Frame for each placeholder
        ph_frame = ctk.CTkFrame(content_frame, fg_color="transparent")
        ph_frame.pack(fill="x", pady=(5, 15))
        
        # Label for placeholder
        ph_label = ctk.CTkLabel(
            ph_frame,
            text=f"{placeholder}:",
            font=("Segoe UI", 13),
            anchor="w"
        )
        ph_label.pack(anchor="w")
        
        # Entry for value
        ph_entry = ctk.CTkEntry(ph_frame, width=700, height=35)
        ph_entry.pack(fill="x", pady=(5, 0))
        
        # 'Leave Raw' checkbox for this placeholder
        leave_raw_var = tk.BooleanVar(value=macro_leave_raw_preferences[macro_name].get(placeholder, False))
        leave_raw_checkbox = ctk.CTkCheckBox(
            ph_frame,
            text="Leave Raw",
            variable=leave_raw_var,
            font=("Segoe UI", 11)
        )
        leave_raw_checkbox.pack(anchor="w", pady=(2, 0))
        
        entries[placeholder] = ph_entry
        leave_raw_vars[placeholder] = leave_raw_var
    
    # Warning label (initially hidden)
    warning_frame = ctk.CTkFrame(dialog, fg_color="transparent")
    warning_frame.pack(fill="x", padx=25, pady=(0, 5))
    
    warning_label = ctk.CTkLabel(
        warning_frame,
        text="Warning: Some placeholders are empty. They will remain as {{placeholder}} in the text.",
        text_color="orange",
        font=("Segoe UI", 13),
        wraplength=700
    )
    warning_label.pack(fill="x")
    warning_frame.pack_forget()  # Initially hidden
    
    # Set focus to the first entry
    if sorted_placeholders:
        entries[sorted_placeholders[0]].focus_set()
    
    # Result variable to store output
    result = [None]
    
    # Button frame
    btn_frame = ctk.CTkFrame(dialog, fg_color="transparent")
    btn_frame.pack(fill="x", pady=(0, 20), padx=25)
    
    def on_cancel():
        result[0] = None
        dialog.destroy()
    
    def on_submit():
        # Save the 'Leave Raw' preferences for each placeholder in this macro
        for placeholder in sorted_placeholders:
            macro_leave_raw_preferences[macro_name][placeholder] = leave_raw_vars[placeholder].get()
        save_leave_raw_preferences()
        
        # Collect values from entries, respecting 'Leave Raw' settings
        values = {}
        for ph in sorted_placeholders:
            if not leave_raw_vars[ph].get():
                value = entries[ph].get()
                if value:  # Only include if a value was provided
                    values[ph] = value
        
        # Check if any non-'Leave Raw' entries are empty
        empty_entries = [ph for ph in sorted_placeholders if not leave_raw_vars[ph].get() and not entries[ph].get().strip()]
        
        if empty_entries:
            # Show warning only for non-'Leave Raw' empty entries
            warning_label.configure(text=f"Warning: {len(empty_entries)} placeholder(s) are empty. They will remain as {{placeholder}} in the text.")
            warning_frame.pack(fill="x", padx=25, pady=(0, 10))
            
            # Ask for confirmation
            confirm_btn = ctk.CTkButton(
                warning_frame,
                text="Continue Anyway",
                command=lambda: confirm_submit(values),
                fg_color="#FF8C00",  # Orange color
                hover_color="#E67300",  # Darker orange
                height=35,
                width=120
            )
            confirm_btn.pack(side="right", pady=(5, 0))
            
            # Highlight empty entries
            for ph in empty_entries:
                entries[ph].configure(border_color="orange", border_width=2)
        else:
            result[0] = values
            dialog.destroy()
    
    def confirm_submit(values):
        result[0] = values
        dialog.destroy()
    
    # Allow pressing Enter to submit
    def handle_enter(event):
        on_submit()
    
    for entry in entries.values():
        entry.bind("<Return>", handle_enter)
    
    # Allow Tab to navigate between entries (default behavior)
    
    # Buttons
    cancel_btn = ctk.CTkButton(
        btn_frame,
        text="Cancel",
        command=on_cancel,
        fg_color="red",
        height=35,
        width=100
    )
    cancel_btn.pack(side="left", padx=(0, 10))
    
    submit_btn = ctk.CTkButton(
        btn_frame,
        text="Submit",
        command=on_submit,
        height=35,
        width=100
    )
    submit_btn.pack(side="left")
    
    # Bind Escape key to cancel
    dialog.bind("<Escape>", lambda event: on_cancel())
    
    # Center on screen
    dialog.update_idletasks()
    width = dialog.winfo_width()
    height = dialog.winfo_height()
    x = (dialog.winfo_screenwidth() // 2) - (width // 2)
    y = (dialog.winfo_screenheight() // 2) - (height // 2)
    dialog.geometry(f"{width}x{height}+{x}+{y}")
    
    # Wait for user interaction
    dialog.wait_window()
    return result[0]

def open_macro_file():
    """Opens the current macro data file using the OS default text editor."""
    global macro_data_file_path
    
    if not macro_data_file_path:
        messagebox.showerror("Error", "Cannot open: Macro data file path is not set.")
        return
        
    log_message(f"Attempting to open macro data file: {macro_data_file_path}")
    
    # Check if file exists
    if not os.path.exists(macro_data_file_path):
        # Ask if user wants to create it
        create_new = messagebox.askyesno(
            "File Not Found", 
            f"The XML data file does not exist at:\n{macro_data_file_path}\n\nDo you want to create a new empty data file?"
        )
        
        if not create_new:
            return
            
        # Try to create empty data file
        try:
            data_dir = os.path.dirname(macro_data_file_path)
            if not os.path.exists(data_dir):
                os.makedirs(data_dir, exist_ok=True)
                
            # Create basic XML structure
            root = ET.Element("macro_data")
            version = ET.SubElement(root, "version")
            version.text = "1.0"
            categories = ET.SubElement(root, "categories")
            
            # Add Uncategorized category
            cat_id = generate_unique_id("CAT")
            cat = ET.SubElement(categories, "category", id=cat_id)
            name = ET.SubElement(cat, "name")
            name.text = "Uncategorized"
            created = ET.SubElement(cat, "created")
            created.text = datetime.now().isoformat()
            modified = ET.SubElement(cat, "modified")
            modified.text = datetime.now().isoformat()
            desc = ET.SubElement(cat, "description")
            desc.text = ""
            
            macros = ET.SubElement(root, "macros")
            cat_order = ET.SubElement(root, "category_order")
            cat_order.text = cat_id
            
            tree = ET.ElementTree(root)
            tree.write(macro_data_file_path, encoding="utf-8", xml_declaration=True)
            
            log_message(f"Created new empty macro data file at: {macro_data_file_path}")
            messagebox.showinfo("Success", f"Created new XML data file at:\n{macro_data_file_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Cannot create macro data file at {macro_data_file_path}:\n{e}")
            log_message(f"Error creating macro data file: {e}")
            return

    # Try to open the file
    try:
        if sys.platform.startswith('win32'):
            os.startfile(macro_data_file_path)
        elif sys.platform.startswith('darwin'):
            subprocess.run(['open', macro_data_file_path], check=True)
        else:
            subprocess.run(['xdg-open', macro_data_file_path], check=True)
        log_message("Successfully launched default editor for macro data file.")
    except Exception as e:
        messagebox.showerror("Error Opening File", f"An error occurred while trying to open the macro data file:\n{e}")
        log_message(f"Error opening macro data file: {e}")

def open_macro_file_location():
    """Opens the folder containing the macro data file."""
    global macro_data_file_path
    if not macro_data_file_path:
        messagebox.showerror("Error", "Cannot open: Macro data file path is not set.")
        return
    if not os.path.exists(macro_data_file_path):
        messagebox.showerror("Error", f"Cannot open: Macro data file not found at:\n{macro_data_file_path}")
        return

    try:
        folder_path = os.path.dirname(macro_data_file_path)
        if sys.platform.startswith('win32'):
            # Just open the folder - much simpler and more reliable
            os.startfile(folder_path)
        elif sys.platform.startswith('darwin'):
            subprocess.run(['open', folder_path], check=True)
        else:
            subprocess.run(['xdg-open', folder_path], check=True)
        log_message(f"Opened folder containing macro data file: {folder_path}")
    except Exception as e:
        messagebox.showerror("Error Opening Folder", f"An error occurred while trying to open the folder:\n{e}")
        log_message(f"Error opening folder: {e}")

def backup_data_files_folder():
    """Creates a zip backup of the entire data files folder."""
    global macro_data_file_path
    if not macro_data_file_path:
        messagebox.showerror("Error", "Cannot backup: Macro data file path is not set.")
        return

    try:
        # Get the data directory path
        data_dir = os.path.dirname(macro_data_file_path)
        if not os.path.exists(data_dir):
            messagebox.showerror("Error", f"Cannot backup: Data directory not found at:\n{data_dir}")
            return

        # Get the default backup filename with timestamp
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        default_filename = f"MacroMouse_Data_Backup_{timestamp}.zip"
        
        # Get the default backup directory (project directory)
        default_dir = os.path.dirname(os.path.abspath(__file__))
        
        # Ask user for backup location
        backup_path = filedialog.asksaveasfilename(
            title="Save Backup As",
            initialdir=default_dir,
            initialfile=default_filename,
            defaultextension=".zip",
            filetypes=[("ZIP files", "*.zip"), ("All files", "*.*")]
        )
        
        if not backup_path:  # User cancelled
            return
            
        # Create zip backup
        import zipfile
        with zipfile.ZipFile(backup_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for root, dirs, files in os.walk(data_dir):
                for file in files:
                    file_path = os.path.join(root, file)
                    arcname = os.path.relpath(file_path, data_dir)
                    zipf.write(file_path, arcname)

        log_message(f"Created backup of data files at: {backup_path}")
        
        # Show simple message with backup location
        backup_dir = os.path.dirname(backup_path)
        msg = f"Data Files backed up to:\n{backup_path}"
        messagebox.showinfo("Backup Complete", msg)
        
        # Open the backup directory
        try:
            if sys.platform.startswith('win32'):
                os.startfile(backup_dir)
            elif sys.platform.startswith('darwin'):
                subprocess.run(['open', backup_dir], check=True)
            else:
                subprocess.run(['xdg-open', backup_dir], check=True)
        except Exception as e:
            log_message(f"Error opening backup location: {e}")
            
    except Exception as e:
        messagebox.showerror("Backup Error", f"An error occurred while creating the backup:\n{e}")
        log_message(f"Error creating backup: {e}")

def restore_data_file():
    """Allows users to restore data files from backups."""
    global macro_data_file_path, config_file_path
    
    # Create the restore dialog
    restore_window = ctk.CTkToplevel()
    restore_window.title("Restore Data Files")
    restore_window.geometry("500x250")
    restore_window.grab_set()
    
    # Apply same theme as main window
    ctk.set_appearance_mode(ctk.get_appearance_mode())
    
    # Header
    header_frame = ctk.CTkFrame(restore_window, fg_color="#181C22", height=44, corner_radius=0)
    header_frame.pack(fill="x", side="top")
    
    title_label = ctk.CTkLabel(
        header_frame,
        text="Restore Data Files",
        font=("Segoe UI", 15, "bold"),
        text_color="white",
        anchor="w"
    )
    title_label.pack(side="left", padx=(15, 0), pady=6)
    
    close_btn = ctk.CTkButton(
        header_frame, text="✕", width=32, fg_color="#23272E", text_color="white",
        hover_color="#B22222", command=restore_window.destroy
    )
    close_btn.pack(side="right", padx=10, pady=6)
    
    # Content
    content_frame = ctk.CTkFrame(restore_window)
    content_frame.pack(fill="both", expand=True, padx=15, pady=15)
    
    # Prompt text
    prompt_label = ctk.CTkLabel(
        content_frame,
        text="Select the type of restore you want to perform:",
        font=("Segoe UI", 12),
        anchor="w"
    )
    prompt_label.pack(pady=(0, 15), anchor="w")
    
    # Restore options
    xml_var = tk.IntVar(value=1)
    xml_radio = ctk.CTkRadioButton(
        content_frame,
        text="Restore Single XML File (only macros.xml)",
        variable=xml_var,
        value=1
    )
    xml_radio.pack(pady=5, anchor="w")
    
    zip_radio = ctk.CTkRadioButton(
        content_frame,
        text="Restore from Zip Backup (all data files)",
        variable=xml_var,
        value=2
    )
    zip_radio.pack(pady=5, anchor="w")
    
    # Buttons
    btn_frame = ctk.CTkFrame(content_frame, fg_color="transparent")
    btn_frame.pack(fill="x", pady=(20, 0))
    
    def handle_restore():
        restore_type = xml_var.get()
        start_dir = os.path.dirname(os.path.abspath(__file__))
        
        if restore_type == 1:
            # Single XML file
            file_path = filedialog.askopenfilename(
                title="Select XML File to Restore",
                initialdir=start_dir,
                filetypes=[("XML files", "*.xml"), ("All files", "*.*")]
            )
            
            if not file_path:
                return
                
            target_path = macro_data_file_path
            if os.path.exists(target_path):
                # Ask for confirmation before overwriting
                if not styled_askyesno("Confirm Overwrite", 
                    f"This will overwrite your existing data file at:\n{target_path}\n\nAre you sure you want to continue?",
                    parent=restore_window):
                    return
            
            try:
                # Ensure directory exists
                os.makedirs(os.path.dirname(target_path), exist_ok=True)
                import shutil
                shutil.copy2(file_path, target_path)
                log_message(f"Restored XML file from {file_path} to {target_path}")
                styled_showinfo("Restore Complete", f"Successfully restored data file from:\n{file_path}", parent=restore_window)
                restore_window.destroy()
            except Exception as e:
                styled_showerror("Restore Error", f"An error occurred during restore:\n{e}", parent=restore_window)
                log_message(f"Restore error: {e}")
        
        else:
            # Zip backup
            zip_path = filedialog.askopenfilename(
                title="Select Zip Backup to Restore",
                initialdir=start_dir,
                filetypes=[("ZIP files", "*.zip"), ("All files", "*.*")]
            )
            
            if not zip_path:
                return
                
            target_dir = os.path.dirname(macro_data_file_path)
            
            # Check if any files will be overwritten
            if os.path.exists(target_dir) and any(os.path.exists(os.path.join(target_dir, f)) for f in ["macros.xml", "config.json"]):
                if not styled_askyesno("Confirm Overwrite", 
                    f"This will overwrite existing data files in:\n{target_dir}\n\nAre you sure you want to continue?",
                    parent=restore_window):
                    return
            
            try:
                # Ensure directory exists
                os.makedirs(target_dir, exist_ok=True)
                
                # Extract zip
                import zipfile
                with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                    zip_ref.extractall(target_dir)
                
                log_message(f"Restored data files from zip: {zip_path} to {target_dir}")
                styled_showinfo("Restore Complete", f"Successfully restored data files from:\n{zip_path}", parent=restore_window)
                restore_window.destroy()
            except Exception as e:
                styled_showerror("Restore Error", f"An error occurred during restore:\n{e}", parent=restore_window)
                log_message(f"Restore error: {e}")
    
    restore_btn = ctk.CTkButton(
        btn_frame,
        text="Select File and Restore",
        command=handle_restore
    )
    restore_btn.pack(side="left", padx=(0, 10))
    
    cancel_btn = ctk.CTkButton(
        btn_frame,
        text="Cancel",
        fg_color="gray",
        command=restore_window.destroy
    )
    cancel_btn.pack(side="left")

# --- POPUP WINDOWS ---
def add_macro_popup(update_list_func, category_dropdown):
    """Popup to add a new macro with category selection."""
    popup = ctk.CTkToplevel()
    popup.title("Add New Macro")
    popup.geometry("500x350")
    popup.grab_set()

    # Macro Name
    name_label = ctk.CTkLabel(popup, text="Macro Name:")
    name_label.pack(padx=10, pady=(10, 0), anchor="w")
    name_entry = ctk.CTkEntry(popup)
    name_entry.pack(padx=10, fill="x")

    # Category
    cat_label = ctk.CTkLabel(popup, text="Category:")
    cat_label.pack(padx=10, pady=(10, 0), anchor="w")
    
    data = load_macro_data()
    categories = [(cat_id, cat_data["name"]) for cat_id, cat_data in data["categories"].items()]
    categories.sort(key=lambda x: x[1])
    category_names = [name for _, name in categories]
    category_names.append("+ New Category")
    
    # Use the currently selected category if it's not "All"
    current_category = selected_category if selected_category != "All" else category_names[0]
    cat_var = tk.StringVar(value=current_category)
    cat_dropdown = ctk.CTkOptionMenu(popup, values=category_names, variable=cat_var)
    cat_dropdown.pack(padx=10, fill="x")

    def on_cat_change(choice):
        if choice == "+ New Category":
            new_cat = ctk.CTkInputDialog(text="Enter new category name:", title="New Category").get_input()
            if new_cat and new_cat.strip():
                new_cat = new_cat.strip()
                cat_id = create_new_category(new_cat)
                if cat_id:
                    category_names.insert(-1, new_cat)
                    cat_var.set(new_cat)
                    cat_dropdown.configure(values=category_names)
                    if category_dropdown:
                        category_dropdown.configure(values=get_categories())
                else:
                    messagebox.showerror("Error", "Failed to create new category.")
                    cat_var.set(category_names[0] if category_names else "Uncategorized")
            else:
                cat_var.set(category_names[0] if category_names else "Uncategorized")
    cat_dropdown.configure(command=on_cat_change)

    # Macro Content
    content_label = ctk.CTkLabel(popup, text="Macro Content:")
    content_label.pack(padx=10, pady=(10, 0), anchor="w")
    content_text = ctk.CTkTextbox(popup, height=6)
    content_text.pack(padx=10, pady=(0, 10), fill="both", expand=True)

    # Buttons
    btn_frame = ctk.CTkFrame(popup)
    btn_frame.pack(fill="x", padx=10, pady=(0, 10))

    def add_macro_action():
        name = name_entry.get().strip()
        category_name = cat_var.get().strip() or "Uncategorized"
        content = content_text.get("1.0", "end").strip()
        
        if not name:
            messagebox.showerror("Error", "Macro name cannot be empty.")
            return
            
        cat_id = get_category_by_name(category_name)
        if not cat_id:
            messagebox.showerror("Error", "Invalid category selected.")
            return
            
        if get_macro_by_name(name, cat_id):
            messagebox.showerror("Error", "A macro with this name already exists in this category.")
            return
            
        macro_id = add_macro_to_data(cat_id, name, content)
        if macro_id:
            macros_dict[(category_name, name)] = content
            update_list_func((category_name, name))
            popup.destroy()
            if category_dropdown:
                category_dropdown.configure(values=get_categories())
                category_dropdown.set("All")
        else:
            messagebox.showerror("Error", "Failed to save macro to data file.")

    add_btn = ctk.CTkButton(btn_frame, text="Add Macro", command=add_macro_action)
    add_btn.pack(side="left", padx=(0, 5), expand=True, fill="x")
    cancel_btn = ctk.CTkButton(btn_frame, text="Cancel", command=popup.destroy, fg_color="red")
    cancel_btn.pack(side="right", padx=(5, 0), expand=True, fill="x")

def edit_macro_popup(macro_name_to_edit, update_list_func, category_dropdown):
    """Popup to edit an existing macro with category selection."""
    data = load_macro_data()
    macro_id = None
    current_cat_id = None
    
    for mid, macro in data["macros"].items():
        if macro["name"] == macro_name_to_edit:
            macro_id = mid
            current_cat_id = macro["category_id"]
            break
    
    if not macro_id:
        messagebox.showerror("Error", "Macro not found in data structure.")
        return
    
    current_category = data["categories"][current_cat_id]["name"]
    
    popup = ctk.CTkToplevel()
    popup.title("Edit Macro")
    popup.geometry("500x350")
    popup.grab_set()

    # Macro Name
    name_label = ctk.CTkLabel(popup, text="Macro Name:")
    name_label.pack(padx=10, pady=(10, 0), anchor="w")
    name_entry = ctk.CTkEntry(popup)
    name_entry.insert(0, macro_name_to_edit)
    name_entry.pack(padx=10, fill="x")

    # Category
    cat_label = ctk.CTkLabel(popup, text="Category:")
    cat_label.pack(padx=10, pady=(10, 0), anchor="w")
    
    categories = [(cat_id, cat_data["name"]) for cat_id, cat_data in data["categories"].items()]
    categories.sort(key=lambda x: x[1])
    category_names = [name for _, name in categories]
    category_names.append("+ New Category")
    
    # Use the currently selected category if it's not "All" and matches the macro's category
    current_category = selected_category if selected_category != "All" and selected_category == current_category else current_category
    cat_var = tk.StringVar(value=current_category)
    cat_dropdown = ctk.CTkOptionMenu(popup, values=category_names, variable=cat_var)
    cat_dropdown.pack(padx=10, fill="x")

    def on_cat_change(choice):
        if choice == "+ New Category":
            new_cat = ctk.CTkInputDialog(text="Enter new category name:", title="New Category").get_input()
            if new_cat and new_cat.strip():
                new_cat = new_cat.strip()
                cat_id = create_new_category(new_cat)
                if cat_id:
                    category_names.insert(-1, new_cat)
                    cat_var.set(new_cat)
                    cat_dropdown.configure(values=category_names)
                    if category_dropdown:
                        category_dropdown.configure(values=get_categories())
                else:
                    messagebox.showerror("Error", "Failed to create new category.")
                    cat_var.set(current_category)
            else:
                cat_var.set(current_category)
    cat_dropdown.configure(command=on_cat_change)

    # Macro Content
    content_label = ctk.CTkLabel(popup, text="Macro Content:")
    content_label.pack(padx=10, pady=(10, 0), anchor="w")
    content_text = ctk.CTkTextbox(popup, height=6)
    content_text.insert("1.0", data["macros"][macro_id]["content"])
    content_text.pack(padx=10, pady=(0, 10), fill="both", expand=True)

    # Buttons
    btn_frame = ctk.CTkFrame(popup)
    btn_frame.pack(fill="x", padx=10, pady=(0, 10))
    
    def save_changes_action():
        new_name = name_entry.get().strip()
        new_category_name = cat_var.get().strip() or "Uncategorized"
        new_content = content_text.get("1.0", "end").strip()
        
        if not new_name:
            messagebox.showerror("Error", "Macro name cannot be empty.")
            return
            
        new_cat_id = get_category_by_name(new_category_name)
        if not new_cat_id:
            messagebox.showerror("Error", "Invalid category selected.")
            return
            
        if new_name != macro_name_to_edit and get_macro_by_name(new_name, new_cat_id):
            messagebox.showerror("Error", "A macro with this name already exists in this category.")
            return
            
        if update_macro_in_data(macro_id, new_cat_id, new_name, new_content):
            if (current_category, macro_name_to_edit) in macros_dict:
                del macros_dict[(current_category, macro_name_to_edit)]
            macros_dict[(new_category_name, new_name)] = new_content
            update_list_func((new_category_name, new_name))
            popup.destroy()
            if category_dropdown:
                category_dropdown.configure(values=get_categories())
                category_dropdown.set("All")
        else:
            messagebox.showerror("Error", "Failed to save macro to data file.")

    save_btn = ctk.CTkButton(btn_frame, text="Save Changes", command=save_changes_action)
    save_btn.pack(side="left", padx=(0, 5), expand=True, fill="x")
    cancel_btn = ctk.CTkButton(btn_frame, text="Cancel", command=popup.destroy, fg_color="red")
    cancel_btn.pack(side="right", padx=(5, 0), expand=True, fill="x")

def create_category_window(update_list_func, category_dropdown):
    """Create and show the category management window."""
    category_window = ctk.CTkToplevel()
    category_window.title("Manage Macro Categories")
    category_window.geometry("506x400")  # 10% wider than 460
    category_window.grab_set()

    # --- HEADER TOOLBAR ---
    header_frame = ctk.CTkFrame(category_window, fg_color="#181C22", height=44, corner_radius=0)
    header_frame.pack(fill="x", side="top")
    # App icon (if you want, else comment out)
    try:
        from PIL import Image, ImageTk
        icon_path = os.path.join(os.path.dirname(__file__), "your_icon.ico")  # Update path as needed
        if os.path.exists(icon_path):
            icon_img = Image.open(icon_path).resize((28, 28))
            icon_photo = ImageTk.PhotoImage(icon_img)
            icon_label = tk.Label(header_frame, image=icon_photo, bg="#181C22")
            icon_label.image = icon_photo
            icon_label.pack(side="left", padx=(12, 8), pady=6)
    except Exception:
        pass
    # Title
    title_label = ctk.CTkLabel(
        header_frame,
        text="Manage Macro Categories",
        font=("Segoe UI", 15, "bold"),
        text_color="white",
        anchor="w"
    )
    title_label.pack(side="left", padx=(8, 0), pady=6)
    # Optional: Add a close button on the right
    close_btn = ctk.CTkButton(
        header_frame, text="✕", width=32, fg_color="#23272E", text_color="white",
        hover_color="#B22222", command=category_window.destroy
    )
    close_btn.pack(side="right", padx=10, pady=6)

    data = load_macro_data()
    categories = data["categories"]

    # Initialize category_order from data or create new
    category_order = data.get("category_order", [])
    if not category_order:  # If no order exists, create one
        category_order = list(categories.keys())
        # Always put "Uncategorized" first if it exists
        unc_id = next((cat_id for cat_id, cat in categories.items() if cat["name"] == "Uncategorized"), None)
        if unc_id:
            category_order.remove(unc_id)
            category_order.insert(0, unc_id)
        # Save the initial order
        save_macro_data(data, category_order)

    list_frame = ctk.CTkScrollableFrame(category_window)
    list_frame.pack(fill="both", expand=True, padx=10, pady=10)

    button_frame = ctk.CTkFrame(category_window)
    button_frame.pack(fill="x", padx=10, pady=5)

    def update_category_list():
        for widget in list_frame.winfo_children():
            widget.destroy()
            
        category_order[:] = [cat_id for cat_id in category_order if cat_id in categories]
        
        for idx, cat_id in enumerate(category_order):
            cat_data = categories[cat_id]
            frame = ctk.CTkFrame(list_frame, fg_color="#23272E", corner_radius=8)
            frame.pack(fill="x", pady=(8 if idx == 1 else 4, 4), padx=2)  # Spacer after "Uncategorized"
            
            # Category label with style
            label = ctk.CTkLabel(
                frame, 
                text=cat_data["name"], 
                font=("Segoe UI", 13),  # <-- removed "bold"
                text_color="#00BFFF" if cat_data["name"] == "Uncategorized" else "white"
            )
            label.pack(side="left", padx=(10, 32))  # <-- 32px right padding for button wall
            
            if cat_data["name"] != "Uncategorized":
                btn_frame = ctk.CTkFrame(frame, fg_color="transparent")
                btn_frame.pack(side="right", padx=5)
                
                btn_style = {"font": ("Segoe UI", 11, "bold"), "text_color": "white"}
                
                # Edit button
                edit_btn = ctk.CTkButton(btn_frame, text="Edit", width=50, **btn_style, command=lambda c=cat_id: edit_category_popup(c, update_category_list, update_list_func, category_dropdown))
                edit_btn.pack(side="left", padx=(0, 6))
                # Delete button
                delete_btn = ctk.CTkButton(btn_frame, text="Delete", width=60, **btn_style, command=lambda c=cat_id: delete_category(c))
                delete_btn.pack(side="left", padx=(0, 6))
                # Top button
                top_btn = ctk.CTkButton(btn_frame, text="Top", width=40, **btn_style, command=lambda c=cat_id: move_category(c, "top"))
                top_btn.pack(side="left", padx=(0, 6))
                # Up button
                up_btn = ctk.CTkButton(btn_frame, text="↑", width=30, **btn_style, command=lambda c=cat_id: move_category(c, "up"))
                up_btn.pack(side="left", padx=(0, 6))
                # Down button
                down_btn = ctk.CTkButton(btn_frame, text="↓", width=30, **btn_style, command=lambda c=cat_id: move_category(c, "down"))
                down_btn.pack(side="left", padx=(0, 6))
                # Bottom button
                bottom_btn = ctk.CTkButton(btn_frame, text="Bottom", width=40, **btn_style, command=lambda c=cat_id: move_category(c, "bottom"))
                bottom_btn.pack(side="left", padx=(0, 0))

    def move_category(cat_id, direction):
        idx = category_order.index(cat_id)
        if direction == "top" and idx > 1:  # Don't move above "Uncategorized"
            # Remove from current position and insert after "Uncategorized"
            category_order.pop(idx)
            category_order.insert(1, cat_id)
        elif direction == "bottom":
            # Remove from current position and append to end
            category_order.pop(idx)
            category_order.append(cat_id)
        elif direction == "up" and idx > 1:  # Don't move above "Uncategorized"
            category_order[idx], category_order[idx-1] = category_order[idx-1], category_order[idx]
        elif direction == "down" and idx < len(category_order)-1:
            category_order[idx], category_order[idx+1] = category_order[idx+1], category_order[idx]
        
        update_category_list()
        save_macro_data(data, category_order)  # Save after move
        if category_dropdown:
            category_dropdown.configure(values=get_categories())
            category_dropdown.set("All")
        update_list_func()

    def add_category():
        name = ctk.CTkInputDialog(text="Enter category name:", title="Add Category").get_input()
        if name and name.strip():
            name = name.strip()
            if not any(cat["name"] == name for cat in categories.values()):
                cat_id = create_new_category(name)
                if cat_id:
                    # Reload data to get the new category
                    data = load_macro_data()
                    categories.clear()
                    categories.update(data["categories"])
                    
                    # Add to order list if not already present
                    if cat_id not in category_order:
                        category_order.append(cat_id)
                    
                    # Save the updated order
                    save_macro_data(data, category_order)
                    
                    # Update the UI
                    update_category_list()
                    if category_dropdown:
                        category_dropdown.configure(values=get_categories())
                        category_dropdown.set("All")
                    update_list_func()
                    log_message(f"Added new category: {name}")
            else:
                messagebox.showerror("Error", "A category with this name already exists.")

    def delete_category(cat_id):
        cat_name = categories[cat_id]["name"]
        associated_macros = [mid for mid, macro in data["macros"].items() if macro["category_id"] == cat_id]
        if associated_macros:
            dialog = ctk.CTkToplevel(category_window)
            dialog.title("Delete Category")
            dialog.geometry("370x170")
            dialog.grab_set()
            ctk.set_appearance_mode(ctk.get_appearance_mode())  # Match app theme
            msg = ctk.CTkLabel(dialog, text=f"Category '{cat_name}' has {len(associated_macros)} macro(s).\nWhat would you like to do?")
            msg.pack(pady=15)
            btn_frame = ctk.CTkFrame(dialog)
            btn_frame.pack(pady=10)
            def delete_macros():
                for mid in associated_macros:
                    del data["macros"][mid]
                del categories[cat_id]
                save_macro_data(data, category_order)
                dialog.destroy()
                update_category_list()
                category_dropdown.configure(values=get_categories())
                category_dropdown.set("All")
                update_list_func()
                log_message(f"Deleted category '{cat_name}' and its macros.")
            def move_to_uncategorized():
                unc_id = get_category_by_name("Uncategorized") or create_new_category("Uncategorized")
                for mid in associated_macros:
                    data["macros"][mid]["category_id"] = unc_id
                del categories[cat_id]
                save_macro_data(data, category_order)
                dialog.destroy()
                update_category_list()
                category_dropdown.configure(values=get_categories())
                category_dropdown.set("All")
                update_list_func()
                log_message(f"Moved macros from '{cat_name}' to Uncategorized and deleted category.")
            del_btn = ctk.CTkButton(btn_frame, text="Delete Macros", fg_color="red", command=delete_macros)
            del_btn.pack(side="left", padx=10)
            move_btn = ctk.CTkButton(btn_frame, text="Move to Uncategorized", command=move_to_uncategorized)
            move_btn.pack(side="left", padx=10)
            cancel_btn = ctk.CTkButton(dialog, text="Cancel", fg_color="gray", command=dialog.destroy)
            cancel_btn.pack(pady=5)
        else:
            if messagebox.askyesno("Confirm Delete", f"Delete category '{cat_name}'?"):
                del categories[cat_id]
                save_macro_data(data, category_order)
                update_category_list()
                category_dropdown.configure(values=get_categories())
                category_dropdown.set("All")
                update_list_func()
                log_message(f"Deleted category: {cat_id}")

    def sort_alphabetically():
        unc_id = [cat_id for cat_id in category_order if categories[cat_id]["name"] == "Uncategorized"][0]
        rest = sorted([cat_id for cat_id in category_order if categories[cat_id]["name"] != "Uncategorized"],
                      key=lambda cid: categories[cid]["name"].lower())
        category_order[:] = [unc_id] + rest
        update_category_list()
        save_macro_data(data, category_order)
        category_dropdown.configure(values=get_categories())
        category_dropdown.set("All")
        update_list_func()

    def on_close():
        if category_dropdown:
            category_dropdown.configure(values=get_categories())
            category_dropdown.set("All")
        update_list_func()
        category_window.destroy()

    sort_btn = ctk.CTkButton(button_frame, text="Sort Alphabetically", command=sort_alphabetically)
    sort_btn.pack(side="left", padx=5)
    add_btn = ctk.CTkButton(button_frame, text="Add Category", command=add_category)
    add_btn.pack(side="left", padx=5)
    close_btn = ctk.CTkButton(button_frame, text="Close", command=on_close)
    close_btn.pack(side="right", padx=5)

    # Initial update of the category list
    update_category_list()

def edit_category_popup(cat_id, update_category_list=None, update_list_func=None, category_dropdown=None):
    """Edit category popup that can be called from various contexts."""
    # Load fresh data
    data = load_macro_data()
    cat_data = data["categories"][cat_id]
    
    popup = ctk.CTkToplevel()
    popup.title("Edit Category")
    popup.geometry("350x180")
    popup.grab_set()
    
    name_label = ctk.CTkLabel(popup, text="Category Name:")
    name_label.pack(pady=(10, 0))
    name_entry = ctk.CTkEntry(popup)
    name_entry.insert(0, cat_data["name"])
    name_entry.pack(pady=5, padx=10, fill="x")
    
    desc_label = ctk.CTkLabel(popup, text="Description (optional):")
    desc_label.pack(pady=(10, 0))
    desc_entry = ctk.CTkEntry(popup)
    desc_entry.insert(0, cat_data.get("description", ""))
    desc_entry.pack(pady=5, padx=10, fill="x")
    
    def save_edit():
        new_name = name_entry.get().strip()
        new_desc = desc_entry.get().strip()
        
        if not new_name:
            messagebox.showerror("Error", "Category name cannot be empty.")
            return
            
        # Update category data
        cat_data["name"] = new_name
        cat_data["description"] = new_desc
        cat_data["modified"] = datetime.now().isoformat()
        
        # Save changes to file
        save_macro_data(data)
        
        # Update UI elements if provided
        if update_category_list:
            update_category_list()
        if category_dropdown:
            category_dropdown.configure(values=get_categories())
            category_dropdown.set("All")
        if update_list_func:
            update_list_func()
            
        popup.destroy()
        
    save_btn = ctk.CTkButton(popup, text="Save", command=save_edit)
    save_btn.pack(pady=10)

# --- MAIN APPLICATION WINDOW ---
def create_macro_window():
    """Main application window with menu bar, category dropdown, and macro list."""
    global selected_macro_name, macro_list_items, selected_category, macros_dict
    global window, update_list_func  # Add this line
    
    # Initialize macros_dict from XML data
    data = load_macro_data()
    macros_dict = {}
    for macro_id, macro in data["macros"].items():
        cat_id = macro["category_id"]
        cat_name = data["categories"].get(cat_id, {}).get("name", "Uncategorized")
        name = macro["name"]
        content = macro["content"]
        macros_dict[(cat_name, name)] = content

    window = ctk.CTk()  # This now sets the global window
    window.title("MacroMouse")
    window.geometry("1000x700")
    window.minsize(900, 600)

    config = load_config()
    if 'icon_path' in config and os.path.exists(config['icon_path']):
        try:
            window.iconbitmap(config['icon_path'])
        except Exception as e:
            log_message(f"Error loading saved icon: {e}")

    menubar = tk.Menu(window)
    window.config(menu=menubar)
    
    file_menu = tk.Menu(menubar, tearoff=0)
    menubar.add_cascade(label="File", menu=file_menu)
    
    # Data File submenu
    data_file_menu = tk.Menu(file_menu, tearoff=0)
    file_menu.add_cascade(label="Data File", menu=data_file_menu)
    data_file_menu.add_command(label="Show Macro XML Datafile", command=open_macro_file)
    data_file_menu.add_command(label="Open Data File Location", command=open_macro_file_location)
    data_file_menu.add_command(label="Back Up Data Files Folder", command=backup_data_files_folder)
    data_file_menu.add_command(label="Restore Data File", command=restore_data_file)
    data_file_menu.add_separator()
    data_file_menu.add_command(label="Show File Paths (Debug)", command=show_file_paths)
    
    file_menu.add_separator()
    file_menu.add_command(label="Change App Icon", command=lambda: change_app_icon(window))
    
    # Reference File submenu
    reference_file_menu = tk.Menu(file_menu, tearoff=0)
    file_menu.add_cascade(label="Reference File", menu=reference_file_menu)
    reference_file_menu.add_command(label="Select Reference File", command=select_reference_file)
    reference_file_menu.add_command(label="View Reference File", command=view_reference_file)
    
    tools_menu = tk.Menu(menubar, tearoff=0)
    menubar.add_cascade(label="Tools", menu=tools_menu)
    tools_menu.add_command(label="Macro Categories", command=lambda: create_category_window(update_list, category_dropdown))
    tools_menu.add_command(label="Delete All Macro Usage Counts", command=delete_all_usage_counts)
    # Theme submenu
    theme_menu = tk.Menu(tools_menu, tearoff=0)
    tools_menu.add_cascade(label="Theme", menu=theme_menu)
    theme_mode = tk.StringVar(value=load_config().get('theme_mode', 'Dark'))
    theme_menu.add_radiobutton(label="Dark", variable=theme_mode, value="Dark", command=lambda: set_theme("Dark"))
    theme_menu.add_radiobutton(label="Light", variable=theme_mode, value="Light", command=lambda: set_theme("Light"))
    
    help_menu = tk.Menu(menubar, tearoff=0)
    menubar.add_cascade(label="Help", menu=help_menu)
    
    # Open README.md instead of showing a generic About box
    def open_readme():
        readme_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "README.md")
        if os.path.exists(readme_path):
            try:
                if sys.platform.startswith('win32'):
                    os.startfile(readme_path)
                elif sys.platform.startswith('darwin'):
                    subprocess.run(['open', readme_path], check=True)
                else:
                    subprocess.run(['xdg-open', readme_path], check=True)
                log_message(f"Opened README.md file")
            except Exception as e:
                messagebox.showerror("Error", f"Could not open README.md file:\n{e}")
                log_message(f"Error opening README.md: {e}")
        else:
            messagebox.showerror("Error", "README.md file not found in application directory.")
            
    help_menu.add_command(label="About / Help", command=open_readme)
    help_menu.add_command(label="Placeholder References", command=show_placeholder_help)

    window.grid_rowconfigure(0, weight=1)
    window.grid_columnconfigure(0, weight=1)
    window.grid_columnconfigure(1, weight=2)

    left_frame = ctk.CTkFrame(window)
    left_frame.grid(row=0, column=0, sticky="nsew")
    # Set row weights for resizing
    left_frame.grid_rowconfigure(0, weight=0)  # category_frame
    left_frame.grid_rowconfigure(1, weight=0)  # search_frame
    left_frame.grid_rowconfigure(2, weight=1)  # macro_list_frame (main resizable area)
    left_frame.grid_rowconfigure(3, weight=0)  # action_button_frame
    left_frame.grid_rowconfigure(4, weight=0)  # close_button

    category_frame = ctk.CTkFrame(left_frame)
    category_frame.grid(row=0, column=0, padx=10, pady=(10, 20), sticky="ew")
    category_label = ctk.CTkLabel(category_frame, text="Category:")
    category_label.pack(side="left", padx=(0, 5))

    def on_category_change(choice):
        global selected_category
        selected_category = choice
        update_list()

    category_dropdown = ctk.CTkOptionMenu(
        category_frame,
        values=get_categories(),
        command=on_category_change
    )
    category_dropdown.pack(side="left", fill="x", expand=True)
    category_dropdown.set("All")

    search_var = tk.StringVar()
    search_frame = ctk.CTkFrame(left_frame)
    search_frame.grid(row=1, column=0, padx=10, pady=(0, 5), sticky="ew")
    search_label = ctk.CTkLabel(search_frame, text="Search Macros:")
    search_label.pack(side="left", padx=(0, 5))
    search_entry = ctk.CTkEntry(search_frame, textvariable=search_var)
    search_entry.pack(side="left", fill="x", expand=True)

    refresh_btn = ctk.CTkButton(search_frame, text="Refresh", width=80, command=lambda: update_list())
    refresh_btn.pack(side="right", padx=(5, 0))

    search_var.trace_add("write", lambda *args: update_list())

    macro_list_frame = ctk.CTkScrollableFrame(left_frame, label_text="Available Macros")
    macro_list_frame.grid(row=2, column=0, padx=10, pady=(0, 0), sticky="nsew")
    macro_list_frame.grid_columnconfigure(0, weight=1)

    action_button_frame = ctk.CTkFrame(left_frame, fg_color="transparent")
    action_button_frame.grid(row=3, column=0, padx=10, pady=(10, 5), sticky="ew")
    action_button_frame.grid_columnconfigure((0, 1, 2, 3), weight=1)

    def add_action():
        add_macro_popup(update_list, category_dropdown)

    def edit_action():
        if selected_macro_name:
            edit_macro_popup(selected_macro_name[1], update_list, category_dropdown)
        else:
            messagebox.showwarning("Select Macro", "Please select a macro from the list to edit.")

    def remove_action():
        global selected_macro_name
        if selected_macro_name:
            if messagebox.askyesno("Confirm Delete", f"Are you sure you want to remove '{selected_macro_name[1]}'?"):
                data = load_macro_data()
                macro_id = None
                for mid, macro in data["macros"].items():
                    if macro["name"] == selected_macro_name[1]:
                        macro_id = mid
                        break
                if macro_id and delete_macro_from_data(macro_id):
                    if selected_macro_name in macros_dict:
                        del macros_dict[selected_macro_name]
                    selected_macro_name = None
                    update_list()
                else:
                    messagebox.showerror("Error", "Failed to delete macro from data file.")

    def reset_order():
        """Reset the macro order by keeping only top 5 counts and sorting the rest alphabetically."""
        global macro_usage_counts
        
        # Get the top 5 most used macros across all categories
        top_macros = sorted(macro_usage_counts.items(), key=lambda x: x[1], reverse=True)[:5]
        
        # Clear all usage counts
        macro_usage_counts.clear()
        
        # Add back only the top 5
        for macro_key, count in top_macros:
            if count > 0:  # Only keep macros that have been used
                macro_usage_counts[macro_key] = count
        
        # Save the updated counts
        save_usage_counts()
        
        # Re-fetch and sort macros
        update_list()
        
        # Show a small popup to confirm reset
        popup = ctk.CTkToplevel(window)
        popup.title("Order Reset")
        popup.geometry("350x120")
        popup.attributes('-topmost', True)
        popup.grab_set()
        
        # Create message based on how many top macros were kept
        top_count = len(macro_usage_counts)
        if top_count > 0:
            message = f"Reset complete. Kept your top {top_count} most used macros\nat the top, all others are now sorted alphabetically."
        else:
            message = "Reset complete. All macros are now sorted alphabetically."
        
        label = ctk.CTkLabel(popup, text=message)
        label.pack(pady=20)
        
        # Auto-close after 2.5 seconds
        popup.after(2500, popup.destroy)

    add_btn = ctk.CTkButton(action_button_frame, text="Add Macro", command=add_action)
    add_btn.grid(row=0, column=0, padx=2, sticky="ew")
    edit_btn = ctk.CTkButton(action_button_frame, text="Edit Macro", command=edit_action)
    edit_btn.grid(row=0, column=1, padx=2, sticky="ew")
    del_btn = ctk.CTkButton(action_button_frame, text="Delete Macro", command=remove_action, fg_color="red")
    del_btn.grid(row=0, column=2, padx=2, sticky="ew")
    reset_btn = ctk.CTkButton(action_button_frame, text="Reset Order", command=reset_order)
    reset_btn.grid(row=0, column=3, padx=2, sticky="ew")
    
    # New bottom action row for Close button
    bottom_action_frame = ctk.CTkFrame(left_frame, fg_color="transparent")
    bottom_action_frame.grid(row=4, column=0, padx=10, pady=(5, 10), sticky="ew")
    # Use the same column configuration as action_button_frame
    bottom_action_frame.grid_columnconfigure((0, 1, 2, 3), weight=1)
    
    # Use the same configuration as add_btn, but only in the first position
    close_button = ctk.CTkButton(bottom_action_frame, text="Close MacroMouse", command=window.destroy, fg_color="red", width=140)
    close_button.grid(row=0, column=0, padx=2, sticky="ew")
    
    # Add invisible placeholder buttons to maintain spacing consistency
    placeholder1 = ctk.CTkButton(bottom_action_frame, text="", fg_color="transparent", hover=False, border_width=0)
    placeholder1.grid(row=0, column=1, padx=2, sticky="ew")
    placeholder2 = ctk.CTkButton(bottom_action_frame, text="", fg_color="transparent", hover=False, border_width=0)
    placeholder2.grid(row=0, column=2, padx=2, sticky="ew")
    placeholder3 = ctk.CTkButton(bottom_action_frame, text="", fg_color="transparent", hover=False, border_width=0)
    placeholder3.grid(row=0, column=3, padx=2, sticky="ew")

    preview_frame = ctk.CTkFrame(window)
    preview_frame.grid(row=0, column=1, padx=(0, 10), pady=10, sticky="nsew")
    preview_frame.grid_rowconfigure(1, weight=1)
    preview_frame.grid_columnconfigure(0, weight=1)
    
    preview_header_frame = ctk.CTkFrame(preview_frame, fg_color="transparent")
    preview_header_frame.grid(row=0, column=0, padx=10, pady=(0, 5), sticky="ew")
    preview_label = ctk.CTkLabel(preview_header_frame, text="Macro Preview:")
    preview_label.pack(side="left", padx=(0, 5))
    # Add reference button
    reference_btn = ctk.CTkButton(preview_header_frame, text="View Reference", width=120, command=view_reference_file)
    reference_btn.pack(side="right")
    
    preview_text = ctk.CTkTextbox(preview_frame, wrap="word", font=("Consolas", 11))
    preview_text.grid(row=1, column=0, padx=10, pady=(0, 10), sticky="nsew")
    preview_text.configure(state="disabled")

    def clear_macro_list_widgets():
        global macro_list_items
        for widget in macro_list_items:
            widget.destroy()
        macro_list_items = []

    def update_list(selected=None):
        clear_macro_list_widgets()
        search_term = search_var.get()
        macros = get_macros_for_ui(selected_category, search_term)
        if not macros:
            no_results_label = ctk.CTkLabel(macro_list_frame, text="No macros found. Click 'Add Macro' to start.", text_color="gray")
            no_results_label.pack(pady=5)
            macro_list_items.append(no_results_label)
            update_preview(None)
            highlight_selected_item(None)
            return
        for cat, name, content in macros:
            macro_button = ctk.CTkButton(
                macro_list_frame, text=name, anchor="w", fg_color="transparent", hover=False, border_width=0,
                text_color=ctk.ThemeManager.theme["CTkLabel"]["text_color"],
                command=lambda c=cat, n=name: on_macro_select(c, n)
            )
            macro_button.pack(fill="x", pady=(0, 1), padx=1)
            # Add double-click event for copy
            def on_double_click(event, c=cat, n=name):
                copy_macro((c, n))
            macro_button.bind("<Double-Button-1>", on_double_click)
            macro_list_items.append(macro_button)
        if selected:
            on_macro_select(*selected)

    def update_preview(selected_key):
        preview_text.configure(state="normal")
        preview_text.delete("1.0", "end")
        if selected_key and selected_key in macros_dict:
            preview_text.insert("1.0", macros_dict[selected_key])
        preview_text.configure(state="disabled")

    def highlight_selected_item(selected_key):
        selected_color = ctk.ThemeManager.theme["CTkButton"]["hover_color"]
        for item_widget in macro_list_items:
            if isinstance(item_widget, ctk.CTkButton):
                item_widget.configure(border_width=0)
        if selected_key and isinstance(selected_key, tuple):
            for item_widget in macro_list_items:
                if isinstance(item_widget, ctk.CTkButton) and item_widget.cget("text") == selected_key[1]:
                    item_widget.configure(border_width=1, border_color=selected_color)
                    break

    def on_macro_select(category, name):
        global selected_macro_name
        selected_macro_name = (category, name)
        update_preview(selected_macro_name)
        highlight_selected_item(selected_macro_name)

    # --- Theme Menu ---
    def set_theme(mode):
        ctk.set_appearance_mode(mode)
        config = load_config()
        config['theme_mode'] = mode
        save_config(config)

    update_list()
    window.mainloop()
    
    log_message("Application window closed.")
    log_message("="*20 + " MacroMouse Session End " + "="*20 + "\n")

    # Store a reference to the update_list function
    update_list_func = update_list

def main():
    """Main entry point for the application."""
    global macro_data_file_path, log_file_path, config_file_path, reference_file_path
    
    # Store data in a subdirectory of the script's location
    script_dir = os.path.dirname(os.path.abspath(__file__))
    app_data_dir = os.path.join(script_dir, "MacroMouse_Data")
    os.makedirs(app_data_dir, exist_ok=True)
    
    # Set file paths - these are the correct locations
    macro_data_file_path = os.path.join(app_data_dir, "macros.xml")
    log_file_path = os.path.join(app_data_dir, "MacroMouse.log")
    config_file_path = os.path.join(app_data_dir, "config.json")
    
    # Log startup information
    log_message("="*20 + " MacroMouse Session Start " + "="*20)
    log_message(f"App data directory: {app_data_dir}")
    log_message(f"Macro data file: {macro_data_file_path}")
    log_message(f"Log file: {log_file_path}")
    log_message(f"Config file: {config_file_path}")
    
    # Initialize data if needed
    data = load_macro_data()
    if not any(cat["name"] == "Uncategorized" for cat in data["categories"].values()):
        create_new_category("Uncategorized")
    
    # Load usage counts
    load_usage_counts()
    
    # Load 'Leave Raw' preferences
    load_leave_raw_preferences()
    
    # Load reference file path from config
    config = load_config()
    reference_file_path = config.get('reference_file', None)
    if reference_file_path:
        log_message(f"Reference file: {reference_file_path}")
    
    # Set theme from config
    theme_mode = config.get('theme_mode', 'Dark')
    ctk.set_appearance_mode(theme_mode)
    ctk.set_default_color_theme("dark-blue")
    
    # Start the main window
    create_macro_window()

# Add a debugging function to check file paths
def show_file_paths():
    """Display current file paths in a message box for debugging."""
    global macro_data_file_path, log_file_path, config_file_path
    
    # Create styled dialog
    paths_dialog = ctk.CTkToplevel()
    paths_dialog.title("File Paths")
    paths_dialog.geometry("500x600")  # Doubled the height from 300 to 600
    paths_dialog.grab_set()
    paths_dialog.resizable(True, True)  # Allow both width and height resizing
    
    # Header with consistent styling
    header_frame = ctk.CTkFrame(paths_dialog, fg_color="#181C22", height=44, corner_radius=0)
    header_frame.pack(fill="x", side="top")
    
    title_label = ctk.CTkLabel(
        header_frame,
        text="File Paths",
        font=("Segoe UI", 15, "bold"),
        text_color="white",
        anchor="w"
    )
    title_label.pack(side="left", padx=(15, 0), pady=6)
    
    # Content frame
    content_frame = ctk.CTkFrame(paths_dialog)
    content_frame.pack(fill="both", expand=True, padx=15, pady=15)
    
    # Paths information with better formatting
    paths_info = [
        ("Macro Data File:", macro_data_file_path, os.path.exists(macro_data_file_path)),
        ("Log File:", log_file_path, os.path.exists(log_file_path)),
        ("Config File:", config_file_path, os.path.exists(config_file_path)),
        ("Current Directory:", os.getcwd(), True)
    ]
    
    for i, (label_text, path, exists) in enumerate(paths_info):
        # Label container
        entry_frame = ctk.CTkFrame(content_frame, fg_color="transparent")
        entry_frame.pack(fill="x", pady=(10 if i > 0 else 0))
        
        # Path label
        label = ctk.CTkLabel(
            entry_frame,
            text=label_text,
            font=("Segoe UI", 12, "bold"),
            anchor="w"
        )
        label.pack(anchor="w")
        
        # Path value with status indicator
        status_color = "#4CAF50" if exists else "#F44336"  # Green if exists, red if not
        path_text = ctk.CTkTextbox(entry_frame, height=40, wrap="word", activate_scrollbars=True)  # Increased height and enabled scrollbars
        path_text.insert("1.0", path)
        path_text.configure(state="disabled")
        path_text.pack(fill="x", pady=(2, 0))
        
        # Exists indicator
        status_text = f"Exists: {'Yes' if exists else 'No'}"
        status_label = ctk.CTkLabel(
            entry_frame,
            text=status_text,
            font=("Segoe UI", 10),
            text_color=status_color,
            anchor="w"
        )
        status_label.pack(anchor="w")
    
    # OK button at bottom
    btn_frame = ctk.CTkFrame(paths_dialog, fg_color="transparent")
    btn_frame.pack(fill="x", pady=(0, 15))
    
    ok_btn = ctk.CTkButton(
        btn_frame,
        text="OK",
        width=100,
        command=paths_dialog.destroy
    )
    ok_btn.pack(side="right", padx=15)
    
    # Center the dialog on screen
    paths_dialog.update_idletasks()
    width = paths_dialog.winfo_width()
    height = paths_dialog.winfo_height()
    x = (paths_dialog.winfo_screenwidth() // 2) - (width // 2)
    y = (paths_dialog.winfo_screenheight() // 2) - (height // 2)
    paths_dialog.geometry(f"{width}x{height}+{x}+{y}")
    
    # Add a note about resizing
    note_label = ctk.CTkLabel(
        btn_frame,
        text="Note: This window can be resized by dragging its edges",
        font=("Segoe UI", 10),
        text_color="gray",
        anchor="w"
    )
    note_label.pack(side="left", padx=15)
    
    paths_dialog.wait_window()

# Apply consistent styling to all message boxes
def create_styled_messagebox(title, message, parent=None, icon=None, buttons=None):
    """Create a styled message box that matches the app theme."""
    # Get current theme
    theme_mode = ctk.get_appearance_mode()
    
    # Create custom dialog
    dialog = ctk.CTkToplevel(parent)
    dialog.title(title)
    dialog.geometry("400x200")
    dialog.grab_set()
    dialog.resizable(False, False)
    
    # Header
    header_frame = ctk.CTkFrame(dialog, fg_color="#181C22", height=44, corner_radius=0)
    header_frame.pack(fill="x", side="top")
    
    title_label = ctk.CTkLabel(
        header_frame,
        text=title,
        font=("Segoe UI", 15, "bold"),
        text_color="white",
        anchor="w"
    )
    title_label.pack(side="left", padx=(15, 0), pady=6)
    
    # Message
    content_frame = ctk.CTkFrame(dialog)
    content_frame.pack(fill="both", expand=True, padx=10, pady=10)
    
    message_label = ctk.CTkLabel(
        content_frame,
        text=message,
        font=("Segoe UI", 12),
        wraplength=350,
        justify="left"
    )
    message_label.pack(pady=15, padx=10)
    
    # Buttons
    btn_frame = ctk.CTkFrame(dialog, fg_color="transparent")
    btn_frame.pack(fill="x", pady=(0, 10), padx=10)
    
    result = [None]  # Use list to store result (mutable)
    
    def on_button(value):
        result[0] = value
        dialog.destroy()
    
    # Default OK button
    if not buttons:
        ok_btn = ctk.CTkButton(
            btn_frame,
            text="OK",
            command=lambda: on_button(True)
        )
        ok_btn.pack(side="right", padx=5)
    else:
        # Custom buttons
        for btn_text, btn_value in buttons:
            btn = ctk.CTkButton(
                btn_frame,
                text=btn_text,
                command=lambda val=btn_value: on_button(val),
                fg_color="gray" if btn_text.lower() in ["cancel", "no"] else None
            )
            btn.pack(side="right", padx=5)
    
    # Center dialog
    dialog.update_idletasks()
    width = dialog.winfo_width()
    height = dialog.winfo_height()
    x = (dialog.winfo_screenwidth() // 2) - (width // 2)
    y = (dialog.winfo_screenheight() // 2) - (height // 2)
    dialog.geometry(f"{width}x{height}+{x}+{y}")
    
    # Wait for user interaction
    dialog.wait_window()
    return result[0]

# Create a styled version of messagebox functions
def styled_showinfo(title, message, parent=None):
    return create_styled_messagebox(title, message, parent)

def styled_askyesno(title, message, parent=None):
    return create_styled_messagebox(title, message, parent, 
                           buttons=[("Yes", True), ("No", False)])

def styled_showerror(title, message, parent=None):
    return create_styled_messagebox(title, message, parent)

def save_usage_counts():
    """Save macro usage counts to a separate JSON file."""
    global macro_data_file_path, macro_usage_counts
    if not macro_data_file_path:
        return False
    
    # Create the usage file path in the same directory as the macro data file
    usage_file_path = os.path.join(os.path.dirname(macro_data_file_path), "usage_counts.json")
    
    try:
        # Convert tuple keys to strings for JSON serialization
        serializable_counts = {}
        for key, count in macro_usage_counts.items():
            # Use a separator unlikely to appear in category or macro names
            serializable_key = f"{key[0]}|||{key[1]}"
            serializable_counts[serializable_key] = count
            
        with open(usage_file_path, 'w') as f:
            json.dump(serializable_counts, f, indent=4)
        return True
    except Exception as e:
        log_message(f"Error saving usage counts: {e}")
        return False

def load_usage_counts():
    """Load macro usage counts from a separate JSON file."""
    global macro_data_file_path, macro_usage_counts
    if not macro_data_file_path:
        return
    
    usage_file_path = os.path.join(os.path.dirname(macro_data_file_path), "usage_counts.json")
    
    if not os.path.exists(usage_file_path):
        return  # No usage data yet
    
    try:
        with open(usage_file_path, 'r') as f:
            serializable_counts = json.load(f)
            
        # Convert string keys back to tuples
        macro_usage_counts.clear()
        for key_str, count in serializable_counts.items():
            parts = key_str.split("|||", 1)
            if len(parts) == 2:
                macro_usage_counts[(parts[0], parts[1])] = count
    except Exception as e:
        log_message(f"Error loading usage counts: {e}")

def delete_all_usage_counts():
    """Delete all macro usage counts after confirmation."""
    global macro_usage_counts, window, update_list_func
    
    # Create styled confirmation dialog
    confirm_dialog = ctk.CTkToplevel(window)
    confirm_dialog.title("Confirm Delete")
    confirm_dialog.geometry("400x180")
    confirm_dialog.grab_set()
    
    # Header with dark background
    header_frame = ctk.CTkFrame(confirm_dialog, fg_color="#181C22", height=44, corner_radius=0)
    header_frame.pack(fill="x", side="top")
    
    title_label = ctk.CTkLabel(
        header_frame,
        text="Delete All Usage Counts",
        font=("Segoe UI", 15, "bold"),
        text_color="white",
        anchor="w"
    )
    title_label.pack(side="left", padx=(15, 0), pady=6)
    
    # Content
    content_frame = ctk.CTkFrame(confirm_dialog)
    content_frame.pack(fill="both", expand=True, padx=15, pady=15)
    
    message_label = ctk.CTkLabel(
        content_frame,
        text="Are you sure you want to delete ALL macro usage counts?\n\nThis action cannot be undone.",
        font=("Segoe UI", 12),
        wraplength=350
    )
    message_label.pack(pady=10)
    
    # Buttons with app-matching colors
    btn_frame = ctk.CTkFrame(content_frame, fg_color="transparent")
    btn_frame.pack(fill="x", pady=(10, 0))
    
    def confirm_delete():
        global macro_usage_counts
        # Clear all counts
        macro_usage_counts.clear()
        # Save the empty counts
        save_usage_counts()
        # Refresh the display
        if update_list_func:
            update_list_func()
        # Close dialog
        confirm_dialog.destroy()
        
        # Show success message
        success_popup = ctk.CTkToplevel(window)
        success_popup.title("Success")
        success_popup.geometry("300x100")
        success_popup.attributes('-topmost', True)
        success_popup.grab_set()
        
        success_label = ctk.CTkLabel(success_popup, text="All macro usage counts have been deleted.")
        success_label.pack(pady=20)
        
        # Auto-close after 2 seconds
        success_popup.after(2000, success_popup.destroy)
    
    def cancel():
        confirm_dialog.destroy()
    
    # No button (red)
    no_btn = ctk.CTkButton(
        btn_frame,
        text="No",
        command=cancel,
        fg_color="#B22222",  # Red color
        hover_color="#8B0000"  # Darker red on hover
    )
    no_btn.pack(side="left", padx=(0, 10), expand=True, fill="x")
    
    # Yes button (blue - default CTk color)
    yes_btn = ctk.CTkButton(
        btn_frame,
        text="Yes, Delete All",
        command=confirm_delete
    )
    yes_btn.pack(side="left", expand=True, fill="x")
    
    # Center the dialog
    confirm_dialog.update_idletasks()
    x = window.winfo_rootx() + (window.winfo_width() // 2) - (confirm_dialog.winfo_width() // 2)
    y = window.winfo_rooty() + (window.winfo_height() // 2) - (confirm_dialog.winfo_height() // 2)
    confirm_dialog.geometry(f"+{x}+{y}")

def select_reference_file():
    """Allows the user to select a reference file."""
    global reference_file_path
    config = load_config()
    
    initial_dir = os.path.dirname(config.get('reference_file', '')) or os.path.expanduser("~")
    file_path = filedialog.askopenfilename(
        title="Select Reference File",
        filetypes=[("Text files", "*.txt"), ("All files", "*.*")],
        initialdir=initial_dir
    )
    
    if not file_path:
        return False
        
    # Save the reference file path to config
    config['reference_file'] = file_path
    reference_file_path = file_path
    if save_config(config):
        log_message(f"Set reference file to: {file_path}")
        return True
    else:
        messagebox.showerror("Error", "Failed to save reference file configuration")
        return False

def view_reference_file():
    """Opens the reference file using the default system text editor."""
    global reference_file_path
    config = load_config()
    
    # Use the reference file path from config if it's not set globally
    if not reference_file_path:
        reference_file_path = config.get('reference_file', None)
        
    if not reference_file_path or not os.path.exists(reference_file_path):
        # Create a default reference file instead of prompting
        app_data_dir = os.path.dirname(macro_data_file_path)
        default_ref_path = os.path.join(app_data_dir, "placeholder_reference.txt")
        
        # Create the default reference file with examples
        try:
            with open(default_ref_path, 'w') as f:
                f.write("""# MacroMouse Placeholder Reference

This file serves as a reference for placeholders used in your macros.

## Common Placeholders

{{customer_name}} - Full name of the customer
{{first_name}} - Customer's first name
{{last_name}} - Customer's last name
{{order_number}} - Order or ticket reference number
{{date}} - Current date (you can use any format)
{{time}} - Current time
{{company}} - Your company name
{{product}} - Product name or service
{{agent_name}} - Your name or support representative name
{{contact_email}} - Contact email address
{{contact_phone}} - Contact phone number

## Email Templates Example

"Hello {{customer_name}},

Thank you for contacting {{company}} about your recent {{product}} purchase (#{{order_number}}). 

Your request has been received and is being processed by {{agent_name}}. We'll get back to you within 24 hours.

Best regards,
{{agent_name}}
{{company}}
{{contact_email}}
{{contact_phone}}"

## Notes

- You can add your own placeholders as needed
- The same placeholder used multiple times will be filled with the same value
- Placeholder names are case-sensitive
- Add your commonly used placeholder patterns below:

""")
            
            # Update config with the new reference file path
            reference_file_path = default_ref_path
            config['reference_file'] = default_ref_path
            save_config(config)
            log_message(f"Created default reference file at: {default_ref_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Could not create default reference file:\n{e}")
            log_message(f"Error creating default reference file: {e}")
            return
        
    # Try to open the file
    try:
        if sys.platform.startswith('win32'):
            os.startfile(reference_file_path)
        elif sys.platform.startswith('darwin'):
            subprocess.run(['open', reference_file_path], check=True)
        else:
            subprocess.run(['xdg-open', reference_file_path], check=True)
        log_message(f"Opened reference file: {reference_file_path}")
    except Exception as e:
        messagebox.showerror("Error Opening File", f"Could not open reference file:\n{e}")
        log_message(f"Error opening reference file: {e}")

def show_unresolved_tags_dialog(macro_name, unresolved_tags):
    """
    Shows a dialog warning about unresolved tags.
    Returns True to continue, False to cancel.
    """
    dialog = ctk.CTkToplevel()
    dialog.title("Unresolved Placeholders")
    # Keep window large but with more reasonable proportions
    dialog.geometry("800x550")
    dialog.minsize(700, 500)
    dialog.grab_set()
    
    # Header matching app style
    header_frame = ctk.CTkFrame(dialog, fg_color="#181C22", height=60, corner_radius=0)
    header_frame.pack(fill="x", side="top")
    
    title_label = ctk.CTkLabel(
        header_frame,
        text="Unresolved Placeholders",
        font=("Segoe UI", 18, "bold"),
        text_color="white",
        anchor="w"
    )
    title_label.pack(side="left", padx=(20, 0), pady=10)
    
    # Main content
    content_frame = ctk.CTkFrame(dialog)
    content_frame.pack(fill="both", expand=True, padx=25, pady=25)
    
    warning_label = ctk.CTkLabel(
        content_frame,
        text=f"The macro '{macro_name}' contains {len(unresolved_tags)} unresolved placeholder(s):",
        font=("Segoe UI", 14),
        wraplength=700,
        justify="left"
    )
    warning_label.pack(anchor="w", pady=(0, 20))
    
    # List of unresolved tags in a scrollable frame
    tags_frame = ctk.CTkScrollableFrame(content_frame, height=200)
    tags_frame.pack(fill="x", pady=(0, 20))
    
    for tag in sorted(unresolved_tags):
        tag_label = ctk.CTkLabel(
            tags_frame,
            text=f"• {{{{{tag}}}}}",
            font=("Segoe UI", 13),
            anchor="w"
        )
        tag_label.pack(anchor="w", pady=4)
    
    question_label = ctk.CTkLabel(
        content_frame,
        text="Do you want to continue with the unresolved placeholders?",
        font=("Segoe UI", 14),
        wraplength=700,
        justify="left"
    )
    question_label.pack(anchor="w", pady=(0, 20))
    
    # Result variable
    result = [False]
    
    # Buttons
    btn_frame = ctk.CTkFrame(content_frame, fg_color="transparent")
    btn_frame.pack(fill="x")
    
    def on_no():
        result[0] = False
        dialog.destroy()
    
    def on_yes():
        result[0] = True
        dialog.destroy()
    
    no_btn = ctk.CTkButton(
        btn_frame,
        text="No",
        command=on_no,
        fg_color="red",
        height=35,
        width=100
    )
    no_btn.pack(side="left", padx=(0, 10))
    
    yes_btn = ctk.CTkButton(
        btn_frame,
        text="Yes",
        command=on_yes,
        height=35,
        width=100
    )
    yes_btn.pack(side="left")
    yes_btn.focus_set()  # Set focus to Yes button
    
    # Keyboard shortcuts
    dialog.bind("<Escape>", lambda event: on_no())  # Esc key for No
    dialog.bind("<Return>", lambda event: on_yes())  # Enter key for Yes
    
    # Center on screen
    dialog.update_idletasks()
    width = dialog.winfo_width()
    height = dialog.winfo_height()
    x = (dialog.winfo_screenwidth() // 2) - (width // 2)
    y = (dialog.winfo_screenheight() // 2) - (height // 2)
    dialog.geometry(f"{width}x{height}+{x}+{y}")
    
    # Wait for user interaction
    dialog.wait_window()
    return result[0]

# Add help for the placeholder feature
def show_placeholder_help():
    help_popup = ctk.CTkToplevel(window)
    help_popup.title("Placeholder References Help")
    help_popup.geometry("900x650")
    help_popup.minsize(800, 600)
    help_popup.grab_set()
    
    # Header
    header_frame = ctk.CTkFrame(help_popup, fg_color="#181C22", height=44, corner_radius=0)
    header_frame.pack(fill="x", side="top")
    
    title_label = ctk.CTkLabel(
        header_frame,
        text="Placeholder References Help",
        font=("Segoe UI", 15, "bold"),
        text_color="white",
        anchor="w"
    )
    title_label.pack(side="left", padx=(15, 0), pady=6)
    
    # Content
    content_frame = ctk.CTkScrollableFrame(help_popup)
    content_frame.pack(fill="both", expand=True, padx=15, pady=15)
    
    help_text = """
    # Dynamic Placeholders
    
    MacroMouse supports dynamic placeholders in your macros using the {{tag}} syntax.
    
    ## How It Works
    
    1. When creating or editing a macro, include placeholders in double curly braces:
       Example: "Hello {{name}}, your order #{{order_number}} has been processed."
    
    2. When you copy the macro:
       - MacroMouse will detect all {{tag}} placeholders
       - For each unique tag, you'll be prompted to enter a value
       - All instances of that tag will be replaced with your input
    
    ## Tips
    
    - Use descriptive placeholder names like {{customer_name}}, {{date}}, etc.
    - You can use the Reference File to keep notes on your placeholder naming system
    - The same placeholder used multiple times will be replaced with the same value
    - Empty placeholders will remain as {{placeholder}} in the final text
    
    ## Reference File
    
    - The Reference File should be a plain text (.txt) file
    - You can create this file in any text editor (Notepad, etc.)
    - Use it to document your commonly used placeholders and their purpose
    - Access your reference file quickly via:
      • File → Reference File → View Reference File
      • The "View Reference" button next to the preview pane
    
    ## Example
    
    Original macro:
    "Hi {{customer}}, this is {{agent_name}} following up on your {{service}} inquiry from {{date}}."
    
    When copied, you'll be prompted for:
    - customer
    - agent_name
    - service
    - date
    
    And the resulting text will have all placeholders replaced with your entries.
    """
    
    help_label = ctk.CTkLabel(
        content_frame, 
        text=help_text,
        font=("Segoe UI", 12),
        wraplength=800,
        justify="left",
        anchor="w"
    )
    help_label.pack(pady=10, fill="both", expand=True)
    
    # OK button
    ok_btn = ctk.CTkButton(
        help_popup, 
        text="OK", 
        command=help_popup.destroy, 
        height=35,
        width=100,
        fg_color="red"
    )
    ok_btn.pack(pady=(0, 15))
    
    # Center on screen
    help_popup.update_idletasks()
    width = help_popup.winfo_width()
    height = help_popup.winfo_height()
    x = (help_popup.winfo_screenwidth() // 2) - (width // 2)
    y = (help_popup.winfo_screenheight() // 2) - (height // 2)
    help_popup.geometry(f"{width}x{height}+{x}+{y}")

def save_leave_raw_preferences():
    """Save 'Leave Raw' preferences for macros to a separate JSON file."""
    global macro_data_file_path, macro_leave_raw_preferences
    if not macro_data_file_path:
        return False
    
    # Create the preferences file path in the same directory as the macro data file
    preferences_file_path = os.path.join(os.path.dirname(macro_data_file_path), "leave_raw_preferences.json")
    
    try:
        with open(preferences_file_path, 'w') as f:
            json.dump(macro_leave_raw_preferences, f, indent=4)
        return True
    except Exception as e:
        log_message(f"Error saving 'Leave Raw' preferences: {e}")
        return False

def load_leave_raw_preferences():
    """Load 'Leave Raw' preferences for macros from a separate JSON file."""
    global macro_data_file_path, macro_leave_raw_preferences
    if not macro_data_file_path:
        return
    
    preferences_file_path = os.path.join(os.path.dirname(macro_data_file_path), "leave_raw_preferences.json")
    
    if not os.path.exists(preferences_file_path):
        return  # No preferences data yet
    
    try:
        with open(preferences_file_path, 'r') as f:
            macro_leave_raw_preferences.clear()
            macro_leave_raw_preferences.update(json.load(f))
    except Exception as e:
        log_message(f"Error loading 'Leave Raw' preferences: {e}")

if __name__ == "__main__":
    if sys.platform.startswith('win32'):
        try:
            import ctypes
            ctypes.windll.shcore.SetProcessDpiAwareness(2)
            print("Note: Set DPI awareness to Per-Monitor Aware V2.")
        except Exception as e:
            print(f"Note: Could not set DPI awareness: {e}")

    ctk.set_appearance_mode("dark")
    ctk.set_default_color_theme("dark-blue")
    main()