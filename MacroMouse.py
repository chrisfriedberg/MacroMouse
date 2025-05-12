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

# --- GLOBALS ---
log_file_path = None
config_file_path = None
selected_macro_name = None
macro_list_items = []
selected_category = "All"
macro_data_file_path = None
macros_dict = {}  # In-memory macro storage

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
    if not macro_data_file_path or not os.path.exists(macro_data_file_path):
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
        
        for cat_elem in root.find("categories").findall("category"):
            cat_id = cat_elem.get("id")
            data["categories"][cat_id] = {
                "name": cat_elem.find("name").text,
                "created": cat_elem.find("created").text,
                "modified": cat_elem.find("modified").text,
                "description": cat_elem.find("description").text if cat_elem.find("description") is not None else ""
            }
        
        for macro_elem in root.find("macros").findall("macro"):
            macro_id = macro_elem.get("id")
            data["macros"][macro_id] = {
                "name": macro_elem.find("name").text,
                "category_id": macro_elem.find("category_id").text,
                "content": macro_elem.find("content").text,
                "created": macro_elem.find("created").text,
                "modified": macro_elem.find("modified").text,
                "version": int(macro_elem.find("version").text)
            }
        
        order_elem = root.find("category_order")
        if order_elem is not None and order_elem.text:
            data["category_order"] = order_elem.text.split(",")
        else:
            data["category_order"] = list(data["categories"].keys())
        
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
        tree.write(macro_data_file_path, encoding="utf-8", xml_declaration=True)
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
    return macros

def get_categories():
    """Return a sorted list of categories from the data."""
    data = load_macro_data()
    order = data.get("category_order", list(data["categories"].keys()))
    names = [data["categories"][cid]["name"] for cid in order if cid in data["categories"]]
    return ["All"] + names

# --- FILE OPERATIONS ---
def copy_macro(macro_key):
    """Copies the content of the specified macro to the clipboard."""
    if macro_key and macro_key in macros_dict:
        try:
            pyperclip.copy(macros_dict[macro_key])
            log_message(f"Copied macro '{macro_key[1]}' to clipboard.")
            print(f"Copied to clipboard: {macro_key[1]}")
        except Exception as e:
            messagebox.showerror("Clipboard Error", f"Could not copy to clipboard: {e}")
            log_message(f"Clipboard error for macro '{macro_key[1]}': {e}")
    elif macro_key:
        log_message(f"Attempted to copy non-existent macro: {macro_key}")
        messagebox.showerror("Error", f"Macro '{macro_key[1]}' not found in current data.")

def open_macro_file():
    """Opens the current macro data file using the OS default text editor."""
    global macro_data_file_path
    if not macro_data_file_path:
        messagebox.showerror("Error", "Cannot open: Macro data file path is not set.")
        return
    if not os.path.exists(macro_data_file_path):
        messagebox.showerror("Error", f"Cannot open: Macro data file not found at:\n{macro_data_file_path}")
        return

    log_message(f"Attempting to open macro data file: {macro_data_file_path}")
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
    
    cat_var = tk.StringVar(value=category_names[0] if category_names else "Uncategorized")
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
    cancel_btn = ctk.CTkButton(btn_frame, text="Cancel", command=popup.destroy, fg_color="gray")
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
    cancel_btn = ctk.CTkButton(btn_frame, text="Cancel", command=popup.destroy, fg_color="gray")
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
                edit_btn = ctk.CTkButton(btn_frame, text="Edit", width=50, **btn_style, command=lambda c=cat_id: edit_category_popup(c))
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

def edit_category_popup(cat_id):
    # Simple popup to edit category name/description
    cat_data = categories[cat_id]
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
        cat_data["name"] = new_name
        cat_data["description"] = new_desc
        cat_data["modified"] = datetime.now().isoformat()
        save_macro_data(data, category_order)
        update_category_list()
        if category_dropdown:
            category_dropdown.configure(values=get_categories())
            category_dropdown.set("All")
        update_list_func()
        popup.destroy()
    save_btn = ctk.CTkButton(popup, text="Save", command=save_edit)
    save_btn.pack(pady=10)

# --- MAIN APPLICATION WINDOW ---
def create_macro_window():
    """Main application window with menu bar, category dropdown, and macro list."""
    global selected_macro_name, macro_list_items, selected_category, macros_dict

    # Initialize macros_dict from XML data
    data = load_macro_data()
    macros_dict = {}
    for macro_id, macro in data["macros"].items():
        cat_id = macro["category_id"]
        cat_name = data["categories"].get(cat_id, {}).get("name", "Uncategorized")
        name = macro["name"]
        content = macro["content"]
        macros_dict[(cat_name, name)] = content

    window = ctk.CTk()
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
    file_menu.add_command(label="Open Data File in Editor", command=open_macro_file)
    file_menu.add_separator()
    file_menu.add_command(label="Change App Icon", command=lambda: change_app_icon(window))
    
    tools_menu = tk.Menu(menubar, tearoff=0)
    menubar.add_cascade(label="Tools", menu=tools_menu)
    tools_menu.add_command(label="Macro Categories", command=lambda: create_category_window(update_list, category_dropdown))
    # Theme submenu
    theme_menu = tk.Menu(tools_menu, tearoff=0)
    tools_menu.add_cascade(label="Theme", menu=theme_menu)
    theme_mode = tk.StringVar(value=load_config().get('theme_mode', 'Dark'))
    theme_menu.add_radiobutton(label="Dark", variable=theme_mode, value="Dark", command=lambda: set_theme("Dark"))
    theme_menu.add_radiobutton(label="Light", variable=theme_mode, value="Light", command=lambda: set_theme("Light"))
    
    help_menu = tk.Menu(menubar, tearoff=0)
    menubar.add_cascade(label="Help", menu=help_menu)
    help_menu.add_command(label="About", command=lambda: messagebox.showinfo("About MacroMouse", "MacroMouse v1.0\nA macro management tool."))

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
    action_button_frame.grid_columnconfigure((0, 1, 2), weight=1)

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

    add_btn = ctk.CTkButton(action_button_frame, text="Add Macro", command=add_action)
    add_btn.grid(row=0, column=0, padx=2, sticky="ew")
    edit_btn = ctk.CTkButton(action_button_frame, text="Edit Macro", command=edit_action)
    edit_btn.grid(row=0, column=1, padx=2, sticky="ew")
    del_btn = ctk.CTkButton(action_button_frame, text="Delete Macro", command=remove_action, fg_color="red")
    del_btn.grid(row=0, column=2, padx=2, sticky="ew")

    close_button = ctk.CTkButton(left_frame, text="Close MacroMouse", command=window.destroy, fg_color="gray", height=35)
    close_button.grid(row=4, column=0, padx=10, pady=(10, 10), sticky="ew")

    preview_frame = ctk.CTkFrame(window)
    preview_frame.grid(row=0, column=1, padx=(0, 10), pady=10, sticky="nsew")
    preview_frame.grid_rowconfigure(1, weight=1)
    preview_frame.grid_columnconfigure(0, weight=1)
    preview_label = ctk.CTkLabel(preview_frame, text="Macro Preview:")
    preview_label.grid(row=0, column=0, padx=10, pady=(0, 5), sticky="w")
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
                show_copied_popup(window, n)
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

    def show_copied_popup(parent, macro_name):
        popup = ctk.CTkToplevel(parent)
        popup.title("Copied")
        popup.geometry("320x90")
        popup.resizable(False, False)
        popup.attributes("-topmost", True)
        popup.grab_set()
        ctk.CTkLabel(popup, text=f"Macro '{macro_name}' copied to clipboard.").pack(pady=18)
        ctk.CTkButton(popup, text="OK", command=popup.destroy, width=80).pack(pady=(0, 10))
        # Center the popup over the parent
        popup.update_idletasks()
        x = parent.winfo_rootx() + (parent.winfo_width() // 2) - (popup.winfo_width() // 2)
        y = parent.winfo_rooty() + (parent.winfo_height() // 2) - (popup.winfo_height() // 2)
        popup.geometry(f"+{x}+{y}")

    update_list()
    window.mainloop()
    
    log_message("Application window closed.")
    log_message("="*20 + " MacroMouse Session End " + "="*20 + "\n")

def main():
    """Main entry point for the application."""
    global macro_data_file_path, log_file_path, config_file_path
    app_data_dir = os.path.join(os.path.expanduser("~"), "MacroMouse_Data")
    os.makedirs(app_data_dir, exist_ok=True)
    macro_data_file_path = os.path.join(app_data_dir, "macros.xml")
    log_file_path = os.path.join(app_data_dir, "MacroMouse.log")
    config_file_path = os.path.join(app_data_dir, "config.json")
    
    data = load_macro_data()
    if not any(cat["name"] == "Uncategorized" for cat in data["categories"].values()):
        create_new_category("Uncategorized")
    config = load_config()
    theme_mode = config.get('theme_mode', 'Dark')
    ctk.set_appearance_mode(theme_mode)
    ctk.set_default_color_theme("dark-blue")
    create_macro_window()

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