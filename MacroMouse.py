import customtkinter as ctk
import tkinter as tk # Still need tkinter for StringVar
from tkinter import messagebox, filedialog
import pyperclip
import subprocess
import os
import sys
import shutil # Used for copy/delete fallback during file move
from datetime import datetime

# --- Globals ---
macro_file_path = None
log_file_path = None
# Keep selected_macro_name accessible within create_macro_window scope
selected_macro_name = None
# Keep a reference to list items (buttons) for easy clearing/updating
macro_list_items = [] # Stores the actual CTkButton widgets in the list

# --- Core Logic Functions ---

def get_log_timestamp():
    """Return a simple timestamp string for logging."""
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

def log_message(message):
    """Append a log message to the log file."""
    global log_file_path
    if log_file_path:
        try:
            # Ensure directory exists before writing log
            log_dir = os.path.dirname(log_file_path)
            if log_dir and not os.path.exists(log_dir):
                 os.makedirs(log_dir, exist_ok=True)
            with open(log_file_path, "a", encoding="utf-8") as logf:
                logf.write(f"[{get_log_timestamp()}] {message}\n")
        except Exception as e:
            # Non-critical error, print to console instead of showing popup
            print(f"Warning: could not write log file '{log_file_path}': {e}")

def load_macros():
    """Loads macro data from the currently set macro_file_path."""
    global macro_file_path
    macros = {}
    if not macro_file_path or not os.path.exists(macro_file_path):
        # Log this, but don't show popup here; handled in main() or file ops
        log_message(f"Macro file not found or path not set during load: {macro_file_path}")
        return macros # Return empty dict

    try:
        with open(macro_file_path, "r", encoding="utf-8") as f:
            current_macro_name = None
            current_macro_content = ""
            for line in f:
                stripped_line = line.strip()
                if stripped_line.startswith("#"):
                    # Save previous macro if one was being built
                    if current_macro_name:
                        macros[current_macro_name] = current_macro_content.strip()
                    # Start new macro
                    current_macro_name = stripped_line[1:].strip()
                    current_macro_content = "" # Reset content
                elif current_macro_name is not None: # Ensure we have a macro name before appending
                    # Append line to current macro content (preserve original line breaks)
                    current_macro_content += line # Keep original newline from file
            # Save the last macro in the file
            if current_macro_name is not None: # Check again for last macro
                macros[current_macro_name] = current_macro_content.strip()
        log_message(f"Successfully loaded {len(macros)} macros from: {macro_file_path}")
        return macros
    except Exception as e:
        messagebox.showerror("Error Loading File", f"Error loading macros from {macro_file_path}:\n{e}")
        log_message(f"Error loading macros: {e}")
        return None # Return None on critical error to distinguish from empty file

def save_macros(macros_to_save):
    """Saves the provided macro dictionary to the current macro_file_path."""
    global macro_file_path
    if not macro_file_path:
        messagebox.showerror("Error", "Cannot save: Macro file path is not set.")
        return False

    try:
        # Ensure directory exists before writing
        macro_dir = os.path.dirname(macro_file_path)
        if macro_dir and not os.path.exists(macro_dir):
            os.makedirs(macro_dir, exist_ok=True)
            log_message(f"Created directory for macro file: {macro_dir}")

        with open(macro_file_path, "w", encoding="utf-8") as f:
            # Sort macros by name before saving for consistency
            for name in sorted(macros_to_save.keys()):
                content = macros_to_save[name]
                f.write(f"# {name}\n") # Write macro name marker
                # Write content exactly as stored (should preserve newlines)
                f.write(f"{content}\n\n") # Add blank line for separation
        log_message(f"Saved {len(macros_to_save)} macros to: {macro_file_path}")
        return True
    except Exception as e:
        messagebox.showerror("Error Saving File", f"Error saving macros to {macro_file_path}:\n{e}")
        log_message(f"Error saving macros: {e}")
        return False

def copy_macro(macro_name_to_copy, macros_dict):
    """Copies the content of the specified macro to the clipboard."""
    if macro_name_to_copy and macro_name_to_copy in macros_dict:
        try:
            pyperclip.copy(macros_dict[macro_name_to_copy])
            log_message(f"Copied macro '{macro_name_to_copy}' to clipboard.")
            # Optional: Add status bar feedback instead of print
            print(f"Copied to clipboard: {macro_name_to_copy}")
        except Exception as e:
             # Pyperclip can sometimes fail
             messagebox.showerror("Clipboard Error", f"Could not copy to clipboard: {e}")
             log_message(f"Clipboard error for macro '{macro_name_to_copy}': {e}")
    elif macro_name_to_copy:
        # Should not happen if selection logic is correct, but handle anyway
        log_message(f"Attempted to copy non-existent macro: {macro_name_to_copy}")
        messagebox.showerror("Error", f"Macro '{macro_name_to_copy}' not found in current data.")

def open_macro_file():
    """Opens the current macro file using the OS default text editor."""
    global macro_file_path
    if not macro_file_path:
        messagebox.showerror("Error", "Cannot open: Macro file path is not set.")
        return
    if not os.path.exists(macro_file_path):
         messagebox.showerror("Error", f"Cannot open: Macro file not found at:\n{macro_file_path}")
         return

    log_message(f"Attempting to open macro file: {macro_file_path}")
    try:
        if sys.platform.startswith('win32'):
            os.startfile(macro_file_path) # Recommended way on Windows
        elif sys.platform.startswith('darwin'):
            subprocess.run(['open', macro_file_path], check=True) # macOS
        else: # Assume Linux/other Unix-like
            subprocess.run(['xdg-open', macro_file_path], check=True) # Standard freedesktop way
        log_message("Successfully launched default editor for macro file.")
    except FileNotFoundError:
         # This might mean the file exists but the associated application doesn't
         messagebox.showerror("Error", f"Could not find an application associated with text files, or the file path is invalid:\n{macro_file_path}")
         log_message(f"Error opening macro file - File or associated application not found: {macro_file_path}")
    except Exception as e:
        messagebox.showerror("Error Opening File", f"An error occurred while trying to open the macro file:\n{e}")
        log_message(f"Error opening macro file: {e}")

# --- File Location Management Functions ---

def use_new_macro_file(macros_dict, update_list_func):
    """
    Prompts user for a parent directory, creates a 'MacroMouse' subfolder,
    initializes new 'macros.txt' and 'MacroMouse.log' files there,
    updates global paths, and reloads the UI.
    """
    global macro_file_path, log_file_path

    # Suggest starting directory (e.g., User's Documents)
    initial_dir = os.path.expanduser("~")
    if sys.platform.startswith('win32'):
        docs = os.path.join(initial_dir, "Documents")
        desktop = os.path.join(initial_dir, "Desktop")
        initial_dir = docs if os.path.exists(docs) else (desktop if os.path.exists(desktop) else initial_dir)

    folder_selected = filedialog.askdirectory(
        title="Select Parent Folder for New Macro File Location",
        initialdir=initial_dir
        )
    if not folder_selected:
        log_message("User cancelled 'Use New Macro File' dialog.")
        return

    log_message(f"User selected new parent folder: {folder_selected}")
    target_folder = os.path.join(folder_selected, "MacroMouse") # Standard subfolder name

    try:
        os.makedirs(target_folder, exist_ok=True) # Create structure if needed
        log_message(f"Ensured target directory exists: {target_folder}")
    except Exception as e:
         messagebox.showerror("Directory Error", f"Could not create directory:\n{target_folder}\n\nError: {e}")
         log_message(f"Failed to create target directory {target_folder}: {e}")
         return

    # Define new file paths
    new_macro_path = os.path.join(target_folder, "macros.txt")
    new_log_path = os.path.join(target_folder, "MacroMouse.log")

    # Create/overwrite the macros file with a starter message
    try:
        with open(new_macro_path, "w", encoding="utf-8") as f:
            f.write("# Welcome to MacroMouse!\nAdd your first macro using the 'Add' button.\n\n")
        log_message(f"Created new macro file: {new_macro_path}")
    except Exception as e:
         messagebox.showerror("File Creation Error", f"Could not create new macro file:\n{new_macro_path}\n\nError: {e}")
         log_message(f"Failed to create new macro file {new_macro_path}: {e}")
         return

    # Create/overwrite the log file
    try:
        with open(new_log_path, "w", encoding="utf-8") as lf:
             lf.write(f"[{get_log_timestamp()}] Created/Reset MacroMouse.log at new location: {new_log_path}\n")
        # Update global log path *only after* successfully writing the new log
        log_file_path = new_log_path
        log_message(f"Created/Reset log file: {new_log_path}") # Log this using the *new* log file
    except Exception as e:
        messagebox.showwarning("Log File Warning", f"Could not create/clear log file:\n{new_log_path}\n\nError: {e}")
        log_message(f"Failed to create/reset log file {new_log_path}: {e}") # Log failure using old log path
        log_file_path = None # Invalidate if creation failed

    # Update the global macro path *after* successful file creation
    macro_file_path = new_macro_path

    messagebox.showinfo("New Macro File Set", f"MacroMouse will now use files in:\n{target_folder}")

    # Clear current data, load from new file, and update UI
    macros_dict.clear()
    reloaded_data = load_macros()
    if reloaded_data is not None:
        macros_dict.update(reloaded_data)
    update_list_func() # Refresh the list (should show the example)

def change_macro_file_location(macros_dict, update_list_func):
    """
    Moves the existing macros.txt and log file to a new parent directory
    selected by the user (within a 'MacroMouse' subfolder).
    Starts the dialog in the current file's directory.
    """
    global macro_file_path, log_file_path

    if not macro_file_path or not os.path.exists(macro_file_path):
        messagebox.showerror("Error", "Cannot move: No valid macro file currently set or found.")
        return

    current_macro_dir = os.path.dirname(macro_file_path)
    current_macro_filename = os.path.basename(macro_file_path)
    current_log_filename = os.path.basename(log_file_path) if log_file_path and os.path.basename(log_file_path) else "MacroMouse.log"

    # Start the directory chooser in the current macro file's location
    folder_selected = filedialog.askdirectory(
        title="Select New Parent Folder (a 'MacroMouse' subfolder will be created/used)",
        initialdir=current_macro_dir
        )
    if not folder_selected:
        log_message("User cancelled 'Move File' dialog.")
        return

    log_message(f"User selected new parent folder for move: {folder_selected}")

    target_folder = os.path.join(folder_selected, "MacroMouse")
    new_macro_path = os.path.join(target_folder, current_macro_filename)
    new_log_path = os.path.join(target_folder, current_log_filename)

    if os.path.abspath(new_macro_path) == os.path.abspath(macro_file_path):
         messagebox.showinfo("Info", "The selected location is the same as the current location. Files not moved.")
         log_message("Move operation cancelled: Target location is same as source.")
         return

    if os.path.exists(new_macro_path) or os.path.exists(new_log_path):
        existing = []
        if os.path.exists(new_macro_path): existing.append(os.path.basename(new_macro_path))
        if os.path.exists(new_log_path): existing.append(os.path.basename(new_log_path))
        if not messagebox.askyesno("Confirm Overwrite",
                                   f"The file(s) '{', '.join(existing)}' already exist(s) in the target 'MacroMouse' subfolder.\n\nOverwrite?"):
            log_message("Move operation cancelled: User chose not to overwrite existing files.")
            return

    try:
        os.makedirs(target_folder, exist_ok=True)
        log_message(f"Ensured target directory exists for move: {target_folder}")
    except Exception as e:
        messagebox.showerror("Directory Error", f"Could not create target directory:\n{target_folder}\n\nError: {e}")
        log_message(f"Failed to create target directory {target_folder} for move: {e}")
        return

    original_macro_path = macro_file_path
    original_log_path = log_file_path
    macro_moved = False
    log_moved = False

    try:
        if not save_macros(macros_dict):
             raise Exception("Failed to save current macros before moving. Aborting move.")

        # Move Macro File
        try:
            os.rename(original_macro_path, new_macro_path)
            log_message(f"Successfully renamed macro file to: {new_macro_path}")
            macro_moved = True
        except OSError:
             log_message(f"os.rename failed for macro file. Attempting shutil.copy2/os.remove...")
             shutil.copy2(original_macro_path, new_macro_path)
             os.remove(original_macro_path)
             log_message(f"Successfully copied/deleted macro file to: {new_macro_path}")
             macro_moved = True

        if macro_moved:
            macro_file_path = new_macro_path

        # Move Log File
        if original_log_path and os.path.exists(original_log_path):
            try:
                os.rename(original_log_path, new_log_path)
                log_message(f"Successfully renamed log file to: {new_log_path}")
                log_moved = True
            except OSError:
                log_message(f"os.rename failed for log file. Attempting shutil.copy2/os.remove...")
                shutil.copy2(original_log_path, new_log_path)
                os.remove(original_log_path)
                log_message(f"Successfully copied/deleted log file to: {new_log_path}")
                log_moved = True

            if log_moved:
                log_file_path = new_log_path
        else:
             log_message("Original log file not found or path not set, skipping log move.")
             if macro_moved: log_file_path = new_log_path

        if macro_moved:
             messagebox.showinfo("Move Successful", f"Files moved successfully to:\n{target_folder}")
             log_message(f"Move operation completed. New macro path: {macro_file_path}, New log path: {log_file_path}")
             macros_dict.clear()
             reloaded_data = load_macros()
             if reloaded_data is not None:
                 macros_dict.update(reloaded_data)
             update_list_func()
        else:
             messagebox.showerror("Move Failed", "Could not move the macro file. Operation aborted.")

    except Exception as e:
        macro_file_path = original_macro_path
        log_file_path = original_log_path
        messagebox.showerror("Error During Move", f"An error occurred: {e}\n\nPaths have been reset to original locations. Please check file permissions and ensure files are not open.")
        log_message(f"Error during move process: {e}. Attempted to restore original paths.")
        macros_dict.clear()
        reloaded_data = load_macros() # Reload original data
        if reloaded_data is not None:
            macros_dict.update(reloaded_data)
        update_list_func()

# --- Popup Windows ---

def add_macro_popup(macros_dict, update_list_func):
    """Popup to add a new macro using CustomTkinter."""
    popup = ctk.CTkToplevel()
    popup.title("Add New Macro")
    popup.geometry("450x380")
    popup.minsize(400, 350)
    popup.grab_set()
    popup.attributes('-topmost', True)
    popup.after(150, lambda: popup.attributes('-topmost', False))

    input_frame = ctk.CTkFrame(popup)
    input_frame.pack(pady=10, padx=10, fill="both", expand=True)
    name_label = ctk.CTkLabel(input_frame, text="Macro Name:")
    name_label.pack(pady=(5, 0), padx=10, anchor="w")
    name_entry = ctk.CTkEntry(input_frame)
    name_entry.pack(pady=(0, 10), padx=10, fill="x")
    name_entry.focus()
    content_label = ctk.CTkLabel(input_frame, text="Macro Content:")
    content_label.pack(pady=(5, 0), padx=10, anchor="w")
    content_text = ctk.CTkTextbox(input_frame, font=("Consolas", 11))
    content_text.pack(pady=(0, 10), padx=10, fill="both", expand=True)

    button_frame = ctk.CTkFrame(popup, fg_color="transparent")
    button_frame.pack(pady=(0, 10), padx=10, fill="x", side="bottom")

    def add_macro_action():
        name = name_entry.get().strip()
        content = content_text.get("1.0", "end-1c").strip()
        if not name: messagebox.showerror("Input Error", "Macro Name cannot be empty.", parent=popup); return
        if not content: messagebox.showerror("Input Error", "Macro Content cannot be empty.", parent=popup); return
        if name in macros_dict: messagebox.showerror("Name Conflict", f"Macro name '{name}' already exists.", parent=popup); return

        macros_dict[name] = content
        if save_macros(macros_dict):
            log_message(f"Added macro '{name}'.")
            update_list_func(name) # Refresh list and select new item
            popup.destroy()
        else:
            if name in macros_dict: del macros_dict[name] # Revert change in memory if save failed
            messagebox.showerror("Save Error", "Could not save macros to file. Macro not added.", parent=popup)

    def cancel_action(): popup.destroy()

    add_button = ctk.CTkButton(button_frame, text="Add Macro", command=add_macro_action)
    add_button.pack(side="left", padx=(0, 10))
    cancel_button = ctk.CTkButton(button_frame, text="Cancel", command=cancel_action, fg_color="gray")
    cancel_button.pack(side="right")

    name_entry.bind("<Return>", lambda event: content_text.focus())
    content_text.bind("<Control-Return>", lambda event: add_macro_action())
    popup.protocol("WM_DELETE_WINDOW", cancel_action)

def edit_macro_popup(macros_dict, macro_name_to_edit, update_list_func):
    """Popup to edit an existing macro using CustomTkinter."""
    if not macro_name_to_edit or macro_name_to_edit not in macros_dict:
        messagebox.showerror("Error", "Invalid or non-existent macro selected for editing.")
        return

    popup = ctk.CTkToplevel()
    popup.title(f"Edit Macro: {macro_name_to_edit}")
    popup.geometry("450x380")
    popup.minsize(400, 350)
    popup.grab_set()
    popup.attributes('-topmost', True)
    popup.after(150, lambda: popup.attributes('-topmost', False))

    original_name = macro_name_to_edit
    original_content = macros_dict[original_name]

    input_frame = ctk.CTkFrame(popup)
    input_frame.pack(pady=10, padx=10, fill="both", expand=True)
    name_label = ctk.CTkLabel(input_frame, text="Macro Name:")
    name_label.pack(pady=(5, 0), padx=10, anchor="w")
    name_entry = ctk.CTkEntry(input_frame)
    name_entry.insert(0, original_name)
    name_entry.pack(pady=(0, 10), padx=10, fill="x")
    content_label = ctk.CTkLabel(input_frame, text="Macro Content:")
    content_label.pack(pady=(5, 0), padx=10, anchor="w")
    content_text = ctk.CTkTextbox(input_frame, font=("Consolas", 11))
    content_text.insert("1.0", original_content)
    content_text.pack(pady=(0, 10), padx=10, fill="both", expand=True)
    content_text.focus()

    button_frame = ctk.CTkFrame(popup, fg_color="transparent")
    button_frame.pack(pady=(0, 10), padx=10, fill="x", side="bottom")

    def save_changes_action():
        new_name = name_entry.get().strip()
        new_content = content_text.get("1.0", "end-1c").strip()
        if not new_name: messagebox.showerror("Input Error", "Macro Name cannot be empty.", parent=popup); return
        if not new_content: messagebox.showerror("Input Error", "Macro Content cannot be empty.", parent=popup); return
        if new_name != original_name and new_name in macros_dict: messagebox.showerror("Name Conflict", f"The name '{new_name}' already exists.", parent=popup); return

        temp_macros = macros_dict.copy()
        if new_name != original_name:
            if original_name in temp_macros: del temp_macros[original_name]
            else: log_message(f"Warning: Original key '{original_name}' not found during edit rename.")
        temp_macros[new_name] = new_content

        if save_macros(temp_macros):
            macros_dict.clear(); macros_dict.update(temp_macros) # Update main dict on success
            log_message(f"Edited macro '{original_name}' (now '{new_name}').")
            update_list_func(new_name) # Refresh list and select edited item
            popup.destroy()
        else:
             messagebox.showerror("Save Error", "Could not save changes to macro file.", parent=popup)

    def cancel_action(): popup.destroy()

    save_button = ctk.CTkButton(button_frame, text="Save Changes", command=save_changes_action)
    save_button.pack(side="left", padx=(0, 10))
    cancel_button = ctk.CTkButton(button_frame, text="Cancel", command=cancel_action, fg_color="gray")
    cancel_button.pack(side="right")

    name_entry.bind("<Return>", lambda event: content_text.focus())
    content_text.bind("<Control-Return>", lambda event: save_changes_action())
    popup.protocol("WM_DELETE_WINDOW", cancel_action)

# --- Refresh Function ---

def refresh_data_from_file(macros_dict, update_list_func):
    """
    Clears the current macro dictionary, reloads from the macro file on disk,
    and updates the UI list. Used after potential external file edits.
    """
    global selected_macro_name # Access global to clear selection if needed

    log_message("User initiated refresh from file.")
    print("Refreshing list from file...") # Console feedback

    # Store current selection to potentially restore later
    current_selection = selected_macro_name

    # Reload data from the file
    reloaded_macros = load_macros()

    # Check if loading failed (load_macros handles message boxes / returns None)
    if reloaded_macros is None or not isinstance(reloaded_macros, dict):
        log_message("Refresh failed because file loading returned invalid data.")
        print("Refresh failed.")
        # Clear the dict even if load failed, to show empty state
        macros_dict.clear()
    else:
        # Update the main dictionary only if load was successful
        macros_dict.clear()
        macros_dict.update(reloaded_macros)

    # Clear selection state before updating list, let update_list handle re-selection
    selected_macro_name = None

    # Update the UI list
    # Try to restore selection if the item still exists after reload
    update_list_func(current_selection if current_selection in macros_dict else None)

    log_message(f"Refresh complete. Loaded {len(macros_dict)} macros.")
    print(f"Refresh complete. Found {len(macros_dict)} macros.")

# --- Main Application Window ---

def create_macro_window(macros_dict):
    """Creates the main GUI window using CustomTkinter."""
    global selected_macro_name, macro_list_items # Allow modification

    window = ctk.CTk()
    window.title("MacroMouse")
    window.geometry("850x600+150+100")
    window.minsize(650, 500)
    window.resizable(True, True)

    window.lift()
    window.attributes('-topmost', True)
    window.after(200, lambda: window.attributes('-topmost', False))

    def on_close():
        log_message("Close button clicked. Exiting application.")
        window.quit()
    window.protocol("WM_DELETE_WINDOW", on_close)

    # --- Configure Grid Layout ---
    window.grid_columnconfigure(0, weight=1, minsize=280) # Left column
    window.grid_columnconfigure(1, weight=3)             # Right column
    window.grid_rowconfigure(0, weight=1)                # Allow row expansion

    # --- Left Frame (Controls and List) ---
    left_frame = ctk.CTkFrame(window, fg_color="transparent")
    left_frame.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")
    left_frame.grid_columnconfigure(0, weight=1)
    # Row Indexing: 0:Search, 1:List, 2:Add/Edit/Del, 3:File Mgmt, 4:Close
    left_frame.grid_rowconfigure(1, weight=1) # Make list area expand

    # --- Search Row (Row 0) ---
    search_frame = ctk.CTkFrame(left_frame, fg_color="transparent")
    search_frame.grid(row=0, column=0, padx=10, pady=(0, 5), sticky="ew")
    search_frame.grid_columnconfigure(0, weight=1)
    search_label = ctk.CTkLabel(search_frame, text="Search Macros:")
    search_label.grid(row=0, column=0, sticky="w")
    search_var = tk.StringVar()
    search_entry = ctk.CTkEntry(search_frame, textvariable=search_var)
    search_entry.grid(row=1, column=0, pady=(0, 5), sticky="ew")

    # --- Macro List Area (Row 1) ---
    macro_list_frame = ctk.CTkScrollableFrame(left_frame, label_text="Available Macros")
    macro_list_frame.grid(row=1, column=0, padx=10, pady=0, sticky="nsew")
    macro_list_frame.grid_columnconfigure(0, weight=1)

    # --- Preview Area (Right Side) ---
    preview_frame = ctk.CTkFrame(window)
    preview_frame.grid(row=0, column=1, padx=(0, 10), pady=10, sticky="nsew")
    preview_frame.grid_rowconfigure(1, weight=1)
    preview_frame.grid_columnconfigure(0, weight=1)
    preview_label = ctk.CTkLabel(preview_frame, text="Macro Preview:")
    preview_label.grid(row=0, column=0, padx=10, pady=(0, 5), sticky="w")
    preview_text = ctk.CTkTextbox(preview_frame, wrap="word", font=("Consolas", 11))
    preview_text.grid(row=1, column=0, padx=10, pady=(0, 10), sticky="nsew")
    preview_text.configure(state="disabled")

    # --- Add/Edit/Remove Button Area (Row 2) ---
    action_button_frame = ctk.CTkFrame(left_frame, fg_color="transparent")
    action_button_frame.grid(row=2, column=0, padx=10, pady=(10, 5), sticky="ew")
    action_button_frame.grid_columnconfigure((0, 1), weight=1)

    # --- File Management Buttons Frame (Row 3) ---
    file_mgmt_frame = ctk.CTkFrame(left_frame)
    file_mgmt_frame.grid(row=3, column=0, padx=10, pady=(10, 5), sticky="ew")
    file_mgmt_frame.grid_columnconfigure(0, weight=1) # Allow expansion via columnspan

    # --- UI Interaction Functions --- Needed by buttons/list below ---
    def update_preview(selected_name_preview):
        preview_text.configure(state="normal")
        preview_text.delete("1.0", "end")
        if selected_name_preview and selected_name_preview in macros_dict:
            preview_text.insert("1.0", macros_dict[selected_name_preview])
        preview_text.configure(state="disabled")

    def highlight_selected_item(button_to_highlight):
         selected_color = ctk.ThemeManager.theme["CTkButton"]["hover_color"]
         for item_widget in macro_list_items:
             if isinstance(item_widget, ctk.CTkButton):
                 item_widget.configure(border_width=0)
         if button_to_highlight and isinstance(button_to_highlight, ctk.CTkButton):
             button_to_highlight.configure(border_width=1, border_color=selected_color)

    def on_macro_select(name, button_widget):
        global selected_macro_name
        if not isinstance(button_widget, ctk.CTkButton): return
        selected_macro_name = name
        update_preview(selected_macro_name)
        highlight_selected_item(button_widget)

    def on_macro_double_click(name): copy_macro(name, macros_dict)
    def clear_macro_list_widgets():
         global macro_list_items
         for widget in macro_list_items: widget.destroy()
         macro_list_items = []

    def update_list(select_name_after_update=None):
        global selected_macro_name, macro_list_items
        clear_macro_list_widgets()
        search_term = search_var.get().lower()
        sorted_names = sorted(macros_dict.keys(), key=str.lower)
        filtered_names = [name for name in sorted_names if search_term in name.lower()]
        selected_button_widget = None; newly_selected_name = None

        if not filtered_names:
             no_results_label = ctk.CTkLabel(macro_list_frame, text="No matching macros found.", text_color="gray")
             no_results_label.pack(pady=5); macro_list_items.append(no_results_label)
             selected_macro_name = None; update_preview(None); highlight_selected_item(None)
        else:
            for name in filtered_names:
                macro_button = ctk.CTkButton(macro_list_frame, text=name, anchor="w", fg_color="transparent", hover=False, border_width=0, text_color=ctk.ThemeManager.theme["CTkLabel"]["text_color"], command=lambda n=name, b=None: on_macro_select(n, b))
                macro_button.configure(command=lambda n=name, b=macro_button: on_macro_select(n, b))
                macro_button.pack(fill="x", pady=(0, 1), padx=1)
                macro_button.bind("<Double-Button-1>", lambda event, n=name: on_macro_double_click(n))
                macro_list_items.append(macro_button)
                if select_name_after_update and name == select_name_after_update: selected_button_widget = macro_button; newly_selected_name = name
                elif not select_name_after_update and selected_macro_name and name == selected_macro_name: selected_button_widget = macro_button; newly_selected_name = name

        if newly_selected_name and selected_button_widget:
            selected_macro_name = newly_selected_name
            highlight_selected_item(selected_button_widget)
            update_preview(selected_macro_name)
            window.after(50, lambda sb=selected_button_widget: macro_list_frame.scroll_to_widget(sb))
        elif not newly_selected_name:
            selected_macro_name = None; update_preview(None); highlight_selected_item(None)

    # --- Action Handlers for Buttons ---
    def edit_action():
        if selected_macro_name: edit_macro_popup(macros_dict, selected_macro_name, update_list)
        else: messagebox.showwarning("Select Macro", "Please select a macro from the list to edit.")
    def remove_action():
         global selected_macro_name
         if selected_macro_name:
             if messagebox.askyesno("Confirm Delete", f"Are you sure you want to permanently remove the macro '{selected_macro_name}'?"):
                 if selected_macro_name in macros_dict:
                     del macros_dict[selected_macro_name]
                     if save_macros(macros_dict): log_message(f"Removed macro '{selected_macro_name}'."); selected_macro_name = None; update_list()
                     else: messagebox.showerror("Error", "Failed to save changes after removing macro."); macros_dict.clear(); reloaded=load_macros(); macros_dict.update(reloaded if reloaded is not None else {}); update_list()
                 else: messagebox.showerror("Error", "Selected macro seems to have already been removed."); selected_macro_name = None; update_list()
         else: messagebox.showwarning("Select Macro", "Please select a macro from the list to remove.")

    # --- Create and Place Action Buttons ---
    add_macro_button = ctk.CTkButton(action_button_frame, text="Add", command=lambda: add_macro_popup(macros_dict, update_list))
    add_macro_button.grid(row=0, column=0, padx=(0, 5), pady=5, sticky="ew")
    edit_macro_button = ctk.CTkButton(action_button_frame, text="Edit", command=edit_action)
    edit_macro_button.grid(row=0, column=1, padx=(5, 0), pady=5, sticky="ew")
    remove_macro_button = ctk.CTkButton(action_button_frame, text="Remove", command=remove_action, fg_color="#D32F2F", hover_color="#B71C1C")
    remove_macro_button.grid(row=1, column=0, columnspan=2, padx=0, pady=(5, 10), sticky="ew")

    # --- Create and Place File Management Buttons ---
    use_new_button = ctk.CTkButton(file_mgmt_frame, text="Use New File...", command=lambda: use_new_macro_file(macros_dict, update_list), height=35)
    use_new_button.grid(row=0, column=0, padx=(0, 5), pady=5, sticky="ew")
    change_location_button = ctk.CTkButton(file_mgmt_frame, text="Move File Location...", command=lambda: change_macro_file_location(macros_dict, update_list), height=35)
    change_location_button.grid(row=0, column=1, padx=(5, 0), pady=5, sticky="ew")
    open_button = ctk.CTkButton(file_mgmt_frame, text="Open File in Editor", command=open_macro_file, height=35)
    open_button.grid(row=1, column=0, columnspan=2, padx=0, pady=5, sticky="ew")
    refresh_list_button = ctk.CTkButton(file_mgmt_frame, text="Refresh Macro List", command=lambda: refresh_data_from_file(macros_dict, update_list), height=35)
    refresh_list_button.grid(row=2, column=0, columnspan=2, padx=0, pady=(5, 10), sticky="ew") # Placed below Open button

    # --- Close Button (Row 4 - Main Left Frame Grid) ---
    close_button = ctk.CTkButton(left_frame, text="Close MacroMouse", command=on_close, fg_color="gray", height=35)
    close_button.grid(row=4, column=0, padx=10, pady=(10, 10), sticky="ew")

    # --- Initial Setup & Bindings ---
    search_var.trace_add("write", lambda name, index, mode: update_list())
    update_list() # Initial population

    window.mainloop() # Start the Tkinter event loop

# --- Main Execution ---

def main():
    """Main function: Sets up paths, loads data, and launches the GUI."""
    global macro_file_path, log_file_path

    # Determine Application Path
    if getattr(sys, 'frozen', False): base_path = os.path.dirname(sys.executable)
    elif __file__: base_path = os.path.dirname(os.path.abspath(__file__))
    else: base_path = os.path.abspath(".")

    # Define Data Storage Location
    data_subfolder = "MacroMouse_Data"
    app_data_path = os.path.join(base_path, data_subfolder)

    # Define Default File Paths
    default_macro_filename = "macros.txt"
    default_log_filename = "MacroMouse.log"
    default_macro_path = os.path.join(app_data_path, default_macro_filename)
    default_log_path = os.path.join(app_data_path, default_log_filename)

    # Ensure Data Directory Exists
    try: os.makedirs(app_data_path, exist_ok=True)
    except Exception as e: messagebox.showerror("Fatal Error", f"Could not create data directory:\n{app_data_path}\nError: {e}\nExiting."); sys.exit(1)

    # Set Global Paths
    macro_file_path = default_macro_path
    log_file_path = default_log_path

    # Initialize Logging
    log_message("="*20 + " MacroMouse Session Start " + "="*20)
    log_message(f"Base path: {base_path}")
    log_message(f"Data path: {app_data_path}")
    log_message(f"Initial macro file: {macro_file_path}")
    log_message(f"Initial log file: {log_file_path}")

    # Load Initial Macros
    macros_data = {}
    if not os.path.exists(macro_file_path):
        log_message(f"Default macro file not found: {macro_file_path}")
        answer = messagebox.askyesno("Setup: Macro File", f"Macro file ({default_macro_filename}) not found in:\n{app_data_path}\n\nCreate a new file here?")
        if answer:
            try:
                with open(macro_file_path, "w", encoding="utf-8") as f: f.write("# Welcome!\nAdd macros using the 'Add' button.\n\n")
                log_message("Created new empty macro file."); macros_data = load_macros()
            except Exception as e: messagebox.showerror("Error", f"Failed to create macro file: {e}"); log_message(f"Failed to create default macro file: {e}")
        else: log_message("User chose not to create default file."); messagebox.showinfo("Info", "Starting without macros.\nUse UI buttons to manage files.")
    else:
        log_message("Default macro file found. Loading...")
        loaded = load_macros()
        if loaded is not None: macros_data = loaded # Only assign if load didn't fail critically

    # Launch GUI
    log_message("Starting main application window.")
    create_macro_window(macros_data) # Pass loaded data
    log_message("Application window closed.")
    log_message("="*20 + " MacroMouse Session End " + "="*20 + "\n")

if __name__ == "__main__":
    # Set high DPI awareness for Windows if possible
    if sys.platform.startswith('win32'):
        try:
            ctk.windll.shcore.SetProcessDpiAwareness(2) # PROCESS_PER_MONITOR_DPI_AWARE
            print("Note: Set DPI awareness to Per-Monitor Aware V2.")
        except Exception as e: print(f"Note: Could not set DPI awareness (may require Windows 8.1+): {e}")

    # Set theme/appearance *before* creating any CTk widgets
    ctk.set_appearance_mode("dark")
    ctk.set_default_color_theme("dark-blue")

    main()