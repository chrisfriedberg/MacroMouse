# <img width="50" height="50" alt="MMouseFAV" src="https://github.com/user-attachments/assets/45244a2e-d909-4722-8956-7da214cf35a0" />MacroMouse  <img width="50" height="50" alt="MMouseFAV" src="https://github.com/user-attachments/assets/45244a2e-d909-4722-8956-7da214cf35a0" />
**A Powerful Desktop Macro Management Tool**  

MacroMouse is a feature-rich Python desktop application for creating, organizing, and executing text macros with ease. It’s built with [CustomTkinter](https://github.com/TomSchimansky/CustomTkinter) for a modern UI and supports advanced placeholder handling, dynamic values, usage tracking, category management, and data backup/restore.  

---

## ✨ Features  

- **Macro Creation & Editing**  
  - Create text snippets (macros) organized by categories.  
  - Edit, duplicate, or delete existing macros.  

- **Dynamic Placeholders**  
  - Use `<date>`, `<time>`, `<datetime>` and other tags that automatically update.  
  - Custom `{{placeholders}}` that prompt for user input before copying.  

- **Category Management**  
  - Add, rename, hide/unhide, and reorder categories.  
  - Keep macros organized for quick access.  

- **Macro Usage Tracking**  
  - Tracks most-used macros and promotes them for quick access.  
  - Reset or delete usage statistics as needed.  

- **Clipboard Integration**  
  - Copy macro content instantly to your clipboard with one click.  

- **Backup & Restore**  
  - Backup your entire macro dataset (XML format) to a ZIP archive.  
  - Restore from backups or replace individual macro files.  

- **Customizable UI**  
  - Dark/Light themes.  
  - Change application icon.  

- **Reference File Support**  
  - Link and quickly open a reference file alongside your macros.  

---

## 🖥️ Screenshots  

<img width="743" height="547" alt="image" src="https://github.com/user-attachments/assets/8613e9ef-8826-483e-8b76-b617485e1534" />


---

## 📦 Installation  

**Requirements:**  
- Python 3.9+  
- `pip`  

**Install dependencies:**  
```bash
pip install customtkinter pyperclip pystray pillow
