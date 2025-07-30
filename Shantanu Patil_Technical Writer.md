# Word Macro: Folder-Based Content Automation

This repository contains a VBA macro for Microsoft Word that automates the insertion and refreshing of content from external `.docx` files using content controls. It is ideal for technical writers and documentation teams working with modular or reusable content.

## ✨ Features
- Detects placeholders like `{{ID}}` in a Word document.
- Searches a selected folder (and subfolders) for matching files.
- Inserts content from those files into the document.
- Wraps inserted content in content controls for easy refresh.
- Allows selective refresh of content by ID.
- Navigate to source files directly from Word.

## 📂 Files
- `InsertContentWithContentControls.vba`: The main macro script.
- `README.md`: This documentation.

## 🚀 How to Use
1. Open Microsoft Word.
2. Press `Alt + F11` to open the VBA editor.
3. Personally created macro assigned `Buttons` can be use for easy run.
4. Insert a new module and paste the contents of [`Word-Automated-content-control-and-content-Insertion`](https://github.com/TejasGiri-TD/tejasgiri-portfolio/tree/Word-Automated-content-control-and-content-Insertion).
5. Run the macro `InsertContentWithContentControls` to begin.
6. Run the macro `RefreshSelectedIDsWithPopup` to refresh the content of user specified IDs to retrive the chsnges in the source.
7. Run the macro `NavigateToSourceFile` to navigate the source document.
8. Run the macro `ExportIDAndNameToExcel` to get the list of IDs used in the document.

## 🛡️ Privacy
This macro contains **no personal information** and is safe to use or modify.

## 📄 License
MIT License – Free to use, modify, and distribute.

---

### 👨‍💻 Author
Created by Tejas Giri  
[LinkedIn](https://www.linkedin.com/in/tejas-giri-96151524)
