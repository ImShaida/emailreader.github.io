# Email Reader
A lightweight Python/Tkinter application to read .eml and .msg email files, view headers/body, and extract attachments — with optional HTML preview and one-click Windows EXE packaging.

# Email Reader & Attachment Extractor

A simple Python/Tkinter app to open `.eml` and `.msg` files, read headers/body, and extract attachments.

## Features
- Open `.eml` files (native support).
- Optional `.msg` support via [`extract_msg`](https://pypi.org/project/extract-msg/).
- View plain text and HTML body.
- Save body as `.txt`.
- Extract one or all attachments.
- Open HTML body in your web browser.
- View raw RFC email source.

---

## Installation
1. Install Python 3.8+ (with Tkinter).
2. Clone the repo:

   ```bash
     git clone https://github.com/zrnge/emailreader.git
     cd email-reader
     pip install -r requirements.txt


Converting to a Windows EXE
   
You can create a standalone .exe so users don’t need Python installed.

Install PyInstaller:

pip install pyinstaller
From the project folder, run:

pyinstaller --onefile --windowed email_reader.py
--onefile puts everything into a single .exe

--windowed prevents a console window from appearing (GUI only)

The executable will appear in the dist/ folder:

dist/email_reader.exe
You can move or share this .exe directly with other Windows users.

