# PDF Crawler

A modern, functional Python app to scan folders for PDFs, match them against an Excel/CSV list, and copy the most recent versions to a destination folder.

---

## 📋 Prerequisites

- **Python 3.8+** installed on your system
- **Git** (optional, for cloning the project)

---

## 🚀 Quick Start

### 1. Create a Virtual Environment

Navigate to the project folder and create a virtual environment:

```bash
cd path\to\Crawler
python -m venv venv
```

### 2. Activate the Virtual Environment

**On Windows (PowerShell):**
```powershell
.\venv\Scripts\Activate.ps1
```

**On Windows (Command Prompt):**
```cmd
venv\Scripts\activate.bat
```

**On macOS/Linux:**
```bash
source venv/bin/activate
```

You should see `(venv)` appear in your terminal prompt.

### 3. Install Dependencies

```bash
pip install -r requirements.txt
```

This installs:
- `pandas` — for reading Excel/CSV files
- `openpyxl` — for Excel file support
- `tkinterdnd2` — for drag-and-drop functionality

### 4. Run the App

```bash
python app.py
```

The GUI window will open. You're ready to go!

---

## 📖 How to Use

### Step 1: Select Source Folder
Click **Browse** next to "Source Folder" and select the folder containing your PDF files. This will be scanned recursively (all subfolders included).

### Step 2: Scan PDFs
Click **Scan PDFs →** to scan the folder.
- A timestamped cache file (`scan_YYYY-MM-DD_HH-MM-SS.json`) will be saved in the `cache/` folder.
- The cache status bar updates to show the active cache.

### Step 3: Load Excel / CSV
You have two options:
- **Drag & Drop**: Drag your Excel or CSV file directly onto the drop zone.
- **Browse**: Click the **Browse** button to select the file.

### Step 4: Configure Columns
- **Column name**: Enter the name of the column containing file names (default: `FileName`).
- **Start row index**: Enter the 0-based row index to start reading from (default: `0`).

### Step 5: Set Destination Folder
Click **Browse** next to "Destination Folder" and select where the matched PDFs should be copied.

### Step 6: Run
Click **⚡ Run — Match & Copy PDFs** to execute:
1. Read file names from your Excel/CSV.
2. Match them against the cache.
3. Copy the most recent version of each matched PDF to the destination.

---

## 💾 Cache Management

### What is the Cache?
The cache is a JSON file that maps PDF file names to their full paths. It's created when you click **Scan PDFs**.

### Why Use It?
- **Speed**: Reuse a scan from earlier without rescanning the entire folder.
- **History**: All scans are timestamped and stored in the `cache/` folder.

### Load a Previous Cache
1. Click **Load Cache** (below the Source Folder section).
2. Select any `.json` file from the `cache/` folder.
3. The cache status bar updates to show which file is in use.

### Auto-Scan
If you run the app without a loaded cache, it automatically scans the source folder and creates a new cache.

---

## 📁 Project Structure

```
Crawler/
├── app.py                          # Main application
├── requirements.txt                # Python dependencies
├── README.md                       # This file
└── cache/                          # Timestamped cache files (auto-created)
    ├── scan_2026-03-25_09-34-54.json
    ├── scan_2026-03-25_10-15-22.json
    └── ...
```

---

## 🎨 Features

- **Modern Dark UI** — Clean, professional interface with Segoe UI font.
- **Drag & Drop** — Drop Excel/CSV files directly onto the app.
- **Timestamped Caches** — Never lose a scan; all are preserved with timestamps.
- **Cache Management** — Load any previous scan without rescanning.
- **Thread-Safe Logging** — Real-time feedback with colour-coded messages.
- **Functional Design** — Pure functions for max testability and simplicity.
- **Smart Matching** — Handles file name variations (with/without `.pdf` extension).
- **Most Recent Priority** — If a PDF exists in multiple locations, copies the newest.

---

## ⚙️ Deactivate Virtual Environment

When you're done, deactivate the virtual environment:

```bash
deactivate
```

---

## 🐛 Troubleshooting

### `ModuleNotFoundError: No module named 'tkinter'`
Tkinter comes with Python by default. If missing, install it:
- **Windows**: Usually included; re-run Python installer and enable Tcl/Tk.
- **Linux**: `sudo apt install python3-tk`
- **macOS**: `brew install python-tk`

### `ModuleNotFoundError: No module named 'pandas'`
Make sure the virtual environment is activated and dependencies are installed:
```bash
pip install -r requirements.txt
```

### Drag-and-drop doesn't work
`tkinterdnd2` has platform-specific quirks. The **Browse** button always works as a fallback.

### Cache file corruption
Simply delete the `.json` file and rescan. Each scan creates a new file.

---

## 📝 Example Workflow

1. **Setup**: Create venv, activate, install dependencies.
2. **Configure**: Select your PDF folder (e.g., `D:\PDFs\`).
3. **Scan**: Click **Scan PDFs →** (creates `cache/scan_2026-03-25_09-34-54.json`).
4. **Prepare Excel**: Create a spreadsheet with a column named `FileName` listing PDFs you want.
5. **Drop Excel**: Drag your file onto the app.
6. **Run**: Select destination folder and click **⚡ Run**.
7. **Results**: Matched PDFs appear in your destination folder.

---

## 📄 License

This project is open source. Use and modify freely.

---

## 💬 Support

For issues or feature requests, check the log panel in the app for detailed error messages and timestamps.
