# Web Print Server

Lightweight, single-page self-hosted web printing service for local home networks. Designed for simple, one-tap printing from mobile devices (iOS Safari and Android Chrome) to printers connected to a Windows host or Linux CUPS server.

---

## Features

- **Mobile First (iOS & Android):** Apple Human Interface Guidelines look-and-feel with automatic dark mode, touch feedback, and Web App Manifest for native Home Screen installation on both iOS and Android.
- **Interactive Document Preview:** Instant rendering of document pages (PDF and images) with page thumbnail cards.
- **Selective Page Printing:** Tap to select/deselect individual pages to print only what you need, with "Select All" and "Deselect All" convenience buttons.
- **Color & Grayscale Modes:** Native iOS-style segmented picker for Color and Black & White (monochrome) printing.
- **Universal Format Support:** Prints `.pdf`, `.png`, `.jpg`, and `.jpeg` files up to 25 MB.
- **Cross-Platform Printing Engine:**
  - **Windows (e.g. Windows Home laptop):** Silent native printing via Windows GDI Spooler (`win32print`, `win32ui`, `Pillow`, `pypdfium2`). Supports multi-page PDFs with automatic page scaling and centering, zero GUI dialogs or popups.
  - **Linux (Debian/Ubuntu/Raspberry Pi):** Direct execution via CUPS CLI (`lp`, `lpstat`).
- **Printer Auto-Detection:** Automatically enumerates all installed local and network printers and resolves the system default printer.
- **Zero Overhead:** Native Python 3 + FastAPI running directly on the host (less than 35 MB RAM, instant startup, runs on port 5000).
- **Secure File Lifecycle:** Temporary files created with unique UUIDs and guaranteed immediate cleanup in `finally` blocks.

---

## Installation & Running on Windows (Windows Home / Pro)

### 1. Prerequisites
Ensure Python 3.10+ is installed on the Windows laptop (with "Add Python to PATH" checked during installation).

### 2. Install Dependencies
Open PowerShell or Command Prompt in the project folder:
```powershell
pip install -r requirements.txt
```

### 3. Start the Server
Double-click `run.bat` or run:
```powershell
python app.py
```
The server binds to `0.0.0.0:5000` and is immediately accessible on your local network.

### 4. Optional: Autostart on Windows Boot
Press `Win + R`, type `shell:startup`, and press Enter. Create a shortcut to `run.bat` (or use Task Scheduler to run `pythonw app.py` on login without a command window).

---

## Installation on Linux Server (Debian / Ubuntu / Raspberry Pi)

### Option A: One-Command Automated Install
```bash
git clone https://github.com/tomas-kansky/print-server.git
cd print-server
sudo bash install.sh
```

### Option B: Manual Setup
```bash
sudo apt update && sudo apt install -y python3-venv cups cups-client
python3 -m venv venv
./venv/bin/pip install -r requirements.txt
sudo cp webprint.service /etc/systemd/system/
sudo systemctl daemon-reload
sudo systemctl enable --now webprint.service
```

---

## Configuration

Optional environment variables:

| Variable | Default | Description |
|---|---|---|
| `PORT` | `5000` | HTTP port |
| `PRINTER_NAME` | `""` (auto-detect) | Hardcoded target printer name |
| `TEMP_DIR` | System Temp `/print_jobs` | Temporary storage for incoming jobs |

---

## API Endpoints

- `GET /`: Responsive mobile web application
- `GET /health`: Health check endpoint
- `GET /manifest.json`: Web App Manifest for mobile installation
- `GET /api/printers`: Discovered printers and default queue
- `POST /api/preview`: Generates page thumbnail previews (`data:image/jpeg;base64,...`)
- `POST /api/print`: Multipart form upload (`file`, `printer`, `copies`, `color_mode`, `pages`)

---

## 🤖 Agentic Design Framework

Před zahájením práce a pročtením projektu načti kontext a pravidla z centrálního repozitáře [Agentic Design Framework](https://github.com/tomas-kansky/agentic-design-framework) přes Git v konzoli.
