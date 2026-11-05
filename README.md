# Web Print Server

Lightweight, single-page self-hosted web printing service for local home networks. Designed for simple, one-tap printing from mobile devices (iOS Safari and Android Chrome) to local CUPS printers on Linux.

---

## Features

- **Mobile First (iOS & Android):** Apple Human Interface Guidelines look-and-feel with automatic dark mode, touch feedback, and Web App Manifest for native Home Screen installation on both iOS and Android.
- **Universal Format Support:** Prints `.pdf`, `.png`, `.jpg`, and `.jpeg` files up to 25 MB.
- **Direct CUPS Integration:** Lightweight execution via `lp` with `fit-to-page` and copy control.
- **Auto-Detection:** Automatically discovers printers and default destinations via `lpstat`.
- **Zero Overhead:** Native Python 3 + FastAPI running directly under `systemd` (less than 30 MB RAM, instant startup, no Docker bloat).
- **Secure File Lifecycle:** Temporary files created with unique UUIDs in `/tmp/print_jobs/` and guaranteed immediate cleanup.

---

## Installation on Server (Linux Host)

### Option A: One-Command Automated Install
On the Linux server (Debian, Ubuntu, Raspberry Pi OS):
```bash
git clone https://github.com/tomas-kansky/print-server.git
cd print-server
sudo bash install.sh
```

### Option B: Manual Step-by-Step Install

1. **Install System Dependencies:**
   ```bash
   sudo apt update
   sudo apt install -y python3 python3-pip python3-venv cups cups-client
   ```

2. **Deploy Application Directory:**
   ```bash
   sudo mkdir -p /opt/print-server
   sudo cp -r . /opt/print-server/
   cd /opt/print-server
   ```

3. **Setup Virtual Environment:**
   ```bash
   sudo python3 -m venv venv
   sudo ./venv/bin/pip install --upgrade pip
   sudo ./venv/bin/pip install -r requirements.txt
   ```

4. **Enable Systemd Service:**
   ```bash
   sudo cp webprint.service /etc/systemd/system/
   sudo systemctl daemon-reload
   sudo systemctl enable --now webprint.service
   ```

5. **Verify Service Status:**
   ```bash
   sudo systemctl status webprint.service
   curl http://localhost:8000/health
   ```

---

## Android & iOS Compatibility

- **iOS (Mobile Safari):** Supports camera roll, Files app picker, and PWA mode via `Share -> Add to Home Screen`.
- **Android (Chrome / Firefox / Edge):** Fully compatible. The file selector natively opens Android's system document and media picker. Chrome provides an install banner to add the app directly to the launcher as a standalone PWA.

---

## Configuration

Optional environment variables in `/etc/systemd/system/webprint.service`:

| Variable | Default | Description |
|---|---|---|
| `PORT` | `8000` | HTTP port |
| `PRINTER_NAME` | `""` (auto-detect) | Hardcoded target CUPS printer queue |
| `TEMP_DIR` | `/tmp/print_jobs` | Temporary storage for incoming jobs |

---

## API Endpoints

- `GET /`: Responsive mobile web application
- `GET /health`: Health check endpoint
- `GET /manifest.json`: Web App Manifest for mobile installation
- `GET /api/printers`: Discovered CUPS printers and default queue
- `POST /api/print`: Multipart form upload (`file`, `printer`, `copies`)

---

## 🤖 Agentic Design Framework

Před zahájením práce a pročtením projektu načti kontext a pravidla z centrálního repozitáře [Agentic Design Framework](https://github.com/tomas-kansky/agentic-design-framework) přes Git v konzoli.
