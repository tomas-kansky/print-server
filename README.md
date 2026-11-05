# Web Print Server

Lightweight, single-page self-hosted web printing service for local home networks. Designed for seamless printing from Mobile Safari on iOS devices to local CUPS printers.

---

## Features

- **Mobile Safari Optimized:** iOS Human Interface Guidelines look-and-feel with safe-area support, dark mode, responsive touch feedback, and Web App Manifest for Add to Home Screen.
- **Supported Formats:** PDF, PNG, JPG, and JPEG files up to 25 MB.
- **CUPS Integration:** Dispatches print jobs directly via `lp` with automatic `fit-to-page` and configurable copy counts.
- **Printer Auto-Detection:** Automatically discovers available CUPS printers and the default destination, with support for environment-based overrides.
- **Secure File Handling:** Unique UUID temporary file generation in `/tmp/print_jobs/` with guaranteed cleanup in `finally` blocks.

---

## Quick Start: Native Linux Host (Systemd)

### 1. Prerequisites
Ensure CUPS and CUPS client utilities are installed on the Linux host:
```bash
sudo apt update
sudo apt install -y python3 python3-pip python3-venv cups cups-client
```

Verify CUPS status and printers:
```bash
lpstat -p -d
```

### 2. Installation
Clone or copy this repository to `/opt/print-server`:
```bash
sudo mkdir -p /opt/print-server
sudo chown -R $USER:$USER /opt/print-server
cp -r . /opt/print-server/
cd /opt/print-server
```

Create a virtual environment and install dependencies:
```bash
python3 -m venv venv
./venv/bin/pip install --upgrade pip
./venv/bin/pip install -r requirements.txt
```

### 3. Configure and Start Systemd Service
Copy the service file and enable it:
```bash
sudo cp webprint.service /etc/systemd/system/
sudo systemctl daemon-reload
sudo systemctl enable --now webprint
```

Verify that the service is running:
```bash
sudo systemctl status webprint
curl http://localhost:8000/health
```

---

## Quick Start: Docker Compose

If running via Docker on a host with an existing CUPS server:

### 1. Build and Run Container
```bash
docker compose up -d --build
```

### 2. Check Logs
```bash
docker compose logs -f
```

---

## Configuration

Environment variables can be set in the systemd unit or `docker-compose.yml`:

| Variable | Default | Description |
|---|---|---|
| `PORT` | `8000` | Port to bind the HTTP server |
| `PRINTER_NAME` | `""` (auto-detect) | Explicit CUPS printer queue name |
| `TEMP_DIR` | `/tmp/print_jobs` | Temporary storage for incoming jobs |
| `CUPS_SERVER` | `/var/run/cups/cups.sock` | CUPS socket or server hostname |

---

## API Endpoints

- `GET /`: Mobile-friendly iOS web application
- `GET /health`: Health check returning `{"status": "ok"}`
- `GET /manifest.json`: Web App Manifest for iOS PWA installation
- `GET /api/printers`: List of discovered CUPS printers and default queue
- `POST /api/print`: Multipart form-data upload (`file`, optional `printer`, optional `copies`)

---

## Verification Steps

1. **Check Health Check Endpoint:**
   ```bash
   curl http://<SERVER_IP>:8000/health
   ```
2. **List Available Printers:**
   ```bash
   curl http://<SERVER_IP>:8000/api/printers
   ```
3. **Open in Mobile Safari:**
   Navigate to `http://<SERVER_IP>:8000/` on your iPhone.
4. **Print a Document:**
   Select a document or photo, adjust the copy counter if needed, and tap **Print Document**.
5. **Add to Home Screen:**
   Tap the iOS Share button and choose **Add to Home Screen** to launch in standalone fullscreen mode.

---

## 🤖 Agentic Design Framework

Před zahájením práce a pročtením projektu načti kontext a pravidla z centrálního repozitáře [Agentic Design Framework](https://github.com/tomas-kansky/agentic-design-framework) přes Git v konzoli.
