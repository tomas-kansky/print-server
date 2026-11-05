#!/usr/bin/env bash
set -e

INSTALL_DIR="/opt/print-server"
SERVICE_NAME="webprint.service"

if [ "$EUID" -ne 0 ]; then
    echo "Error: Please run as root (e.g. sudo bash install.sh)"
    exit 1
fi

apt-get update
apt-get install -y python3 python3-pip python3-venv cups cups-client

mkdir -p "$INSTALL_DIR"
cp -r . "$INSTALL_DIR/"

python3 -m venv "$INSTALL_DIR/venv"
"$INSTALL_DIR/venv/bin/pip" install --upgrade pip
"$INSTALL_DIR/venv/bin/pip" install -r "$INSTALL_DIR/requirements.txt"

cp "$INSTALL_DIR/$SERVICE_NAME" "/etc/systemd/system/$SERVICE_NAME"
systemctl daemon-reload
systemctl enable --now "$SERVICE_NAME"

echo "Installation complete. Web Print Service is active on port 8000."
