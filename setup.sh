#!/usr/bin/env bash
set -euo pipefail

echo "Installing Python dependencies..."
python3 -m pip install -r requirements.txt

if command -v apt-get >/dev/null 2>&1; then
  echo "Installing LibreOffice via apt-get..."
  sudo apt-get update
  sudo apt-get install -y libreoffice-core libreoffice-writer
elif command -v brew >/dev/null 2>&1; then
  echo "Installing LibreOffice via Homebrew..."
  brew install --cask libreoffice
else
  echo "Could not detect apt-get or brew."
  echo "Please install LibreOffice manually and make sure 'soffice' is on PATH."
  exit 1
fi

echo "Setup complete."
