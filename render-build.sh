#!/usr/bin/env bash

# Install Python dependencies
pip install -r requirements.txt

# Install LibreOffice for PDF conversion
apt-get update
apt-get install -y libreoffice --no-install-recommends

# Clean up
apt-get clean
rm -rf /var/lib/apt/lists/*
