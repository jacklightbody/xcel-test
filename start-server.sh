#!/bin/bash

# Excel Test Runner - Local HTTPS Server Startup Script

echo "Excel Unit Test Runner - Local Server Setup"
echo "============================================"
echo ""

# Check if http-server is installed
if ! command -v http-server &> /dev/null; then
    echo "http-server is not installed."
    echo "Installing http-server..."
    npm install -g http-server
    if [ $? -ne 0 ]; then
        echo "Error: Failed to install http-server"
        echo "Please install manually: npm install -g http-server"
        exit 1
    fi
fi

# Create certs directory if it doesn't exist
mkdir -p certs

# Check if mkcert is available (preferred method)
if command -v mkcert &> /dev/null; then
    echo "Using mkcert for certificate generation (recommended)..."
    
    # Check if local CA is installed
    if [ ! -d "$(mkcert -CAROOT)" ]; then
        echo "Installing mkcert local CA..."
        mkcert -install
    fi
    
    # Generate certificate if it doesn't exist
    if [ ! -f "certs/cert.pem" ] || [ ! -f "certs/key.pem" ]; then
        echo "Generating trusted certificate with mkcert..."
        mkcert -key-file certs/key.pem -cert-file certs/cert.pem localhost 127.0.0.1 ::1
        echo "Certificate generated successfully!"
    else
        echo "Certificate already exists."
    fi
else
    echo "mkcert not found. Using OpenSSL (self-signed certificate)..."
    echo "Note: For Office.js add-ins, mkcert is recommended for trusted certificates."
    echo "Install with: brew install mkcert (macOS) or see setup-local-server.md"
    echo ""
    
    # Generate self-signed certificate if it doesn't exist
    if [ ! -f "certs/cert.pem" ] || [ ! -f "certs/key.pem" ]; then
        echo "Generating self-signed certificate..."
        openssl genrsa -out certs/key.pem 2048
        openssl req -new -x509 -key certs/key.pem -out certs/cert.pem -days 365 \
          -subj "/CN=localhost" \
          -addext "subjectAltName=DNS:localhost,DNS:*.localhost,IP:127.0.0.1,IP:::1"
        echo "Certificate generated!"
        echo ""
        echo "⚠️  WARNING: Self-signed certificates may not work with Office.js."
        echo "You may need to trust the certificate in your OS:"
        echo "  macOS: sudo security add-trusted-cert -d -r trustRoot -k /Library/Keychains/System.keychain certs/cert.pem"
        echo "  Or install mkcert for automatically trusted certificates."
    fi
fi
echo ""

echo "Starting HTTPS server at https://localhost:3000"
echo ""
echo "Note: Your browser may show a security warning about the self-signed certificate."
echo "This is normal for local development. Click 'Advanced' and proceed."
echo ""
echo "To sideload the add-in in Excel:"
echo "  1. Insert > Add-ins > My Add-ins > Upload My Add-in"
echo "  2. Select manifest.xml from this directory"
echo ""
echo "Press Ctrl+C to stop the server"
echo ""

# Start the server
http-server -p 3000 -S -C certs/cert.pem -K certs/key.pem
