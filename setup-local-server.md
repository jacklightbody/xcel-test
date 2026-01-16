# Setting Up Local HTTPS Server for Excel Add-in

Office.js add-ins **require HTTPS** even for local development. Here are several ways to host the files locally:

## Option 1: Using mkcert (Recommended - Creates Trusted Certificates)

`mkcert` creates locally-trusted certificates that work seamlessly with Office.js add-ins.

### Step 1: Install mkcert

**macOS:**
```bash
brew install mkcert
brew install nss  # for Firefox support (optional)
```

**Windows:**
```bash
# Using Chocolatey
choco install mkcert

# Or using Scoop
scoop bucket add extras
scoop install mkcert
```

**Linux:**
```bash
# Ubuntu/Debian
sudo apt install libnss3-tools
# Then download from https://github.com/FiloSottile/mkcert/releases
```

### Step 2: Install Local CA
```bash
mkcert -install
```

This installs a local Certificate Authority (CA) that your system will trust.

### Step 3: Generate Certificate for localhost
```bash
# Create certs directory
mkdir -p certs

# Generate certificate (creates both cert.pem and key.pem)
mkcert -key-file certs/key.pem -cert-file certs/cert.pem localhost 127.0.0.1 ::1
```

### Step 4: Start Server with http-server
```bash
# Install http-server if you haven't
npm install -g http-server

# Start server
http-server -p 3000 -S -C certs/cert.pem -K certs/key.pem
```

The certificate will be trusted by your system, so no browser warnings!

## Option 2: Using http-server with Self-Signed Certificate (Alternative)

If you can't use mkcert, you can use a self-signed certificate, but you'll need to trust it:

### Step 1: Install http-server
```bash
npm install -g http-server
```

### Step 2: Generate Self-Signed Certificate

Create a certificate for localhost:

```bash
# Create certs directory
mkdir -p certs
cd certs

# Generate private key
openssl genrsa -out key.pem 2048

# Generate certificate with proper extensions for localhost
openssl req -new -x509 -key key.pem -out cert.pem -days 365 \
  -subj "/CN=localhost" \
  -addext "subjectAltName=DNS:localhost,DNS:*.localhost,IP:127.0.0.1,IP:::1"
```

### Step 3: Trust the Certificate (Required for Office.js)

**macOS:**
```bash
# Add certificate to macOS keychain
sudo security add-trusted-cert -d -r trustRoot -k /Library/Keychains/System.keychain certs/cert.pem
```

**Windows:**
1. Double-click `certs/cert.pem`
2. Click "Install Certificate"
3. Choose "Local Machine" → "Place all certificates in the following store"
4. Browse and select "Trusted Root Certification Authorities"
5. Click Next → Finish

**Linux:**
```bash
# Copy certificate to system trust store
sudo cp certs/cert.pem /usr/local/share/ca-certificates/localhost.crt
sudo update-ca-certificates
```

### Step 4: Start the Server

From the project root directory:

```bash
http-server -p 3000 -S -C certs/cert.pem -K certs/key.pem
```

The server will run at `https://localhost:3000`

### Step 5: Update manifest.xml

If needed, update the URLs in `manifest.xml` to match your local server:
- Change `https://localhost:3000` to match your server URL and port

## Option 2: Using serve (Simple Alternative)

### Step 1: Install serve
```bash
npm install -g serve
```

### Step 2: Start with HTTPS (requires certificate)
```bash
serve -s . --ssl-cert certs/cert.pem --ssl-key certs/key.pem -l 3000
```

## Option 3: Using Python (Built-in, no installation needed)

### Step 1: Create a simple HTTPS server script

Create `start-server.py`:

```python
#!/usr/bin/env python3
import http.server
import ssl
import socketserver

PORT = 3000
DIRECTORY = "."

class MyHTTPRequestHandler(http.server.SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=DIRECTORY, **kwargs)

Handler = MyHTTPRequestHandler

with socketserver.TCPServer(("", PORT), Handler) as httpd:
    context = ssl.SSLContext(ssl.PROTOCOL_TLS_SERVER)
    context.load_cert_chain('certs/cert.pem', 'certs/key.pem')
    httpd.socket = context.wrap_socket(httpd.socket, server_side=True)
    print(f"Serving at https://localhost:{PORT}")
    httpd.serve_forever()
```

### Step 2: Generate certificate (same as Option 1, Step 2)

### Step 3: Run
```bash
python3 start-server.py
```

## Option 4: Using Node.js Express (More control)

### Step 1: Create package.json

Create `package.json` in project root:

```json
{
  "name": "excel-test-runner",
  "version": "1.0.0",
  "scripts": {
    "start": "node server.js"
  },
  "dependencies": {
    "express": "^4.18.0"
  }
}
```

### Step 2: Create server.js

```javascript
const express = require('express');
const https = require('https');
const fs = require('fs');
const path = require('path');

const app = express();
const PORT = 3000;

// Serve static files
app.use(express.static(__dirname));

// HTTPS options
const options = {
  key: fs.readFileSync('certs/key.pem'),
  cert: fs.readFileSync('certs/cert.pem')
};

https.createServer(options, app).listen(PORT, () => {
  console.log(`Server running at https://localhost:${PORT}`);
});
```

### Step 3: Install and run
```bash
npm install
npm start
```

## Troubleshooting Certificate Issues

### If you get "content isn't signed by a valid cert" error:

**Solution 1: Use mkcert (Recommended)**
- mkcert creates certificates that are automatically trusted by your system
- This is the easiest solution and works best with Office.js

**Solution 2: Trust the self-signed certificate in your OS**
- You must trust the certificate at the OS level, not just in the browser
- See Step 3 in Option 2 above for platform-specific instructions
- After trusting, restart your browser and Excel

**Solution 3: Check certificate validity**
- Make sure the certificate includes `subjectAltName` extensions
- Verify the certificate is for `localhost` or `127.0.0.1`
- Check that the certificate hasn't expired

### Browser Certificate Warnings

If you're still seeing browser warnings (even with mkcert):
1. **Chrome/Edge**: Click "Advanced" → "Proceed to localhost (unsafe)"
2. **Firefox**: Click "Advanced" → "Accept the Risk and Continue"
3. **Safari**: Click "Show Details" → "visit this website"

However, if the certificate is properly trusted, you shouldn't see these warnings.

## Sideloading the Add-in

Once your server is running:

1. Open Excel
2. Go to **Insert** → **Add-ins** → **My Add-ins** → **Upload My Add-in**
3. Select your `manifest.xml` file
4. The add-in should load from your local server

## Troubleshooting

### Certificate Errors
- Make sure the certificate files (`cert.pem`, `key.pem`) exist in the `certs/` directory
- Regenerate certificates if they're expired or corrupted

### Port Already in Use
- Change the port number in the server command (e.g., `-p 3001`)
- Update `manifest.xml` URLs to match

### CORS Issues
- Office.js handles CORS automatically for add-ins
- Make sure your server is actually running and accessible

### Manifest URL Issues
- Verify the URLs in `manifest.xml` match your server URL exactly
- Check that `taskpane.html` and other files are accessible in the browser

## Quick Start Script

You can create a simple script to start everything:

**start.sh** (Mac/Linux):
```bash
#!/bin/bash
if [ ! -f "certs/cert.pem" ]; then
    echo "Generating certificates..."
    mkdir -p certs
    openssl genrsa -out certs/key.pem 2048
    openssl req -new -x509 -key certs/key.pem -out certs/cert.pem -days 365 -subj "/CN=localhost"
fi
echo "Starting server at https://localhost:3000"
http-server -p 3000 -S -C certs/cert.pem -K certs/key.pem
```

Make it executable: `chmod +x start.sh`

Then run: `./start.sh`
