#!/bin/bash
# Check if certificates exist
if [ ! -f "certs/cert.pem" ] || [ ! -f "certs/key.pem" ]; then
    echo "❌ Certificates not found. Please run ./setup.sh first."
    exit 1
fi

# Check if Python3 is available
if ! command -v python3 &> /dev/null; then
    echo "❌ Python3 not found. Please install Python 3."
    exit 1
fi

echo "🚀 Starting Excel Test Runner server..."
echo "Server will run at: https://localhost:3000"
echo ""
echo "To sideload in Excel:"
echo "  Insert > Add-ins > My Add-ins > Upload My Add-in"
echo "  Select manifest.xml"
echo ""
echo "Press Ctrl+C to stop the server"
echo ""

python3 -c "
import http.server
import ssl
import socketserver
import os

PORT = 3000
DIRECTORY = os.getcwd()

class MyHTTPRequestHandler(http.server.SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=DIRECTORY, **kwargs)

Handler = MyHTTPRequestHandler

with socketserver.TCPServer(('', PORT), Handler) as httpd:
    context = ssl.SSLContext(ssl.PROTOCOL_TLS_SERVER)
    context.load_cert_chain('certs/cert.pem', 'certs/key.pem')
    httpd.socket = context.wrap_socket(httpd.socket, server_side=True)
    print(f'Serving at https://localhost:{PORT}')
    httpd.serve_forever()
"
