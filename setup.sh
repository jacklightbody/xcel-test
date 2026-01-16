#!/bin/bash

# Excel Test Runner - One-Command Setup Script
# This script sets up HTTPS server, generates certificates, and prepares for sideloading

set -e  # Exit on any error

echo "🚀 Excel Unit Test Runner - One-Command Setup"
echo "=============================================="
echo ""

# Detect OS
OS="unknown"
if [[ "$OSTYPE" == "darwin"* ]]; then
    OS="macos"
elif [[ "$OSTYPE" == "linux-gnu"* ]]; then
    OS="linux"
elif [[ "$OSTYPE" == "msys" ]] || [[ "$OSTYPE" == "cygwin" ]]; then
    OS="windows"
fi

echo "Detected OS: $OS"

# Function to check if command exists
command_exists() {
    command -v "$1" >/dev/null 2>&1
}

# Install dependencies based on OS
install_dependencies() {
    echo "📦 Installing dependencies..."
    
    if [ "$OS" = "macos" ]; then
        # Check for Homebrew
        if ! command_exists brew; then
            echo "❌ Homebrew not found. Please install Homebrew first:"
            echo "   /bin/bash -c \"\$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)\""
            exit 1
        fi
        
        # Install mkcert if not present (Python3 is built into macOS)
        if ! command_exists mkcert; then
            echo "Installing mkcert (for trusted certificates)..."
            brew install mkcert
            brew install nss  # for Firefox support
        fi
        
    elif [ "$OS" = "linux" ]; then
        # Install mkcert if not present (Python3 is usually available)
        if ! command_exists mkcert; then
            echo "Installing mkcert (for trusted certificates)..."
            sudo apt-get update
            sudo apt-get install -y libnss3-tools
            echo "Please download mkcert from https://github.com/FiloSottile/mkcert/releases"
            echo "and place it in your PATH, then run this script again."
            exit 1
        fi
        
    elif [ "$OS" = "windows" ]; then
        echo "Windows detected. Please ensure you have:"
        echo "  1. Python 3 installed: https://www.python.org/downloads/"
        echo "  2. Chocolatey installed: https://chocolatey.org/install"
        echo ""
        echo "Then run these commands manually:"
        echo "  choco install mkcert"
        echo ""
        echo "After installing, run this script again."
        if ! command_exists mkcert; then
            exit 1
        fi
    fi
}

# Setup certificates
setup_certificates() {
    echo "🔐 Setting up HTTPS certificates..."
    mkdir -p certs
    
    if command_exists mkcert; then
        echo "Using mkcert for trusted certificates..."
        
        # Install local CA if not already done
        if [ ! -d "$(mkcert -CAROOT 2>/dev/null)" ]; then
            echo "Installing mkcert local CA..."
            mkcert -install
        fi
        
        # Generate certificate if needed
        if [ ! -f "certs/cert.pem" ] || [ ! -f "certs/key.pem" ]; then
            echo "Generating trusted certificate..."
            mkcert -key-file certs/key.pem -cert-file certs/cert.pem localhost 127.0.0.1 ::1
            echo "✅ Trusted certificate generated!"
        else
            echo "✅ Certificate already exists."
        fi
        
    else
        echo "⚠️  mkcert not available. Using self-signed certificate..."
        if [ ! -f "certs/cert.pem" ] || [ ! -f "certs/key.pem" ]; then
            openssl genrsa -out certs/key.pem 2048
            openssl req -new -x509 -key certs/key.pem -out certs/cert.pem -days 365 \
              -subj "/CN=localhost" \
              -addext "subjectAltName=DNS:localhost,DNS:*.localhost,IP:127.0.0.1,IP:::1"
            echo "✅ Self-signed certificate generated!"
            echo ""
            echo "⚠️  WARNING: You may need to trust this certificate:"
            if [ "$OS" = "macos" ]; then
                echo "   sudo security add-trusted-cert -d -r trustRoot -k /Library/Keychains/System.keychain certs/cert.pem"
            elif [ "$OS" = "linux" ]; then
                echo "   sudo cp certs/cert.pem /usr/local/share/ca-certificates/localhost.crt"
                echo "   sudo update-ca-certificates"
            fi
        else
            echo "✅ Certificate already exists."
        fi
    fi
}

# Install manifest for Mac Excel
install_manifest_mac() {
    echo "📋 Installing manifest for Excel on Mac..."
    
    # Get the current user
    USER=$(whoami)
    EXCEL_WEF_DIR="/Users/$USER/Library/Containers/com.microsoft.Excel/Data/Documents/wef"
    
    # Create the directory if it doesn't exist
    if [ ! -d "$EXCEL_WEF_DIR" ]; then
        echo "Creating Excel wef directory: $EXCEL_WEF_DIR"
        mkdir -p "$EXCEL_WEF_DIR"
    fi
    
    # Copy the manifest
    if [ -f "manifest.xml" ]; then
        cp manifest.xml "$EXCEL_WEF_DIR/"
        echo "✅ Manifest copied to $EXCEL_WEF_DIR"
        echo ""
        echo "🎯 The add-in will now appear automatically in Excel!"
        echo "   Just open Excel and look for it in Home > Add-ins"
    else
        echo "❌ manifest.xml not found in current directory"
        return 1
    fi
}

# Create start script for future use
create_start_script() {
    cat > start.sh << 'EOF'
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
EOF
    chmod +x start.sh
}



# Main execution
main() {
    # Install dependencies
    install_dependencies
    
    # Setup certificates
    setup_certificates
    
    # Create quick start script
    create_start_script
    
    echo ""
    echo "🎉 Setup completed successfully!"
    echo ""
    
    # Auto-install manifest on Mac
    if [ "$OS" = "macos" ]; then
        install_manifest_mac
        echo ""
        echo "📋 Setup complete! To start the server:"
        echo "   • Run: ./start.sh"
        echo "📯 Open Excel and find the add-in:"
        echo "   • Home > Add-ins > Excel Unit Test Runner"
        echo "   • The manifest is already installed!"
        echo ""
        echo "🎯 The add-in will be available at: https://localhost:3000"
        echo ""
        echo "📄 For detailed documentation, see README.md"
        
    else
        echo "📋 Setup complete! To start the server:"
        echo "   • Run: ./start.sh"
        echo ""
        echo "📯 Sideload the add-in in Excel:"
        echo "   • Open Excel"
        echo "   • Go to Insert > Add-ins > My Add-ins > Upload My Add-in"
        echo "   • Select 'manifest.xml' from this directory"
        echo ""
        echo "🎯 The add-in will be available at: https://localhost:3000"
        echo ""
        echo "📄 For detailed documentation, see README.md"
    fi
}

# Run main function
main "$@"