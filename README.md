# Excel Unit Test Runner

An Office.js Excel add-in that allows you to write and run unit tests for Excel workbooks to validate the correctness of a set of formulas you've created.

## Usage
![Demo](https://github.com/jacklightbody/xcel-testraw/master/documentation/example-test-run.gif)

## Test Format

Test files are JSON files that contain an array of test sets


### Example

```json
[
  {
    "name": "Base case revenue",
    "inputs": {
      "Assumptions!B2": 0.05,
      "Assumptions!B3": 100000
    },
    "assertions": [
      {
        "cell": "Outputs!E12",
        "equals": 1234567,
        "tolerance": 1
      }
    ]
  },
  {
    "name": "High growth scenario",
    "inputs": {
      "Assumptions!B2": 0.10,
      "Assumptions!B3": 200000
    },
    "assertions": [
      {
        "cell": "Outputs!E12",
        "equals": 2469134,
        "tolerance": 1
      }
    ]
  }
]
```

## How It Works

For each test, the add-in performs the following steps:

1. **Snapshot State**: Captures current values and formulas for all cells referenced in inputs and assertions
2. **Apply Inputs**: Sets the input values as specified in the test
3. **Force Calculation**: Triggers Excel's full calculation to ensure all dependent formulas recalculate
4. **Read Outputs**: Retrieves the actual calculated values from assertion cells
5. **Evaluate Assertions**: Compares actual vs expected values (with tolerance for numeric comparisons)
6. **Restore State**: Restores all original values and formulas, ensuring the workbook is unchanged

This means the unit tests both **preserves state** and **exactly match** the native excel behavior.
## Setup

### Quick Setup (Recommended)

Run `setup.sh` to automatically install dependencies, generate trusted certificates, and prepare everything:
```bash
./setup.sh && ./start.sh
```

This script will:
- Install mkcert (if needed) for trusted certificates
- Generate trusted HTTPS certificates
- **Auto-install the manifest for Mac Excel users**
- **Start the server** (via `./start.sh`)

After the first initialization, you can call `./start.sh` to start the server.

### Manual Setup (Fallback)

If the automated setup fails, follow these manual steps:

1. **Install dependencies**:
   ```bash
   # Install mkcert for trusted certificates:
   # macOS: brew install mkcert
   # Windows: choco install mkcert
   # Linux: sudo apt-get install libnss3-tools (then download mkcert)
   ```

2. **Generate certificates**:
   ```bash
   mkcert -install
   mkdir -p certs
   mkcert -key-file certs/key.pem -cert-file certs/cert.pem localhost 127.0.0.1 ::1
   ```

3. **Start the server**:
   ```bash
   python3 -c "
import http.server
import ssl
import socketserver

PORT = 3000
DIRECTORY = '.'

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
   ```

4. **Launch Excel**:
   - **Mac users**: The manifest is auto-installed! Just go to **Inert** → **My Add-ins** and select "Excel Unit Test Runner"
   - **Other users**: Go to **Insert** → **Add-ins** → **My Add-ins** → **Upload My Add-in** and select `manifest.xml`

![where to find the add-in](https://github.com/jacklightbody/xcel-testraw/master/documentation/add-add-in.png)


The `manifest.xml` is already configured for `https://localhost:3000`.

**Important**: Office.js add-ins require HTTPS, even for local development. The setup script handles this automatically using mkcert for system-trusted certificates.

## Usage

1. Open the Excel workbook you want to test
2. Open the add-in task pane (via the ribbon button or Insert > My Add-ins)
3. Choose your input method:
   - **Paste JSON**: Copy and paste JSON test content directly
   - **Load File**: Select a JSON test file from your computer
4. Review the test inputs and assertions that are displayed
5. Click "Run Test" to execute
6. View the results showing which assertions passed or failed

### Loading Test Files

You can load test files in two ways:

**Method 1: Paste JSON**
- Copy your JSON test content to the clipboard
- Paste it directly into the textarea
- Press `Ctrl+Enter` (or `Cmd+Enter` on Mac) to run the test

**Method 2: Load File** (Recommended for larger test files)
- Click the "Load File" tab
- Click "Choose File" and select your `.json` test file
- The filename will be displayed once loaded
- Click "Run Test" to execute

## File Structure

```
/
├── manifest.xml              # Office.js add-in manifest
├── taskpane/
│   ├── taskpane.html        # Task pane UI
│   ├── taskpane.js          # UI logic and test execution
│   └── taskpane.css         # Styling
├── scripts/
│   └── test-runner.js       # Core test execution logic (reference implementation)
├── tests/
│   └── sample-test.json     # Example test file
└── README.md                # This file
```

## Safety Guarantees

- **State Snapshot**: All cell values and formulas are captured before test execution
- **Atomic Restore**: State restoration happens in a `finally` block, ensuring it executes even if assertions fail
- **Formula Preservation**: Original formulas are restored if they were overwritten by input values
- **Error Handling**: Restore operations are wrapped in error handling to prevent data loss

## Limitations

- Tests should only modify cells specified in the inputs - other cells are snapshotted but restoring them may affect unrelated workbook state
- Large workbooks with extensive calculations may take time to snapshot/restore
- The add-in requires a web server to function (cannot run from `file://` protocol)

## Troubleshooting

- **"Failed to access worksheet"**: Ensure worksheet names match exactly (case-sensitive)
- **"Invalid cell address format"**: Cell addresses must be in format "SheetName!A1"
- **Calculation not updating**: The add-in waits 100ms after forcing calculation; complex models may need more time
- **State not restoring**: Check browser console for restore errors; formulas may need to be restored before values

## Todo

- Allow for relative references (i.e. 2 cells to the right of "total income" on sheet 3)
- UI to help create tests
- Bundle and deploy to msft so installation is easy
- Guard mode to retrigger on save automatically
  - Also keyboard shortcuts
- Locking. Prevent (or at least detect) user edits while tests are running

