# Excel Unit Test Runner

An Office.js Excel add-in that allows you to write and run unit tests for Excel workbooks. The add-in ensures correctness by using Excel's native calculation engine and guarantees safety by automatically snapshotting and restoring workbook state.

## Features

- **State Preservation**: Automatically snapshots workbook state before each test and restores it afterward
- **Native Excel Calculations**: Uses Excel's built-in calculation engine for accurate results
- **Flexible Assertions**: Supports numeric comparisons with configurable tolerance
- **JSON-Based Tests**: Simple JSON format for defining test cases

## Test Format

Test files are JSON files that can contain either:
- A **single test object** (for backward compatibility)
- An **array of test objects** (for multiple tests in one file)

### Single Test Format

```json
{
  "name": "Test name",
  "inputs": {
    "SheetName!CellAddress": value,
    ...
  },
  "assertions": [
    {
      "cell": "SheetName!CellAddress",
      "equals": expectedValue,
      "tolerance": optionalTolerance
    },
    ...
  ]
}
```

### Multiple Tests Format

```json
[
  {
    "name": "Test name 1",
    "inputs": { ... },
    "assertions": [ ... ]
  },
  {
    "name": "Test name 2",
    "inputs": { ... },
    "assertions": [ ... ]
  }
]
```

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

When a test file contains multiple tests, they will be executed sequentially and all results will be displayed together.

## How It Works

For each test, the add-in performs the following steps:

1. **Snapshot State**: Captures current values and formulas for all cells referenced in inputs and assertions
2. **Apply Inputs**: Sets the input values as specified in the test
3. **Force Calculation**: Triggers Excel's full calculation to ensure all dependent formulas recalculate
4. **Read Outputs**: Retrieves the actual calculated values from assertion cells
5. **Evaluate Assertions**: Compares actual vs expected values (with tolerance for numeric comparisons)
6. **Restore State**: Restores all original values and formulas, ensuring the workbook is unchanged

## Setup

### Prerequisites

- Excel for Windows, Excel for Mac, or Excel Online
- A web server to host the add-in files (for local development)

### Installation

1. **Host the files**: The add-in needs to be served over HTTPS. For local development, you can use:
   - [Office Add-in CLI](https://github.com/OfficeDev/Office-Addin-TaskPane-SSO)
   - A local web server with HTTPS (e.g., using `http-server` with SSL)

2. **Sideload the manifest**:
   - Open Excel
   - Go to File > Options > Trust Center > Trust Center Settings > Trusted Add-in Catalogs
   - Add your web server URL
   - Go to Insert > My Add-ins > Upload My Add-in
   - Select `manifest.xml`

### Quick Start (Local Development)

**Important**: Office.js add-ins require HTTPS, even for local development.

#### Option 1: Quick Setup (Recommended - Uses mkcert)

1. Install mkcert (creates trusted certificates):
   ```bash
   # macOS
   brew install mkcert
   
   # Windows (with Chocolatey)
   choco install mkcert
   ```

2. Install local CA and generate certificate:
   ```bash
   mkcert -install
   mkdir -p certs
   mkcert -key-file certs/key.pem -cert-file certs/cert.pem localhost 127.0.0.1 ::1
   ```

3. Install and start http-server:
   ```bash
   npm install -g http-server
   http-server -p 3000 -S -C certs/cert.pem -K certs/key.pem
   ```

Or simply run: `./start-server.sh` (automatically uses mkcert if available)

4. Sideload the manifest in Excel:
   - Go to **Insert** → **Add-ins** → **My Add-ins** → **Upload My Add-in**
   - Select `manifest.xml`

The `manifest.xml` is already configured for `https://localhost:3000`.

**Note**: If you get certificate errors, mkcert is the recommended solution as it creates system-trusted certificates. See [setup-local-server.md](setup-local-server.md) for alternatives and troubleshooting.

## Usage

1. Open the Excel workbook you want to test
2. Open the add-in task pane (via the ribbon button or Insert > My Add-ins)
3. Click "Load Test File" and select a JSON test file
4. Review the test inputs and assertions
5. Click "Run Test" to execute
6. View the results showing which assertions passed or failed

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
- Run tests from files, not pasted in
- display options (hide passed?)
- fix icon
- bundle and deploy
- easier setup script

