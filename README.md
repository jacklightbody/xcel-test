# Excel Unit Test Runner

An Office.js Excel add-in that allows you to write and run unit tests for Excel workbooks to validate the correctness of a set of formulas you've created.

## Usage
![Demo](https://raw.githubusercontent.com/jacklightbody/xcel-test/refs/heads/main/examples/example-test-run.gif)
1. Open the Excel workbook you want to test
2. Open the add-in task pane (via the ribbon button or Insert > My Add-ins)
3. Choose your input method:
   - **Paste JSON**: Copy and paste JSON test content directly
   - **Load File**: Select a JSON test file from your computer
4. Review the test inputs and assertions that are displayed
5. Click "Run Test" to execute (or use **Ctrl+Enter** / **Cmd+Enter** hotkey)
6. View the results showing which assertions passed or failed

## Test Format

Test files are JSON files that contain a list of tests you want to run. Each test has a set of input cells and the values to override to, as well as a set of output cells and their expected values given the inputs.

### Example 1: Simple Cell Reference
If you know all cell addresses, you can easily test that specific addresses evaluate to given values with certain inputs.
```json
[
  {
    "name": "Base case revenue",
    "inputs": [
      {
        "cell": "Sheet1!B2",
        "value": 3
      },
      {
        "cell": "Sheet1!B3", 
        "value": 120000
      }
    ],
    "assertions": [
      {
        "cell": "Sheet1!B5",
        "equals": 40000,
        "tolerance": 1
      }
    ]
  }
]
```

### Example 2: Offset reference
If instead you only know a cell position relative to another cell

```json
[
  {
    "name": "Offset references",
    "inputs": [
      {
        "relativeTo": {
          "sheet": "Assumptions",
          "referenceCell": "Growth Rate",
          "colOffset": -1, # 1 col to the left
          "rowOffset": 1
        },
        "value": 0.05
      },
    ],
    "assertions": [
      {
        "cell": "Outputs!E12",
        "equals": 1234567,
        "tolerance": 1
      },
    ]
  }
]
```

### Example 3: Relative reference
If instead you only know a cell position relative to two other cells

```json
[
  {
    "name": "Relative References",
    "inputs": [
      {
        "cell": "Sheet1!B2",
        "value": 3
      },
    ],
    "assertions": [
      {
        "relativeTo": {
          "sheet": "Sheet1",
          "referenceColCell": "Last Fiscal Year",
          "referenceRowCell": "Total"
        },
        "equals": 100,
        "tolerance": 1
      },
    ]
  }
]
```
### Validation Rules

- **Mutually Exclusive**: Each input/assertion must have either `cell` OR `relativeTo`, never both
- **Required Properties**:
  - Inputs: `value` is required
  - Assertions: `equals` is required
- **Validation Timing**: All test cases are validated before any tests run
- **Error Handling**: Multiple text matches throw clear errors to prevent ambiguity

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
   ./start.sh
   ```

4. **Launch Excel**:
   - **Mac users**: The manifest is auto-installed! Just go to **Inert** → **My Add-ins** and select "Excel Unit Test Runner"
   - **Other users**: Go to **Insert** → **Add-ins** → **My Add-ins** → **Upload My Add-in** and select `manifest.xml`

![where to find the add-in](https://raw.githubusercontent.com/jacklightbody/xcel-test/refs/heads/main/examples/add-add-in.png)


## How It Works

For each test, the add-in performs the following steps:

1. **Snapshot State**: Captures current values and formulas for all cells referenced in inputs and assertions
2. **Apply Inputs**: Sets the input values as specified in the test
3. **Force Calculation**: Triggers Excel's full calculation to ensure all dependent formulas recalculate
4. **Read Outputs**: Retrieves the actual calculated values from assertion cells
5. **Evaluate Assertions**: Compares actual vs expected values (with tolerance for numeric comparisons)
6. **Restore State**: Restores all original values and formulas, ensuring the workbook is unchanged

This means the unit tests both **preserves state** and **exactly match** the native excel behavior.

## Repo Structure

```
/
├── manifest.xml              # Office.js add-in manifest
├── taskpane/
│   ├── taskpane.html        # Task pane UI
│   ├── taskpane.js          # UI logic and test execution
│   └── taskpane.css         # Styling
├── scripts/
│   ├── test-runner.js       # Core test execution logic
│   └── cell-resolver.js     # Cell reference resolution utilities
├── tests/
│   └── sample-test.json     # Example test file
└── README.md                # This file
```

## Limitations

- Tests should only modify cells specified in the inputs - other cells are snapshotted but restoring them may affect unrelated workbook state
- Large workbooks with extensive calculations may take time to snapshot/restore
- The add-in requires a web server to function (cannot run from `file://` protocol)

## Troubleshooting

- **"Failed to access worksheet"**: Ensure worksheet names match exactly (case-sensitive)
- **"Invalid cell address format"**: Cell addresses must be in format "SheetName!A1"
- **"No cell found containing text"**: Ensure the reference text exists exactly as written in the specified worksheet
- **"Multiple cells found containing text"**: Reference text must be unique within the worksheet
- **"Invalid offset"**: Column and row offsets cannot result in negative cell positions
- **"Validation failed"**: Check that each input/assertion has either `cell` OR `relativeTo`, not both
- **Calculation not updating**: The add-in waits 100ms after forcing calculation; complex models may need more time
- **State not restoring**: Check browser console for restore errors; formulas may need to be restored before values

## Todo

- UI to help create tests
- Bundle and deploy to msft so installation is easy
- Guard mode to retrigger on save automatically
- Locking. Prevent (or at least detect) user edits while tests are running
- Snapshot immprovements. Can we snapshot and restore once across every test case rather than one per test?
- Parallelism or some other method to speed up for large tests suites
- Allow interrupting/cancelling a test
- Handle different types of outputs (true/false, string) not just floats