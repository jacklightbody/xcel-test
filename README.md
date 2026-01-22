# Excel Unit Test Runner

Test your Excel formulas like code. Write unit tests to validate calculations, catch errors early, and refactor with confidence.

## Why Use This?

- **Catch errors before they matter**: Validate complex formula logic automatically
- **Refactor fearlessly**: Change formulas knowing tests will catch breaks
- **Document assumptions**: Tests serve as living documentation of expected behavior
- **Save time**: Run comprehensive checks in seconds vs. manual verification

## Quick Start

![Demo](https://raw.githubusercontent.com/jacklightbody/xcel-test/refs/heads/main/examples/example-test-run.gif)

1. Run setup: `./setup.sh && ./start.sh`
2. Open your Excel workbook
3. Open the add-in (Insert → My Add-ins → "Excel Unit Test Runner")
4. Load or paste your test JSON
5. Hit **Run Test** (or **Ctrl+Enter** / **Cmd+Enter**)
6. See which assertions passed or failed

## How It Works

Each test runs in isolation without permanently changing your workbook:

1. **Snapshot**: Captures original values and formulas from all referenced cells
2. **Apply**: Sets input values as specified
3. **Calculate**: Forces full Excel recalculation
4. **Assert**: Compares actual vs. expected values (with optional tolerance)
5. **Restore**: Returns workbook to original state

Your workbook is **always restored** to its original state, so you can run tests repeatedly without side effects.

## Test Format

Tests are JSON files with inputs (values to set) and assertions (expected results).

### Test Property Reference

#### Input Properties

| Property | Type | Required | Notes |
|----------|------|----------|-------|
| `cell` | string | One of `cell` or `relativeTo` | Direct cell reference (e.g., "Sheet1!B2") |
| `relativeTo` | object | One of `cell` or `relativeTo` | Position relative to reference cell(s) |
| `value` | any | ✓ Yes | Value to set in the input cell |

#### Assertion Properties

| Property | Type | Required | Notes |
|----------|------|----------|-------|
| `cell` | string | One of `cell` or `relativeTo` | Direct cell reference (e.g., "Sheet1!B5") |
| `relativeTo` | object | One of `cell` or `relativeTo` | Position relative to reference cell(s) |
| `equals` | any | ✓ Yes | Expected value |
| `tolerance` | number | No | Allowed difference for numeric comparisons |

#### RelativeTo Object (Offset-based)

| Property | Type | Required | Notes |
|----------|------|----------|-------|
| `sheet` | string | ✓ Yes | Sheet name |
| `referenceCell` | string | ✓ Yes | Text content of reference cell |
| `colOffset` | number | ✓ Yes | Columns from reference (-1 = left, 1 = right) |
| `rowOffset` | number | ✓ Yes | Rows from reference (-1 = up, 1 = down) |

#### RelativeTo Object (Intersection-based)

| Property | Type | Required | Notes |
|----------|------|----------|-------|
| `sheet` | string | ✓ Yes | Sheet name |
| `referenceColCell` | string | ✓ Yes | Text content of cell defining column |
| `referenceRowCell` | string | ✓ Yes | Text content of cell defining row |

### Examples

**Example 1: Direct Cell References**

Test specific cells with known addresses:

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

**Example 2: Offset References**

Find cells relative to labeled cells:

```json
[
  {
    "name": "Growth rate scenario",
    "inputs": [
      {
        "relativeTo": {
          "sheet": "Assumptions",
          "referenceCell": "Growth Rate",
          "colOffset": 1,
          "rowOffset": 0
        },
        "value": 0.05
      }
    ],
    "assertions": [
      {
        "cell": "Outputs!E12",
        "equals": 1234567,
        "tolerance": 1
      }
    ]
  }
]
```

**Example 3: Intersection References**

Find cells at the intersection of row/column headers:

```json
[
  {
    "name": "Fiscal year total",
    "inputs": [
      {
        "cell": "Sheet1!B2",
        "value": 3
      }
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
      }
    ]
  }
]
```

## Setup

### Automated Setup (Recommended)

```bash
./setup.sh && ./start.sh
```

This will:
- Install mkcert (if needed) and generate trusted HTTPS certificates
- Auto-install the manifest for Mac Excel users
- Start the development server

After initial setup, just run `./start.sh` to start the server.

### Manual Setup

If automated setup fails:

1. **Install mkcert**:
   ```bash
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

3. **Start server**: `./start.sh`

4. **Install add-in**:
   - **Mac**: Auto-installed! Go to Insert → My Add-ins → "Excel Unit Test Runner"
   - **Other**: Insert → Add-ins → My Add-ins → Upload My Add-in → select `manifest.xml`

![Where to find the add-in](https://raw.githubusercontent.com/jacklightbody/xcel-test/refs/heads/main/examples/add-add-in.png)

## Project Structure

```
/
├── manifest.xml              # Office.js add-in manifest
├── taskpane/
│   ├── taskpane.html        # Task pane UI
│   ├── taskpane.js          # UI logic and test execution
│   └── taskpane.css         # Styling
├── scripts/
│   ├── test-runner.js       # Core test execution logic
│   └── cell-resolver.js     # Cell reference resolution
├── tests/
│   └── sample-test.json     # Example test file
└── README.md
```

## Troubleshooting

| Issue | Solution |
|-------|----------|
| "Failed to access worksheet" | Sheet names are case-sensitive—verify exact match |
| "Invalid cell address format" | Use format "SheetName!A1" |
| "No cell found containing text" | Ensure reference text exists exactly in specified sheet |
| "Multiple cells found containing text" | Reference text must be unique in worksheet |
| "Invalid offset" | Offsets cannot result in negative cell positions |
| Calculations not updating | Complex models may need more than default 100ms wait |
| State not restoring properly | Check browser console; formulas restore before values |

## Known Limitations

- Tests modify only specified input cells, but restoration touches all referenced cells
- Large workbooks with extensive calculations may be slow to snapshot/restore
- Requires web server (cannot run from `file://` protocol)
- Reference text must be unique within each worksheet

## Roadmap

- [ ] UI to help create tests interactively
- [ ] Publish to Microsoft add-in store for easy installation
- [ ] Guard mode: auto-run tests on workbook save
- [ ] Lock workbook during test execution
- [ ] Optimize snapshot/restore to run once per suite
- [ ] Parallel test execution for large suites
- [ ] Allow cancelling in-progress tests
- [ ] Support for boolean and string assertions