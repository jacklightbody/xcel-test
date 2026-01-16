/**
 * Excel Unit Test Runner
 * Core logic for executing tests with state snapshot/restore
 */

/**
 * Parses a cell address like "Assumptions!B2" into {worksheetName, cellAddress}
 */
function parseCellAddress(fullAddress) {
    const parts = fullAddress.split('!');
    if (parts.length !== 2) {
        throw new Error(`Invalid cell address format: ${fullAddress}. Expected format: "SheetName!A1"`);
    }
    return {
        worksheetName: parts[0],
        cellAddress: parts[1]
    };
}

/**
 * Snapshot the current state of all cells referenced in inputs and assertions
 */
async function snapshotWorksheetState(context, cellAddresses) {
    const snapshot = {};
    const workbook = context.workbook;
    
    // Group cells by worksheet
    const cellsByWorksheet = {};
    for (const fullAddress of cellAddresses) {
        const parsed = parseCellAddress(fullAddress);
        if (!cellsByWorksheet[parsed.worksheetName]) {
            cellsByWorksheet[parsed.worksheetName] = [];
        }
        cellsByWorksheet[parsed.worksheetName].push(parsed.cellAddress);
    }
    
    // Use parallel arrays to preserve proxy object references
    const fullAddresses = [];
    const ranges = [];
    
    // Step 1: Get all ranges and load their properties
    for (const [worksheetName, cellAddresses] of Object.entries(cellsByWorksheet)) {
        try {
            const worksheet = workbook.worksheets.getItem(worksheetName);
            
            for (const cellAddress of cellAddresses) {
                const range = worksheet.getRange(cellAddress);
                const fullAddress = `${worksheetName}!${cellAddress}`;
                
                // Store in parallel arrays - preserves proxy reference
                fullAddresses.push(fullAddress);
                ranges.push(range);
                
                // Load properties on the range object
                range.load("values");
                range.load("formulas");
            }
        } catch (error) {
            throw new Error(`Failed to access worksheet "${worksheetName}": ${error.message}`);
        }
    }
    
    // Step 2: Sync to populate properties
    await context.sync();
    
    // Step 3: Extract values using same array indices - preserves reference
    for (let i = 0; i < ranges.length; i++) {
        const fullAddress = fullAddresses[i];
        const range = ranges[i]; // Access range from parallel array
        
        // Access properties on the exact range object that was loaded
        snapshot[fullAddress] = {
            value: range.values[0][0],
            formula: range.formulas[0][0] || null
        };
    }
    
    return snapshot;
}

/**
 * Apply input values to the specified cells
 */
async function applyInputs(context, inputs) {
    const workbook = context.workbook;
    
    // Group inputs by worksheet for batch operations
    const inputsByWorksheet = {};
    for (const [fullAddress, value] of Object.entries(inputs)) {
        const parsed = parseCellAddress(fullAddress);
        if (!inputsByWorksheet[parsed.worksheetName]) {
            inputsByWorksheet[parsed.worksheetName] = [];
        }
        inputsByWorksheet[parsed.worksheetName].push({
            cellAddress: parsed.cellAddress,
            value: value
        });
    }
    
    // Apply inputs to each worksheet
    for (const [worksheetName, inputList] of Object.entries(inputsByWorksheet)) {
        try {
            const worksheet = workbook.worksheets.getItem(worksheetName);
            
            for (const input of inputList) {
                const range = worksheet.getRange(input.cellAddress);
                // Set value (this will overwrite any existing formula)
                range.values = [[input.value]];
            }
        } catch (error) {
            throw new Error(`Failed to apply input to worksheet "${worksheetName}": ${error.message}`);
        }
    }
    
    await context.sync();
}

/**
 * Force Excel to recalculate all formulas
 */
async function forceRecalculate(context) {
    const application = context.workbook.application;
    application.calculate(Excel.CalculationType.full);
    await context.sync();
    
    // Wait a bit to ensure calculation completes
    // Note: There's no direct event for calculation completion in Office.js,
    // so we use a small delay. For production, you might want to poll
    // application.getCalculationState() if needed.
    await new Promise(resolve => setTimeout(resolve, 100));
}

/**
 * Read the actual values from assertion cells
 */
async function readOutputs(context, assertionCells) {
    const workbook = context.workbook;
    const outputs = {};
    
    // Use parallel arrays to preserve proxy object references
    const fullAddresses = [];
    const ranges = [];
    
    // Step 1: Get all ranges and load their properties
    for (const cellAddress of assertionCells) {
        try {
            const parsed = parseCellAddress(cellAddress);
            const worksheet = workbook.worksheets.getItem(parsed.worksheetName);
            const range = worksheet.getRange(parsed.cellAddress);
            
            // Store in parallel arrays - preserves proxy reference
            fullAddresses.push(cellAddress);
            ranges.push(range);
            
            // Load properties on the range object
            range.load("values");
        } catch (error) {
            throw new Error(`Failed to read output from cell "${cellAddress}": ${error.message}`);
        }
    }
    
    // Step 2: Sync to populate properties
    await context.sync();
    
    // Step 3: Extract values using same array indices - preserves reference
    for (let i = 0; i < ranges.length; i++) {
        const fullAddress = fullAddresses[i];
        const range = ranges[i]; // Access range from parallel array
        
        // Access properties on the exact range object that was loaded
        outputs[fullAddress] = range.values[0][0];
    }
    
    return outputs;
}

/**
 * Evaluate assertions against actual values
 */
function evaluateAssertions(outputs, assertions) {
    const results = [];
    let allPassed = true;
    
    for (const assertion of assertions) {
        const cellAddress = assertion.cell;
        const expected = assertion.equals;
        const tolerance = assertion.tolerance || 0;
        const actual = outputs[cellAddress];
        
        let passed = false;
        let difference = null;
        
        if (actual === undefined || actual === null) {
            passed = false;
            difference = null;
        } else if (typeof actual === 'number' && typeof expected === 'number') {
            difference = Math.abs(actual - expected);
            passed = difference <= tolerance;
        } else {
            // For non-numeric values, exact match
            passed = actual === expected;
            if (!passed && typeof actual === 'number' && typeof expected === 'number') {
                difference = Math.abs(actual - expected);
            }
        }
        
        if (!passed) {
            allPassed = false;
        }
        
        results.push({
            cell: cellAddress,
            expected: expected,
            actual: actual,
            tolerance: tolerance,
            difference: difference,
            passed: passed
        });
    }
    
    return {
        allPassed: allPassed,
        results: results
    };
}

/**
 * Restore the workbook state from a snapshot
 */
async function restoreState(context, snapshot) {
    const workbook = context.workbook;
    
    // Group by worksheet
    const cellsByWorksheet = {};
    for (const [fullAddress, state] of Object.entries(snapshot)) {
        const parsed = parseCellAddress(fullAddress);
        if (!cellsByWorksheet[parsed.worksheetName]) {
            cellsByWorksheet[parsed.worksheetName] = [];
        }
        cellsByWorksheet[parsed.worksheetName].push({
            cellAddress: parsed.cellAddress,
            state: state
        });
    }
    
    // Restore each worksheet
    for (const [worksheetName, cellList] of Object.entries(cellsByWorksheet)) {
        try {
            const worksheet = workbook.worksheets.getItem(worksheetName);
            
            for (const cell of cellList) {
                const range = worksheet.getRange(cell.cellAddress);
                const state = cell.state;
                
                // Restore formula if it existed, otherwise restore value
                if (state.formula && state.formula !== '') {
                    range.formulas = [[state.formula]];
                } else {
                    range.values = [[state.value]];
                }
            }
        } catch (error) {
            // Log error but continue restoring other cells
            console.error(`Failed to restore cell in worksheet "${worksheetName}": ${error.message}`);
        }
    }
    
    await context.sync();
}

/**
 * Main function to run a single test case
 */
async function runTest(testCase) {
    return Excel.run(async (context) => {
        // Collect all cell addresses that need to be snapshotted
        const allCellAddresses = new Set();
        
        // Add input cells
        if (testCase.inputs) {
            for (const cellAddress of Object.keys(testCase.inputs)) {
                allCellAddresses.add(cellAddress);
            }
        }
        
        // Add assertion cells
        if (testCase.assertions) {
            for (const assertion of testCase.assertions) {
                allCellAddresses.add(assertion.cell);
            }
        }
        
        const cellAddressArray = Array.from(allCellAddresses);
        
        // Snapshot current state
        let snapshot = null;
        try {
            snapshot = await snapshotWorksheetState(context, cellAddressArray);
        } catch (error) {
            throw new Error(`Failed to snapshot state: ${error.message}`);
        }
        
        try {
            // Apply inputs
            if (testCase.inputs && Object.keys(testCase.inputs).length > 0) {
                await applyInputs(context, testCase.inputs);
            }
            
            // Force recalculation
            await forceRecalculate(context);
            
            // Read outputs
            const assertionCells = testCase.assertions.map(a => a.cell);
            const outputs = await readOutputs(context, assertionCells);
            
            // Evaluate assertions
            const evaluation = evaluateAssertions(outputs, testCase.assertions);
            
            return {
                testName: testCase.name || 'Unnamed Test',
                passed: evaluation.allPassed,
                assertionResults: evaluation.results
            };
            
        } finally {
            // Always restore state, even if there was an error
            if (snapshot) {
                try {
                    await restoreState(context, snapshot);
                } catch (restoreError) {
                    // Log but don't throw - we want the test results even if restore fails
                    console.error(`Warning: Failed to restore state: ${restoreError.message}`);
                }
            }
        }
    });
}

// Export functions globally for Office.js add-in
if (typeof window !== 'undefined') {
    window.ExcelTestRunner = {
        runTest: runTest,
        parseCellAddress: parseCellAddress
    };
}

// Also support Node.js/CommonJS for reference
if (typeof module !== 'undefined' && module.exports) {
    module.exports = {
        runTest: runTest,
        parseCellAddress: parseCellAddress
    };
}