/**
 * Excel Unit Test Runner
 * Core logic for executing tests with state snapshot/restore
 */

console.log('=== test-runner.js LOADING ===');

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
            throw new Error(`Failed to access worksheet "${worksheetName}": ${error.message}`, error);
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
            throw new Error(`Failed to apply input to worksheet "${worksheetName}": ${error.message}`, error);
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
            throw new Error(`Failed to read output from cell "${cellAddress}": ${error.message}`, error);
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
    
    // Group cells by worksheet
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
    
    // Restore cells
    for (const [worksheetName, cellList] of Object.entries(cellsByWorksheet)) {
        try {
            const worksheet = workbook.worksheets.getItem(worksheetName);
            
            for (const cell of cellList) {
                try {
                    const range = worksheet.getRange(cell.cellAddress);
                    const state = cell.state;
                    
                    // Try formula first, otherwise value
                    if (state.formula && state.formula !== '') {
                        range.formulas = [[state.formula]];
                    } else {
                        range.values = [[state.value]];
                    }
                } catch (cellError) {
                    console.error(`Failed to restore cell ${cell.cellAddress}:`, cellError);
                }
            }
        } catch (error) {
            console.error(`Failed to restore worksheet "${worksheetName}":`, error);
        }
    }
    
    await context.sync();
}

/**
 * Public function to run multiple tests with suite-level locking
 */
async function runTestSuite(testCases) {
    return Excel.run(async (context) => {
        // Collect all cell addresses for snapshot
        const allCellAddresses = new Set();
        
        for (const testCase of testCases) {
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
        }
        
        // Create snapshot of current state
        let snapshot = null;
        try {
            const cellAddressArray = Array.from(allCellAddresses);
            snapshot = await snapshotWorksheetState(context, cellAddressArray);
        } catch (error) {
            console.error(`Warning: Failed to create suite-level snapshot:`, error);
        }
        
        const allResults = [];
        let passedCount = 0;
        
        try {
            // Run tests sequentially
            for (let i = 0; i < testCases.length; i++) {
                console.log(`Running test ${i + 1}/${testCases.length}: ${testCases[i].name || 'Unnamed Test'}`);
                
                try {
                    const result = await runTestWithoutProtection(testCases[i], context);
                    allResults.push(result);
                    if (result.passed) {
                        passedCount++;
                    }
                } catch (error) {
                    const testName = testCases[i].name || `Test ${i + 1}`;
                    console.log(`Error running test ${testName}:`, error);
                    // If a test fails, add error result but continue with other tests
                    allResults.push({
                        testName,
                        passed: false,
                        assertionResults: [],
                        error: error.message
                    });
                }
            }
            
            return {
                results: allResults,
                passedCount: passedCount,
                totalCount: testCases.length
            };
            
        } finally {
            // Restore state at the end of the suite
            if (snapshot) {
                try {
                    await restoreState(context, snapshot);
                    console.log("State restored successfully");
                } catch (restoreError) {
                    console.error("Failed to restore state:", restoreError);
                }
            }
        }
    });
}
/**
 * Private function to run a single test without protection (for use within test suites)
 */
async function runTestWithoutProtection(testCase, context) {
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
    const results = [];
    for (let i = 0; i < testCase.assertions.length; i++) {
        const assertion = testCase.assertions[i];
        const actualValue = outputs[assertion.cell];
        
        // Convert expected value to match actual value type
        let expectedValue = assertion.equals; // Test data uses 'equals' field
        if (typeof actualValue === 'number' && typeof expectedValue === 'string') {
            expectedValue = parseFloat(expectedValue);
        } else if (typeof actualValue === 'boolean') {
            expectedValue = expectedValue.toLowerCase() === 'true';
        }
        
        const passed = actualValue === expectedValue;
        
        // Calculate difference for numeric values
        let difference = null;
        let tolerance = null;
        if (typeof actualValue === 'number' && typeof expectedValue === 'number') {
            difference = actualValue - expectedValue;
            tolerance = assertion.tolerance || 0;
        }
        
        results.push({
            cell: assertion.cell,
            expected: expectedValue,
            actual: actualValue,
            passed: passed,
            message: assertion.message || `Cell ${assertion.cell} should be ${expectedValue}`,
            difference: difference,
            tolerance: tolerance
        });
    }
    
    return {
        name: testCase.name || 'Unnamed Test',
        passed: results.every(r => r.passed),
        assertionResults: results,
        error: null
    };
}

// Export functions globally for Office.js add-in
console.log('=== EXPORTING ExcelTestRunner ===');
window.ExcelTestRunner = {
    runTestSuite: runTestSuite,
    parseCellAddress: parseCellAddress
};
console.log('=== ExcelTestRunner EXPORTED ===:', window.ExcelTestRunner);
console.log('=== SCRIPT COMPLETE ===');

// Also support Node.js/CommonJS for reference
if (typeof module !== 'undefined' && module.exports) {
    module.exports = {
        parseCellAddress: parseCellAddress
    };
}