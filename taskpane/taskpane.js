/* global Excel, Office */

let currentTest = null;
let currentTests = null; // Array of tests if multiple tests are loaded

Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        // Initialize UI once Office.js and DOM are ready
        if (document.readyState === 'loading') {
            document.addEventListener("DOMContentLoaded", () => {
                initializeUI();
            });
        } else {
            // DOM is already ready
            initializeUI();
        }
    }
});

function initializeUI() {
    // Test runner should be loaded via script tag in HTML
    setupEventHandlers();
    
    // Show test section by default
    document.getElementById('test-section').style.display = 'block';
}

// Ensure UI is initialized when DOM is ready
if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', initializeUI);
} else {
    initializeUI();
}

// Track if handlers are already set up to avoid duplicates
let handlersSetup = false;

function setupEventHandlers() {
    // Prevent duplicate setup
    if (handlersSetup) {
        console.log('Event handlers already set up, skipping...');
        return;
    }
    
    const testJsonInput = document.getElementById('test-json-input');
    const runTestButton = document.getElementById('run-test-button');
    
    console.log('Setting up event handlers...', {
        testJsonInput: !!testJsonInput,
        runTestButton: !!runTestButton
    });
    
    // Load & Run test button - combines loading and execution
    if (runTestButton && testJsonInput) {
        runTestButton.addEventListener('click', async function(e) {
            console.log('Load & Run Test button clicked');
            await handleLoadAndRunTest();
        });
    }
    
    // Also allow Enter+Ctrl/Cmd to load and run test
    if (testJsonInput) {
        testJsonInput.addEventListener('keydown', async function(e) {
            if ((e.ctrlKey || e.metaKey) && e.key === 'Enter') {
                e.preventDefault();
                await handleLoadAndRunTest();
            }
        });
    }
    
    handlersSetup = true;
}

async function handleLoadAndRunTest() {
    const testJsonInput = document.getElementById('test-json-input');
    const runTestButton = document.getElementById('run-test-button');
    
    if (!testJsonInput) {
        showError('Test input element not found');
        return;
    }
    
    const jsonText = testJsonInput.value.trim();
    if (!jsonText) {
        showError('Please paste or type JSON test content');
        return;
    }
    
    // Disable button during processing
    if (runTestButton) {
        runTestButton.disabled = true;
        runTestButton.querySelector('.ms-Button-label').textContent = 'Loading...';
    }
    
    clearResults();
    clearErrors();
    
    try {
        console.log('Parsing JSON from textarea...');
        const testData = JSON.parse(jsonText);
        console.log('JSON parsed successfully', testData);
        
        // Support both single test object and array of tests
        if (Array.isArray(testData)) {
            currentTests = testData;
            currentTest = null;
            displayMultipleTestInfo(testData);
            console.log('Loaded', testData.length, 'tests');
        } else {
            currentTests = null;
            currentTest = testData;
            displayTestInfo(testData);
            console.log('Loaded single test:', testData.name);
        }
        
        // Show success feedback
        testJsonInput.style.borderColor = '#107c10';
        
        // Now run the test(s)
        if (runTestButton) {
            runTestButton.querySelector('.ms-Button-label').textContent = 'Running...';
        }
        
        // Determine which tests to run and execute them
        const testsToRun = currentTests || [currentTest];
        await executeTests(testsToRun, runTestButton);
        
        // Reset border color
        setTimeout(function() {
            testJsonInput.style.borderColor = '';
        }, 1000);
    } catch (error) {
        console.error('Error parsing or running test:', error);
        showError(`Failed to parse JSON: ${error.message}`);
        testJsonInput.style.borderColor = '#d13438';
        setTimeout(function() {
            testJsonInput.style.borderColor = '';
        }, 2000);
    } finally {
        if (runTestButton) {
            runTestButton.disabled = false;
            runTestButton.querySelector('.ms-Button-label').textContent = 'Load & Run Test';
        }
    }
}


function displayTestInfo(testData) {
    const testInfoDiv = document.getElementById('current-test-info');
    
    let inputsHtml = '';
    if (testData.inputs) {
        inputsHtml = '<p><strong>Inputs:</strong></p><ul>';
        for (const [cell, value] of Object.entries(testData.inputs)) {
            inputsHtml += `<li>${cell} = ${value}</li>`;
        }
        inputsHtml += '</ul>';
    }
    
    let assertionsHtml = '';
    if (testData.assertions) {
        assertionsHtml = '<p><strong>Assertions:</strong></p><ul>';
        for (const assertion of testData.assertions) {
            const tolerance = assertion.tolerance !== undefined ? ` (tolerance: ${assertion.tolerance})` : '';
            assertionsHtml += `<li>${assertion.cell} should equal ${assertion.equals}${tolerance}</li>`;
        }
        assertionsHtml += '</ul>';
    }
    
    testInfoDiv.innerHTML = `
        <h3>${testData.name || 'Unnamed Test'}</h3>
        ${inputsHtml}
        ${assertionsHtml}
    `;
}

function displayMultipleTestInfo(tests) {
    const testInfoDiv = document.getElementById('current-test-info');
    
    let html = `<h3>Test Suite (${tests.length} test${tests.length > 1 ? 's' : ''})</h3>`;
    
    for (let i = 0; i < tests.length; i++) {
        const test = tests[i];
        html += `<div style="margin: 15px 0; padding: 10px; border-left: 3px solid #0078d4; background-color: #f3f2f1;">`;
        html += `<strong>${i + 1}. ${test.name || 'Unnamed Test'}</strong>`;
        
        if (test.inputs && Object.keys(test.inputs).length > 0) {
            html += '<p style="margin: 5px 0;"><small><strong>Inputs:</strong> ';
            const inputs = Object.entries(test.inputs).map(([cell, value]) => `${cell}=${value}`).join(', ');
            html += inputs;
            html += '</small></p>';
        }
        
        if (test.assertions && test.assertions.length > 0) {
            html += `<p style="margin: 5px 0;"><small><strong>Assertions:</strong> ${test.assertions.length}</small></p>`;
        }
        
        html += '</div>';
    }
    
    testInfoDiv.innerHTML = html;
}

// Shared function to execute tests
async function executeTests(testsToRun, buttonElement) {
    const allResults = [];
    let passedCount = 0;
    
    // Run tests sequentially
    for (let i = 0; i < testsToRun.length; i++) {
        // Update button text to show progress
        if (buttonElement) {
            buttonElement.querySelector('.ms-Button-label').textContent = `Running test ${i + 1}/${testsToRun.length}...`;
        }
        
        try {
            let result;
            if (typeof window.ExcelTestRunner === 'undefined' || !window.ExcelTestRunner.runTest) {
                result = await runTestInline(testsToRun[i]);
            } else {
                result = await window.ExcelTestRunner.runTest(testsToRun[i]);
            }
            allResults.push(result);
            if (result.passed) {
                passedCount++;
            }
        } catch (error) {
            // If a test fails, add error result but continue with other tests
            allResults.push({
                testName: testsToRun[i].name || `Test ${i + 1}`,
                passed: false,
                assertionResults: [],
                error: error.message
            });
        }
    }
    
    // Display results
    displayMultipleResults(allResults, passedCount, testsToRun.length);
}

// Inline test runner implementation as fallback for Office.js context
function runTestInline(testCase) {
    // This will use the functions defined in test-runner.js
    // Since we're loading it as a script tag, we need to make sure the functions are accessible
    // For Office.js, we'll recreate the logic here to avoid module issues
    
    return Excel.run(async (context) => {
            // Parse cell address helper
            function parseCellAddress(fullAddress) {
                const parts = fullAddress.split('!');
                if (parts.length !== 2) {
                    throw new Error(`Invalid cell address format: ${fullAddress}`);
                }
                return {
                    worksheetName: parts[0],
                    cellAddress: parts[1]
                };
            }
            
            // Collect all cell addresses
            const allCellAddresses = new Set();
            if (testCase.inputs) {
                for (const cellAddress of Object.keys(testCase.inputs)) {
                    allCellAddresses.add(cellAddress);
                }
            }
            if (testCase.assertions) {
                for (const assertion of testCase.assertions) {
                    allCellAddresses.add(assertion.cell);
                }
            }
            
            const cellAddressArray = Array.from(allCellAddresses);
            
            // Snapshot state
            const snapshot = {};
            const workbook = context.workbook;
            const cellsByWorksheet = {};
            
            // Group cells by worksheet
            for (const fullAddress of cellAddressArray) {
                const parsed = parseCellAddress(fullAddress);
                if (!cellsByWorksheet[parsed.worksheetName]) {
                    cellsByWorksheet[parsed.worksheetName] = [];
                }
                cellsByWorksheet[parsed.worksheetName].push(parsed.cellAddress);
            }
            
            // CRITICAL FIX: Use parallel arrays to preserve proxy object references
            // Office.js proxy objects must be accessed directly without wrapping in objects
            const fullAddresses = [];
            const ranges = [];
            
            // Step 1: Get all ranges and load their properties
            for (const [worksheetName, cellAddresses] of Object.entries(cellsByWorksheet)) {
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
            
            try {
                // Apply inputs
                if (testCase.inputs && Object.keys(testCase.inputs).length > 0) {
                    const inputsByWorksheet = {};
                    for (const [fullAddress, value] of Object.entries(testCase.inputs)) {
                        const parsed = parseCellAddress(fullAddress);
                        if (!inputsByWorksheet[parsed.worksheetName]) {
                            inputsByWorksheet[parsed.worksheetName] = [];
                        }
                        inputsByWorksheet[parsed.worksheetName].push({
                            cellAddress: parsed.cellAddress,
                            value: value
                        });
                    }
                    
                    for (const [worksheetName, inputList] of Object.entries(inputsByWorksheet)) {
                        const worksheet = workbook.worksheets.getItem(worksheetName);
                        for (const input of inputList) {
                            const range = worksheet.getRange(input.cellAddress);
                            range.values = [[input.value]];
                        }
                    }
                    await context.sync();
                }
                
                // Force recalculation
                const application = context.workbook.application;
                application.calculate(Excel.CalculationType.full);
                await context.sync();
                await new Promise(resolve => setTimeout(resolve, 100));
                
                // Read outputs
                const outputs = {};
                const assertionCells = testCase.assertions.map(a => a.cell);
                const outputCellsByWorksheet = {};
                
                for (const cellAddress of assertionCells) {
                    const parsed = parseCellAddress(cellAddress);
                    if (!outputCellsByWorksheet[parsed.worksheetName]) {
                        outputCellsByWorksheet[parsed.worksheetName] = [];
                    }
                    outputCellsByWorksheet[parsed.worksheetName].push({
                        cellAddress: parsed.cellAddress,
                        fullAddress: cellAddress
                    });
                }
                
                // Store output ranges before loading
                const outputRanges = [];
                for (const [worksheetName, cellList] of Object.entries(outputCellsByWorksheet)) {
                    const worksheet = workbook.worksheets.getItem(worksheetName);
                    for (const cell of cellList) {
                        const range = worksheet.getRange(cell.cellAddress);
                        outputRanges.push({ fullAddress: cell.fullAddress, range });
                        range.load("values");
                    }
                }
                
                await context.sync();
                
                // Extract values from loaded output ranges
                for (const { fullAddress, range } of outputRanges) {
                    outputs[fullAddress] = range.values[0][0];
                }
                
                // Evaluate assertions
                const results = [];
                let allPassed = true;
                
                for (const assertion of testCase.assertions) {
                    const cellAddress = assertion.cell;
                    const expected = assertion.equals;
                    const tolerance = assertion.tolerance || 0;
                    const actual = outputs[cellAddress];
                    
                    let passed = false;
                    let difference = null;
                    
                    if (actual === undefined || actual === null) {
                        passed = false;
                    } else if (typeof actual === 'number' && typeof expected === 'number') {
                        difference = Math.abs(actual - expected);
                        passed = difference <= tolerance;
                    } else {
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
                    testName: testCase.name || 'Unnamed Test',
                    passed: allPassed,
                    assertionResults: results
                };
                
            } finally {
                // Restore state
                if (snapshot) {
                    const restoreCellsByWorksheet = {};
                    for (const [fullAddress, state] of Object.entries(snapshot)) {
                        const parsed = parseCellAddress(fullAddress);
                        if (!restoreCellsByWorksheet[parsed.worksheetName]) {
                            restoreCellsByWorksheet[parsed.worksheetName] = [];
                        }
                        restoreCellsByWorksheet[parsed.worksheetName].push({
                            cellAddress: parsed.cellAddress,
                            state: state
                        });
                    }
                    
                    for (const [worksheetName, cellList] of Object.entries(restoreCellsByWorksheet)) {
                        try {
                            const worksheet = workbook.worksheets.getItem(worksheetName);
                            for (const cell of cellList) {
                                const range = worksheet.getRange(cell.cellAddress);
                                const state = cell.state;
                                if (state.formula && state.formula !== '') {
                                    range.formulas = [[state.formula]];
                                } else {
                                    range.values = [[state.value]];
                                }
                            }
                        } catch (error) {
                            console.error(`Failed to restore cell in worksheet "${worksheetName}": ${error.message}`);
                        }
                    }
                    await context.sync();
                }
            }
        });
}

function displayMultipleResults(results, passedCount, totalCount) {
    const resultsSection = document.getElementById('results-section');
    const resultsContent = document.getElementById('results-content');
    
    const allPassed = passedCount === totalCount;
    const summaryClass = allPassed ? 'pass' : 'fail';
    const summaryText = allPassed ? 'ALL PASSED' : `${passedCount}/${totalCount} PASSED`;
    
    let html = `
        <div class="test-summary ${summaryClass}">
            Test Suite: ${summaryText}
        </div>
    `;
    
    for (let i = 0; i < results.length; i++) {
        const result = results[i];
        const resultClass = result.passed ? 'pass' : 'fail';
        const resultText = result.passed ? 'PASSED' : 'FAILED';
        
        html += `
            <div class="result-item ${resultClass}" style="margin-top: 15px;">
                <h4>${i + 1}. ${result.testName} - ${resultText}</h4>
        `;
        
        if (result.error) {
            html += `<div class="error-message" style="margin: 5px 0; padding: 10px;">Error: ${result.error}</div>`;
        }
        
        for (const assertionResult of result.assertionResults) {
            const assertionClass = assertionResult.passed ? 'pass' : 'fail';
            let detailsHtml = '';
            
            if (assertionResult.passed) {
                if (assertionResult.difference !== null) {
                    detailsHtml = `<div class="assertion-details">Actual: ${assertionResult.actual}, Expected: ${assertionResult.expected}, Difference: ${assertionResult.difference}</div>`;
                } else {
                    detailsHtml = `<div class="assertion-details">Actual: ${assertionResult.actual}, Expected: ${assertionResult.expected}</div>`;
                }
            } else {
                if (assertionResult.difference !== null) {
                    detailsHtml = `<div class="assertion-details">Actual: ${assertionResult.actual}, Expected: ${assertionResult.expected}, Difference: ${assertionResult.difference} (tolerance: ${assertionResult.tolerance})</div>`;
                } else {
                    detailsHtml = `<div class="assertion-details">Actual: ${assertionResult.actual}, Expected: ${assertionResult.expected}</div>`;
                }
            }
            
            html += `
                <div class="assertion ${assertionClass}">
                    <strong>${assertionResult.cell}</strong>
                    ${detailsHtml}
                </div>
            `;
        }
        
        html += '</div>';
    }
    
    resultsContent.innerHTML = html;
    resultsSection.style.display = 'block';
}

function showError(message) {
    const errorSection = document.getElementById('error-section');
    const errorContent = document.getElementById('error-content');
    errorContent.textContent = message;
    errorSection.style.display = 'block';
}

function clearResults() {
    document.getElementById('results-section').style.display = 'none';
    document.getElementById('results-content').innerHTML = '';
}

function clearErrors() {
    document.getElementById('error-section').style.display = 'none';
    document.getElementById('error-content').textContent = '';
}