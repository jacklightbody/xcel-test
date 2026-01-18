/* global Excel, Office */

let currentTest = null;
let currentTests = null; // Array of tests if multiple tests are loaded
let currentInputMethod = 'file'; // 'paste' or 'file'
let loadedFileName = null;
let enableLocking = false; // Locking toggle - OFF by default
let isTestRunning = false; // Track if tests are currently running

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

function switchInputMethod(method) {
    currentInputMethod = method;
    
    const pasteTab = document.getElementById('paste-tab');
    const fileTab = document.getElementById('file-tab');
    const pasteSection = document.getElementById('paste-input-section');
    const fileSection = document.getElementById('file-input-section');
    
    // Update tab active states
    if (method === 'paste') {
        pasteTab.classList.add('active');
        fileTab.classList.remove('active');
        pasteSection.style.display = 'block';
        fileSection.style.display = 'none';
    } else {
        fileTab.classList.add('active');
        pasteTab.classList.remove('active');
        fileSection.style.display = 'block';
        pasteSection.style.display = 'none';
    }
}

function initializeUI() {
    setupEventHandlers();
    setupHotkey();
    
    // Show test section by default
    document.getElementById('test-section').style.display = 'block';
    switchInputMethod(currentInputMethod);
}

// Ensure UI is initialized when DOM is ready
if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', initializeUI);
} else {
    initializeUI();
}

// Track if handlers are already set up to avoid duplicates
let handlersSetup = false;

// Store current results to enable filtering
let currentResults = [];
let currentPassedCount = 0;
let currentTotalCount = 0;

function setupEventHandlers() {
    // Prevent duplicate setup
    if (handlersSetup) {
        return;
    }
    
    const testJsonInput = document.getElementById('test-json-input');
    const runTestButton = document.getElementById('run-test-button');
    const pasteTab = document.getElementById('paste-tab');
    const fileTab = document.getElementById('file-tab');
    const fileSelectButton = document.getElementById('file-select-button');
    const testFileInput = document.getElementById('test-file-input');
    const hidePassedTestsCheckbox = document.getElementById('hide-passed-tests');
    

    
    // Tab switching
    if (pasteTab && fileTab) {
        pasteTab.addEventListener('click', () => switchInputMethod('paste'));
        fileTab.addEventListener('click', () => switchInputMethod('file'));
    }
    
    // File selection
    if (fileSelectButton && testFileInput) {
        fileSelectButton.addEventListener('click', () => testFileInput.click());
        testFileInput.addEventListener('change', handleFileSelect);
    }
    
    // Hide passed tests checkbox
    if (hidePassedTestsCheckbox) {
        hidePassedTestsCheckbox.addEventListener('change', filterAndDisplayResults);
    }
    

    
    // Display options toggle
    const displayOptionsToggle = document.getElementById('display-options-toggle');
    if (displayOptionsToggle) {
        displayOptionsToggle.addEventListener('click', toggleDisplayOptions);
    }
    
    // Load & Run test button - combines loading and execution
    if (runTestButton) {
        runTestButton.addEventListener('click', async function(e) {
            await handleLoadAndRunTest();
        });
    }
    
    // Also allow Enter+Ctrl/Cmd to load and run test (only for paste method)
    if (testJsonInput) {
        testJsonInput.addEventListener('keydown', async function(e) {
            if ((e.ctrlKey || e.metaKey) && e.key === 'Enter' && currentInputMethod === 'paste') {
                e.preventDefault();
                await handleLoadAndRunTest();
            }
        });
    }
    
    handlersSetup = true;
}

function setupHotkey() {
    // Add Ctrl/Cmd + Enter hotkey to run tests
    document.addEventListener('keydown', async function(e) {
        // Check for Ctrl/Cmd + Enter combination
        if ((e.ctrlKey || e.metaKey) && e.key === 'Enter') {
            e.preventDefault();
            await handleLoadAndRunTest();
        }
    });
}

function toggleDisplayOptions() {
    const displayOptions = document.getElementById('display-options');
    const toggleIcon = document.getElementById('toggle-icon');
    
    if (displayOptions.style.display === 'none' || displayOptions.style.display === '') {
        displayOptions.style.display = 'block';
        toggleIcon.textContent = '▲';
    } else {
        displayOptions.style.display = 'none';
        toggleIcon.textContent = '▼';
    }
}

function handleFileSelect(event) {
    const file = event.target.files[0];
    if (!file) {
        return;
    }
    
    if (!file.name.toLowerCase().endsWith('.json')) {
        showError('Please select a JSON file');
        return;
    }
    
    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const jsonText = e.target.result;
            const testJsonInput = document.getElementById('test-json-input');
            const fileNameDisplay = document.getElementById('file-name-display');
            
            // Fill the textarea with file content (for consistency with existing logic)
            testJsonInput.value = jsonText;
            
            // Update filename display
            fileNameDisplay.textContent = file.name;
            loadedFileName = file.name;
            
            // Clear any previous results/errors
            clearResults();
            clearErrors();
            
            // Show success feedback
            fileNameDisplay.style.color = '#107c10';
            setTimeout(() => {
                fileNameDisplay.style.color = '';
            }, 1000);
            
        } catch (error) {
            showError(`Failed to read file: ${error.message}`);
        }
    };
    
    reader.onerror = function() {
        showError('Failed to read file');
    };
    
    reader.readAsText(file);
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
        if (currentInputMethod === 'file') {
            showError('Please select a JSON test file first');
        } else {
            showError('Please paste or type JSON test content');
        }
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
        const testData = JSON.parse(jsonText);
        
        // Support both single test object and array of tests
        if (Array.isArray(testData)) {
            currentTests = testData;
            currentTest = null;
            displayMultipleTestInfo(testData);
        } else {
            currentTests = null;
            currentTest = testData;
            displayTestInfo(testData);
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
            runTestButton.querySelector('.ms-Button-label').textContent = 'Run';
        }
    }
}


function displayTestInfo(testData) {
    const testInfoDiv = document.getElementById('current-test-info');
    testInfoDiv.style.display = 'block';
    
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
    // Set global running state
    isTestRunning = true;
    updateUIForTestState(true);
    
    try {
        const suiteResult = await window.ExcelTestRunner.runTestSuite(testsToRun);
        
        displayMultipleResults(suiteResult.results, suiteResult.passedCount, suiteResult.totalCount);
        
    } finally {
        // Always reset the running state
        isTestRunning = false;
        updateUIForTestState(false);
    }
}



function displayMultipleResults(results, passedCount, totalCount) {
    // Store current results for filtering
    currentResults = results;
    currentPassedCount = passedCount;
    currentTotalCount = totalCount;
    
    // Show display options and filter results
    filterAndDisplayResults();
}

function filterAndDisplayResults() {
    const resultsSection = document.getElementById('results-section');
    const resultsContent = document.getElementById('results-content');
    const displayOptions = document.getElementById('display-options');
    const hidePassedTestsCheckbox = document.getElementById('hide-passed-tests');
    const visibleCountSpan = document.getElementById('visible-count');
    const totalCountSpan = document.getElementById('total-count');
    
    // If no results yet, just update the checkbox state
    if (!currentResults || currentResults.length === 0) {
        return;
    }
    
    const allPassed = currentPassedCount === currentTotalCount;
    const summaryClass = allPassed ? 'pass' : 'fail';
    const summaryText = allPassed ? 'ALL PASSED' : `${currentPassedCount}/${currentTotalCount} PASSED`;
    
    // Filter results based on checkbox
    const hidePassedTests = hidePassedTestsCheckbox && hidePassedTestsCheckbox.checked;
    let filteredResults = currentResults;
    let visibleCount = currentResults.length;
    
    if (hidePassedTests) {
        filteredResults = currentResults.filter(result => !result.passed);
        visibleCount = filteredResults.length;
    }
    
    let html = `
        <div class="test-summary ${summaryClass}">
            Test Suite: ${summaryText}
        </div>
    `;
    
    // Show filtered results
    for (let i = 0; i < filteredResults.length; i++) {
        const result = filteredResults[i];
        const resultClass = result.passed ? 'pass' : 'fail';
        const resultText = result.passed ? 'PASSED' : 'FAILED';
        
        // For failed tests, always show details
        // For passed tests (when shown), simplify the display
        const showFullDetails = !result.passed || !hidePassedTests;
        
        html += `
            <div class="result-item ${resultClass}" style="margin-top: 15px;">
                <h4>${result.testName} - ${resultText}</h4>
        `;
        
        if (result.error) {
            html += `<div class="error-message" style="margin: 5px 0; padding: 10px;">Error: ${result.error}</div>`;
        }
        
        if (showFullDetails) {
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
        } else {
            // For passed tests when hiding details, just show a summary
            html += `<div class="assertion-summary">✓ ${result.assertionResults.length} assertions passed</div>`;
        }
        
        html += '</div>';
    }
    
    resultsContent.innerHTML = html;
    resultsSection.style.display = 'block';
    
    // Show display options and update counts
    if (displayOptions && currentResults.length > 0) {
        const testCountDisplay = document.querySelector('.test-count-display');
        if (testCountDisplay) {
            testCountDisplay.style.display = 'flex';
        }
        if (visibleCountSpan) visibleCountSpan.textContent = visibleCount;
        if (totalCountSpan) totalCountSpan.textContent = currentTotalCount;
    }
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

function updateUIForTestState(running) {
    const runTestButton = document.getElementById('run-test-button');
    if (runTestButton) {
        runTestButton.disabled = running;
        runTestButton.querySelector('.ms-Button-label').textContent = running ? 'Running...' : 'Run';
    }
}