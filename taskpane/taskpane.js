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
        
        // Add real-time validation feedback
        testJsonInput.addEventListener('input', function(e) {
            validateJSONInput(e.target.value);
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
        const errorMessage = currentInputMethod === 'file' 
            ? 'Please select a JSON test file first'
            : 'Please paste or type JSON test content';
        showError(errorMessage);
        return;
    }
    
    setLoadingState(true, 'Loading...');
    clearResults();
    clearErrors();
    
    try {
        const testData = parseTestData(jsonText);
        prepareTests(testData);
        
        showSuccessFeedback(testJsonInput);
        setLoadingState(true, 'Running...');
        
        const testsToRun = currentTests || [currentTest];
        await executeTests(testsToRun, runTestButton);
        
        resetBorderColor(testJsonInput, 1000);
    } catch (error) {
        handleTestError(error, testJsonInput);
    } finally {
        setLoadingState(false, 'Run');
    }
}

function parseTestData(jsonText) {
    const cleanJsonText = jsonText
        .replace(/[\u201C\u201D\u201E\u201F\u2033\u2036]/g, '"')  // Replace smart quotes
        .replace(/[\u2018\u2019\u201A\u201B\u2032\u2035]/g, "'")  // Replace smart single quotes
        .replace(/\u00A0/g, " ")  // Replace non-breaking spaces
        .trim();
    
    return JSON.parse(cleanJsonText);
}

function prepareTests(testData) {
    if (Array.isArray(testData)) {
        currentTests = testData;
        currentTest = null;
        displayMultipleTestInfo(testData);
    } else {
        currentTests = null;
        currentTest = testData;
        displayTestInfo(testData);
    }
}

function setLoadingState(loading, buttonText) {
    const runTestButton = document.getElementById('run-test-button');
    if (runTestButton) {
        runTestButton.disabled = loading;
        runTestButton.querySelector('.ms-Button-label').textContent = buttonText;
    }
}

function showSuccessFeedback(testJsonInput) {
    testJsonInput.style.borderColor = '#107c10';
}

function resetBorderColor(testJsonInput, delay) {
    setTimeout(function() {
        testJsonInput.style.borderColor = '';
    }, delay);
}

function handleTestError(error, testJsonInput) {
    console.error('Error parsing or running test:', error);
    
    let errorMessage = `Failed to parse JSON: ${error.message}`;
    if (error instanceof SyntaxError) {
        errorMessage += '\n\nCommon issues:\n• Replace smart quotes ("") with regular quotes (")\n• Check for missing commas\n• Verify brackets and braces are balanced';
    }
    
    showError(errorMessage);
    testJsonInput.style.borderColor = '#d13438';
    resetBorderColor(testJsonInput, 2000);
}


function displayTestInfo(testData) {
    const testInfoDiv = document.getElementById('current-test-info');
    testInfoDiv.style.display = 'block';
    
    const testName = testData.name || 'Unnamed Test';
    const inputsHtml = createInputsHTML(testData.inputs);
    const assertionsHtml = createAssertionsHTML(testData.assertions);
    
    testInfoDiv.innerHTML = `
        <h3>${testName}</h3>
        ${inputsHtml}
        ${assertionsHtml}
    `;
}

function createInputsHTML(inputs) {
    if (!inputs || inputs.length === 0) {
        return '';
    }
    
    let html = '<p><strong>Inputs:</strong></p><ul>';
    for (const input of inputs) {
        const locationInfo = window.CellResolver.createLocationString(input);
        html += `<li>${locationInfo} = ${input.value}</li>`;
    }
    html += '</ul>';
    
    return html;
}

function createAssertionsHTML(assertions) {
    if (!assertions || assertions.length === 0) {
        return '';
    }
    
    let html = '<p><strong>Assertions:</strong></p><ul>';
    for (const assertion of assertions) {
        const tolerance = assertion.tolerance !== undefined ? ` (tolerance: ${assertion.tolerance})` : '';
        const locationInfo = window.CellResolver.createLocationString(assertion);
        html += `<li>${locationInfo} should equal ${assertion.equals}${tolerance}</li>`;
    }
    html += '</ul>';
    
    return html;
}

function displayMultipleTestInfo(tests) {
    const testInfoDiv = document.getElementById('current-test-info');
    
    let html = `<h3>Test Suite (${tests.length} test${tests.length > 1 ? 's' : ''})</h3>`;
    
    for (let i = 0; i < tests.length; i++) {
        const test = tests[i];
        html += `<div style="margin: 15px 0; padding: 10px; border-left: 3px solid #0078d4; background-color: #f3f2f1;">`;
        html += `<strong>${i + 1}. ${test.name || 'Unnamed Test'}</strong>`;
        
        if (test.inputs && test.inputs.length > 0) {
            html += '<p style="margin: 5px 0;"><small><strong>Inputs:</strong> ';
            const inputs = test.inputs.map(input => {
                const locationInfo = window.CellResolver.createLocationString(input);
                return `${locationInfo}=${input.value}`;
            }).join(', ');
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
    if (!currentResults || currentResults.length === 0) {
        return;
    }
    
    const hidePassedTests = shouldHidePassedTests();
    const filteredResults = getFilteredResults(hidePassedTests);
    
    const summaryHTML = createSummaryHTML();
    const resultsHTML = createResultsListHTML(filteredResults, hidePassedTests);
    
    displayResults(summaryHTML, resultsHTML);
    updateResultCounts(filteredResults.length, currentTotalCount);
}

function shouldHidePassedTests() {
    const hidePassedTestsCheckbox = document.getElementById('hide-passed-tests');
    return hidePassedTestsCheckbox && hidePassedTestsCheckbox.checked;
}

function getFilteredResults(hidePassedTests) {
    if (hidePassedTests) {
        return currentResults.filter(result => !result.passed);
    }
    return currentResults;
}

function createSummaryHTML() {
    const allPassed = currentPassedCount === currentTotalCount;
    const summaryClass = allPassed ? 'pass' : 'fail';
    const summaryText = allPassed ? 'ALL PASSED' : `${currentPassedCount}/${currentTotalCount} PASSED`;
    
    return `
        <div class="test-summary ${summaryClass}">
            Test Suite: ${summaryText}
        </div>
    `;
}

function createResultsListHTML(results, hidePassedTests) {
    let html = '';
    
    for (const result of results) {
        html += createSingleResultHTML(result, hidePassedTests);
    }
    
    return html;
}

function createSingleResultHTML(result, hidePassedTests) {
    const resultClass = result.passed ? 'pass' : 'fail';
    const resultText = result.passed ? 'PASSED' : 'FAILED';
    const showFullDetails = !result.passed || !hidePassedTests;
    
    let html = `
        <div class="result-item ${resultClass}" style="margin-top: 15px;">
            <h4>${result.testName} - ${resultText}</h4>
    `;
    
    if (result.error) {
        html += `<div class="error-message" style="margin: 5px 0; padding: 10px;">Error: ${result.error}</div>`;
    }
    
    if (showFullDetails) {
        html += createAssertionDetailsHTML(result.assertionResults);
    } else {
        html += `<div class="assertion-summary">✓ ${result.assertionResults.length} assertions passed</div>`;
    }
    
    html += '</div>';
    return html;
}

function createAssertionDetailsHTML(assertionResults) {
    let html = '';
    
    for (const assertionResult of assertionResults) {
        const assertionClass = assertionResult.passed ? 'pass' : 'fail';
        const detailsHtml = createAssertionDetailsText(assertionResult);
        
        html += `
            <div class="assertion ${assertionClass}">
                <strong>${assertionResult.cell}</strong>
                ${detailsHtml}
            </div>
        `;
    }
    
    return html;
}

function createAssertionDetailsText(assertionResult) {
    if (assertionResult.passed) {
        if (assertionResult.difference !== null) {
            return `<div class="assertion-details">Actual: ${assertionResult.actual}, Expected: ${assertionResult.expected}, Difference: ${assertionResult.difference}</div>`;
        } else {
            return `<div class="assertion-details">Actual: ${assertionResult.actual}, Expected: ${assertionResult.expected}</div>`;
        }
    } else {
        if (assertionResult.difference !== null) {
            return `<div class="assertion-details">Actual: ${assertionResult.actual}, Expected: ${assertionResult.expected}, Difference: ${assertionResult.difference} (tolerance: ${assertionResult.tolerance})</div>`;
        } else {
            return `<div class="assertion-details">Actual: ${assertionResult.actual}, Expected: ${assertionResult.expected}</div>`;
        }
    }
}

function displayResults(summaryHTML, resultsHTML) {
    const resultsSection = document.getElementById('results-section');
    const resultsContent = document.getElementById('results-content');
    const displayOptions = document.getElementById('display-options');
    
    resultsContent.innerHTML = summaryHTML + resultsHTML;
    resultsSection.style.display = 'block';
    
    // Show display options if we have results
    if (displayOptions && currentResults.length > 0) {
        const testCountDisplay = document.querySelector('.test-count-display');
        if (testCountDisplay) {
            testCountDisplay.style.display = 'flex';
        }
    }
}

function updateResultCounts(visibleCount, totalCount) {
    const visibleCountSpan = document.getElementById('visible-count');
    const totalCountSpan = document.getElementById('total-count');
    
    if (visibleCountSpan) visibleCountSpan.textContent = visibleCount;
    if (totalCountSpan) totalCountSpan.textContent = totalCount;
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

/**
 * Validate JSON input and provide visual feedback
 */
function validateJSONInput(jsonText) {
    const testJsonInput = document.getElementById('test-json-input');
    
    if (!jsonText.trim()) {
        testJsonInput.style.borderColor = '';
        return;
    }
    
    try {
        // Fix common quote issues before parsing
        const cleanJsonText = jsonText
            .replace(/[\u201C\u201D\u201E\u201F\u2033\u2036]/g, '"')  // Replace smart quotes
            .replace(/[\u2018\u2019\u201A\u201B\u2032\u2035]/g, "'")  // Replace smart single quotes
            .replace(/\u00A0/g, " ")  // Replace non-breaking spaces
            .trim();
        
        JSON.parse(cleanJsonText);
        testJsonInput.style.borderColor = '#107c10';  // Green for valid
    } catch (error) {
        testJsonInput.style.borderColor = '#d13438';  // Red for invalid
    }
}

function updateUIForTestState(running) {
    const runTestButton = document.getElementById('run-test-button');
    if (runTestButton) {
        runTestButton.disabled = running;
        runTestButton.querySelector('.ms-Button-label').textContent = running ? 'Running...' : 'Run';
    }
}