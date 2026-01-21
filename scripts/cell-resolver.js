/**
 * Cell Resolver - Functions to handle relative cell references
 * Supports text-based cell finding, offset calculations, and intersections
 */

/**
 * Find a cell containing specific text in a worksheet
 * @param {Excel.Worksheet} worksheet - The worksheet to search
 * @param {string} searchText - Exact text to search for
 * @param {Excel.RequestContext} context - Office.js context
 * @returns {Promise<string>} - Full cell address (e.g., "Sheet1!A1")
 */
async function findCellByText(worksheet, searchText, context) {
    const usedRange = worksheet.getUsedRange();
    usedRange.load("values");
    worksheet.load("name");
    await context.sync();

    const matches = [];
    
    // Search through all cells in the used range
    for (let row = 0; row < usedRange.values.length; row++) {
        for (let col = 0; col < usedRange.values[row].length; col++) {
            const cellValue = usedRange.values[row][col];
            if (cellValue && cellValue.toString() === searchText) {
                matches.push({ row, col });
            }
        }
    }

    if (matches.length === 0) {
        throw new Error(`No cell found containing text: "${searchText}"`);
    }

    if (matches.length > 1) {
        const matchAddresses = matches.map(m => `${columnNumberToLetter(m.col + 1)}${m.row + 1}`);
        throw new Error(`Multiple cells found containing text: "${searchText}". Found: ${matchAddresses.join(', ')}`);
    }

    const match = matches[0];
    const cellAddress = `${columnNumberToLetter(match.col + 1)}${match.row + 1}`;
    return `${worksheet.name}!${cellAddress}`;
}

/**
 * Get a cell at a specific offset from a base cell
 * @param {string} baseCellAddress - Base cell address (e.g., "Sheet1!A1")
 * @param {number} colOffset - Column offset (positive = right, negative = left)
 * @param {number} rowOffset - Row offset (positive = down, negative = up)
 * @returns {string} - Offset cell address
 */
function getOffsetCell(baseCellAddress, colOffset, rowOffset) {
    const parsed = parseCellAddress(baseCellAddress);
    
    // Convert column letter to number
    const baseCol = columnLetterToNumber(parsed.cellAddress.match(/[A-Z]+/)[0]);
    const baseRow = parseInt(parsed.cellAddress.match(/\d+/)[0]);
    
    // Calculate offset
    const newCol = baseCol + colOffset;
    const newRow = baseRow + rowOffset;
    
    // Validate new position
    if (newCol < 1 || newRow < 1) {
        throw new Error(`Invalid offset from ${baseCellAddress}: column ${newCol}, row ${newRow}`);
    }
    
    const newCellAddress = `${columnNumberToLetter(newCol)}${newRow}`;
    return `${parsed.worksheetName}!${newCellAddress}`;
}

/**
 * Find the intersection of a column header and row header
 * @param {Excel.Worksheet} worksheet - The worksheet to search
 * @param {string} colHeaderText - Text for column header
 * @param {string} rowHeaderText - Text for row header
 * @param {Excel.RequestContext} context - Office.js context
 * @returns {Promise<string>} - Intersection cell address
 */
async function findIntersection(worksheet, colHeaderText, rowHeaderText, context) {
    const usedRange = worksheet.getUsedRange();
    usedRange.load("values");
    worksheet.load("name");
    await context.sync();

    let colMatch = null;
    let rowMatch = null;
    const colMatches = [];
    const rowMatches = [];

    // Find column header (text in any row)
    for (let row = 0; row < usedRange.values.length; row++) {
        for (let col = 0; col < usedRange.values[row].length; col++) {
            const cellValue = usedRange.values[row][col];
            if (cellValue && cellValue.toString() === colHeaderText) {
                colMatches.push({ row, col });
            }
        }
    }

    // Find row header (text in any column)
    for (let row = 0; row < usedRange.values.length; row++) {
        for (let col = 0; col < usedRange.values[row].length; col++) {
            const cellValue = usedRange.values[row][col];
            if (cellValue && cellValue.toString() === rowHeaderText) {
                rowMatches.push({ row, col });
            }
        }
    }

    // Validate matches
    if (colMatches.length === 0) {
        throw new Error(`No cell found containing column header text: "${colHeaderText}"`);
    }
    if (rowMatches.length === 0) {
        throw new Error(`No cell found containing row header text: "${rowHeaderText}"`);
    }
    if (colMatches.length > 1) {
        throw new Error(`Multiple cells found containing column header text: "${colHeaderText}"`);
    }
    if (rowMatches.length > 1) {
        throw new Error(`Multiple cells found containing row header text: "${rowHeaderText}"`);
    }

    // Intersection is at column from colMatch and row from rowMatch
    const intersectionCol = colMatches[0].col;
    const intersectionRow = rowMatches[0].row;
    
    const cellAddress = `${columnNumberToLetter(intersectionCol + 1)}${intersectionRow + 1}`;
    return cellAddress;
}

/**
 * Resolve a relative reference to an absolute cell address
 * @param {Object} relativeTo - Relative reference object
 * @param {Excel.Workbook} workbook - Excel workbook
 * @param {Excel.RequestContext} context - Office.js context
 * @returns {Promise<string>} - Absolute cell address
 */
async function resolveRelativeReference(relativeTo, workbook, context) {
    const { sheet } = relativeTo;
    
    // Get worksheet
    const worksheet = workbook.worksheets.getItem(sheet);
    
    // Handle offset reference
    if (relativeTo.referenceCell && relativeTo.colOffset !== undefined && relativeTo.rowOffset !== undefined) {
        const baseCellAddress = await findCellByText(worksheet, relativeTo.referenceCell, context);
        const fullBaseAddress = `${sheet}!${baseCellAddress}`;
        return getOffsetCell(fullBaseAddress, relativeTo.colOffset, relativeTo.rowOffset);
    }
    
    // Handle intersection reference
    if (relativeTo.referenceColCell && relativeTo.referenceRowCell) {
        const cellAddress = await findIntersection(worksheet, relativeTo.referenceColCell, relativeTo.referenceRowCell, context);
        return `${sheet}!${cellAddress}`;
    }
    
    throw new Error('Invalid relative reference: must contain either (referenceCell, colOffset, rowOffset) or (referenceColCell, referenceRowCell)');
}

/**
 * Convert column letter to number (A=1, B=2, Z=26, AA=27, etc.)
 * @param {string} column - Column letter(s)
 * @returns {number} - Column number
 */
function columnLetterToNumber(column) {
    let result = 0;
    for (let i = 0; i < column.length; i++) {
        result = result * 26 + (column.charCodeAt(i) - 'A'.charCodeAt(0) + 1);
    }
    return result;
}

/**
 * Convert column number to letter (1=A, 2=B, 26=Z, 27=AA, 28=AB, etc.)
 * @param {number} column - Column number
 * @returns {string} - Column letter(s)
 */
function columnNumberToLetter(column) {
    let result = '';
    while (column > 0) {
        column--;
        result = String.fromCharCode('A'.charCodeAt(0) + (column % 26)) + result;
        column = Math.floor(column / 26);
    }
    return result;
}

/**
 * Parse a cell address into worksheet name and cell address
 * @param {string} fullAddress - Full address like "Sheet1!A1"
 * @returns {Object} - {worksheetName, cellAddress}
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
 * Main resolver function that handles both direct and relative cell references
 * @param {Object|string} reference - Either a cell address string or input/assertion object
 * @param {Excel.Workbook} workbook - Excel workbook
 * @param {Excel.RequestContext} context - Office.js context
 * @returns {Promise<string>} - Absolute cell address
 */
async function resolveCell(reference, workbook, context) {
    // If it's a string, it's a direct cell reference
    if (typeof reference === 'string') {
        return reference;
    }
    
    // If it's an object with cell property, return the direct cell reference
    if (reference.cell) {
        return reference.cell;
    }
    
    // If it's an object with relativeTo property, resolve the relative reference
    if (reference.relativeTo) {
        return await resolveRelativeReference(reference.relativeTo, workbook, context);
    }
    
    throw new Error('Invalid reference: must be a string, object with cell property, or object with relativeTo property');
}

/**
 * Create a user-friendly location string for display purposes
 * @param {Object|string} reference - Either a cell address string or input/assertion object
 * @returns {string} - Human-readable location description
 */
function createLocationString(reference) {
    // If it's a string, return as-is
    if (typeof reference === 'string') {
        return reference;
    }
    
    // If it's an input or assertion object
    if (reference.cell) {
        return reference.cell;
    }
    
    if (reference.relativeTo) {
        if (reference.relativeTo.referenceCell) {
            return `${reference.relativeTo.referenceCell}+(${reference.relativeTo.colOffset},${reference.relativeTo.rowOffset})`;
        } else {
            return `${reference.relativeTo.referenceColCell}×${reference.relativeTo.referenceRowCell}`;
        }
    }
    
    return 'Unknown location';
}

// Export functions globally for Office.js add-in
window.CellResolver = {
    resolveCell,
    createLocationString,
    parseCellAddress
};

// Also support Node.js/CommonJS for testing
if (typeof module !== 'undefined' && module.exports) {
    module.exports = {
        resolveCell,
        createLocationString,
        parseCellAddress,
        // Keep internal functions available for testing
        _findCellByText: findCellByText,
        _getOffsetCell: getOffsetCell,
        _findIntersection: findIntersection,
        _columnLetterToNumber: columnLetterToNumber,
        _columnNumberToLetter: columnNumberToLetter
    };
}