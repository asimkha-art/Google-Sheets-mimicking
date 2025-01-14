// Initialize necessary variables
let sheetDB = [];
let rows = 100;
let cols = 26;

// Initialize a 2D matrix for cells
for (let i = 0; i < rows; i++) {
    let sheetRow = [];
    for (let j = 0; j < cols; j++) {
        let cellProp = {
            value: "",
            formula: "",
            children: [],
        };
        sheetRow.push(cellProp);
    }
    sheetDB.push(sheetRow);
}

// Selecting all cells and attaching event listeners for manual data entry
for (let i = 0; i < rows; i++) {
    for (let j = 0; j < cols; j++) {
        let cell = document.querySelector(`.cell[rid="${i}"][cid="${j}"]`);
        
        // Event listener for data entry
        cell.addEventListener('blur', (e) => {
            let address = addressBar.value;
            let [activeCell, cellProp] = getCellAndCellProp(address);
            let enteredData = activeCell.innerText;

            if (enteredData === cellProp.value) return;

            cellProp.value = enteredData;
            removeChildFromParent(cellProp.formula);
            cellProp.formula = '';
            updateChildrenCells(address);
        });
    }
}

// Formula bar handling for formula input and execution
let formulaBar = document.querySelector('.formula-bar');
formulaBar.addEventListener('keydown', async (e) => {
    if (e.key === 'Enter') {
        let inputFormula = formulaBar.value.trim();
        let address = addressBar.value;
        let [cell, cellProp] = getCellAndCellProp(address);

        // If the formula is modified, remove the old relationship
        if (inputFormula !== cellProp.formula) {
            removeChildFromParent(cellProp.formula);
        }

        // Evaluate the formula and update the UI
        let evaluatedValue = evaluateFormula(inputFormula);
        if (evaluatedValue !== null) {
            setCellUIAndCellProp(evaluatedValue, inputFormula, address);
            addChildToParent(inputFormula);
            updateChildrenCells(address);
        } else {
            alert("Invalid formula or range!");
        }
    }
});

// Function to evaluate formulas with SUM, AVERAGE, MAX, MIN, COUNT
function evaluateFormula(formula) {
    formula = formula.trim();
    let [funcName, range] = formula.split(' ');
    funcName = funcName.toUpperCase();

    if (!range) {
        alert('Invalid formula syntax! Use: SUM A1:B2');
        return null;
    }

    let values = getValuesFromRange(range);
    switch (funcName) {
        case "SUM":
            return values.reduce((acc, val) => acc + val, 0);
        case "AVERAGE":
            return values.reduce((acc, val) => acc + val, 0) / values.length;
        case "MAX":
            return Math.max(...values);
        case "MIN":
            return Math.min(...values);
        case "COUNT":
            return values.filter(val => !isNaN(val)).length;
        default:
            alert('Unsupported formula. Use SUM, AVERAGE, MAX, MIN, COUNT.');
            return null;
    }
}

// Helper function to fetch values from a cell range
function getValuesFromRange(range) {
    let [startCell, endCell] = range.split(':');
    let [startRow, startCol] = decodeRidCidFromAddress(startCell);
    let [endRow, endCol] = decodeRidCidFromAddress(endCell);

    let values = [];
    for (let i = startRow; i <= endRow; i++) {
        for (let j = startCol; j <= endCol; j++) {
            let cell = document.querySelector(`.cell[rid="${i}"][cid="${j}"]`);
            let cellValue = Number(cell.innerText) || 0;
            values.push(cellValue);
        }
    }
    return values;
}

// Decodes cell address like A1 into row and column indices
function decodeRidCidFromAddress(address) {
    let col = address.charCodeAt(0) - 65;  // Convert 'A' to 0, 'B' to 1
    let row = parseInt(address.slice(1)) - 1; // Convert "1" to 0 (0-based index)
    return [row, col];
}

// Fetch the cell element and its properties from the database
function getCellAndCellProp(address) {
    let [row, col] = decodeRidCidFromAddress(address);
    let cell = document.querySelector(`.cell[rid="${row}"][cid="${col}"]`);
    let cellProp = sheetDB[row][col];
    return [cell, cellProp];
}

// Update cell UI and properties in the database
function setCellUIAndCellProp(value, formula, address) {
    let [cell, cellProp] = getCellAndCellProp(address);
    cell.innerText = value;
    cellProp.value = value;
    cellProp.formula = formula;
}

// Add a parent-child relationship for formula dependency tracking
function addChildToParent(formula) {
    let childAddress = addressBar.value;
    let encodedFormula = formula.split(' ');

    for (let token of encodedFormula) {
        let asciiValue = token.charCodeAt(0);
        if (asciiValue >= 65 && asciiValue <= 90) { // If a cell reference
            let [parentCell, parentCellProp] = getCellAndCellProp(token);
            parentCellProp.children.push(childAddress);
        }
    }
}

// Remove the parent-child relationship when formula is deleted
function removeChildFromParent(formula) {
    let childAddress = addressBar.value;
    let encodedFormula = formula.split(' ');

    for (let token of encodedFormula) {
        let asciiValue = token.charCodeAt(0);
        if (asciiValue >= 65 && asciiValue <= 90) {
            let [parentCell, parentCellProp] = getCellAndCellProp(token);
            let index = parentCellProp.children.indexOf(childAddress);
            if (index > -1) {
                parentCellProp.children.splice(index, 1);
            }
        }
    }
}

// Update all child cells when the parent value changes
function updateChildrenCells(parentAddress) {
    let [parentCell, parentCellProp] = getCellAndCellProp(parentAddress);
    let children = parentCellProp.children;

    for (let childAddress of children) {
        let [childCell, childCellProp] = getCellAndCellProp(childAddress);
        let evaluatedValue = evaluateFormula(childCellProp.formula);
        setCellUIAndCellProp(evaluatedValue, childCellProp.formula, childAddress);
        updateChildrenCells(childAddress);
    }
}
