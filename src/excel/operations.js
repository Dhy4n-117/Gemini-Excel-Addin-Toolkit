/**
 * Excel Operations
 * Read/write cells, format ranges, sort, filter, and manipulate data via Office JS API.
 */

/**
 * Get context about the current spreadsheet state
 * @returns {Promise<object>} - Sheet context for the AI
 */
export async function getSpreadsheetContext() {
    return Excel.run(async (ctx) => {
        const sheet = ctx.workbook.worksheets.getActiveWorksheet();
        sheet.load("name");

        const usedRange = sheet.getUsedRangeOrNullObject();
        usedRange.load(["address", "rowCount", "columnCount", "values"]);

        const selectedRange = ctx.workbook.getSelectedRange();
        selectedRange.load(["address", "values"]);

        await ctx.sync();

        const context = {
            sheetName: sheet.name,
            selectedRange: selectedRange.address || null,
            selectedData: selectedRange.values || [],
            headers: [],
            totalRows: 0,
            totalCols: 0,
        };

        if (!usedRange.isNullObject) {
            context.totalRows = usedRange.rowCount;
            context.totalCols = usedRange.columnCount;

            // Extract headers (first row)
            if (usedRange.values && usedRange.values.length > 0) {
                context.headers = usedRange.values[0].map((h) => (h !== null && h !== undefined ? String(h) : ""));
            }
        }

        return context;
    });
}

/**
 * Write a value or formula to a single cell
 */
export async function writeCell(address, value) {
    return Excel.run(async (ctx) => {
        const sheet = ctx.workbook.worksheets.getActiveWorksheet();
        const range = sheet.getRange(address);

        if (typeof value === "string" && value.startsWith("=")) {
            range.formulas = [[value]];
        } else {
            range.values = [[value]];
        }

        await ctx.sync();
    });
}

/**
 * Write a 2D array of values starting at a given address
 */
export async function writeRange(startAddress, values) {
    return Excel.run(async (ctx) => {
        const sheet = ctx.workbook.worksheets.getActiveWorksheet();
        const startRange = sheet.getRange(startAddress);

        const rows = values.length;
        const cols = values[0]?.length || 1;

        const targetRange = startRange.getResizedRange(rows - 1, cols - 1);

        // Separate formulas from regular values
        const hasFormulas = values.some((row) => row.some((cell) => typeof cell === "string" && cell.startsWith("=")));

        if (hasFormulas) {
            targetRange.formulas = values;
        } else {
            targetRange.values = values;
        }

        await ctx.sync();
    });
}

/**
 * Format a range of cells
 */
export async function formatRange(address, format) {
    return Excel.run(async (ctx) => {
        const sheet = ctx.workbook.worksheets.getActiveWorksheet();
        const range = sheet.getRange(address);

        if (format.bold !== undefined) range.format.font.bold = format.bold;
        if (format.italic !== undefined) range.format.font.italic = format.italic;
        if (format.color) range.format.font.color = format.color;
        if (format.fontSize) range.format.font.size = format.fontSize;
        if (format.fillColor) range.format.fill.color = format.fillColor;
        if (format.numberFormat) range.numberFormat = [[format.numberFormat]];
        if (format.horizontalAlignment) range.format.horizontalAlignment = format.horizontalAlignment;
        if (format.borderStyle) {
            const borders = range.format.borders;
            borders.getItem("EdgeBottom").style = format.borderStyle;
            borders.getItem("EdgeTop").style = format.borderStyle;
            borders.getItem("EdgeLeft").style = format.borderStyle;
            borders.getItem("EdgeRight").style = format.borderStyle;
        }

        await ctx.sync();
    });
}

/**
 * Sort a data range by a column
 */
export async function sortRange(address, sortColumnIndex, ascending = true) {
    return Excel.run(async (ctx) => {
        const sheet = ctx.workbook.worksheets.getActiveWorksheet();
        const range = sheet.getRange(address);

        range.sort.apply([
            {
                key: sortColumnIndex,
                ascending: ascending,
            },
        ]);

        await ctx.sync();
    });
}

/**
 * Filter data and write matching rows to a new location
 */
export async function filterData(sourceRange, columnLetter, operator, filterValue, outputStart) {
    return Excel.run(async (ctx) => {
        const sheet = ctx.workbook.worksheets.getActiveWorksheet();
        const range = sheet.getRange(sourceRange);
        range.load("values");

        await ctx.sync();

        const data = range.values;
        if (!data || data.length === 0) return;

        // Find column index from letter
        const colIndex = columnLetter.toUpperCase().charCodeAt(0) - 65;
        const headers = data[0];
        const filteredRows = [headers];

        for (let i = 1; i < data.length; i++) {
            const cellValue = data[i][colIndex];
            let matches = false;

            switch (operator) {
                case "equals":
                    matches = cellValue == filterValue;
                    break;
                case "notEquals":
                    matches = cellValue != filterValue;
                    break;
                case "greaterThan":
                    matches = Number(cellValue) > Number(filterValue);
                    break;
                case "lessThan":
                    matches = Number(cellValue) < Number(filterValue);
                    break;
                case "greaterThanOrEqual":
                    matches = Number(cellValue) >= Number(filterValue);
                    break;
                case "lessThanOrEqual":
                    matches = Number(cellValue) <= Number(filterValue);
                    break;
                case "contains":
                    matches = String(cellValue).toLowerCase().includes(String(filterValue).toLowerCase());
                    break;
                case "notContains":
                    matches = !String(cellValue).toLowerCase().includes(String(filterValue).toLowerCase());
                    break;
            }

            if (matches) {
                filteredRows.push(data[i]);
            }
        }

        // Write filtered results to output location
        if (filteredRows.length > 0) {
            const outputRange = sheet.getRange(outputStart);
            const targetRange = outputRange.getResizedRange(filteredRows.length - 1, filteredRows[0].length - 1);
            targetRange.values = filteredRows;
        }

        await ctx.sync();
    });
}

/**
 * Insert a new row with data after a specified row number
 */
export async function insertRow(afterRow, values) {
    return Excel.run(async (ctx) => {
        const sheet = ctx.workbook.worksheets.getActiveWorksheet();

        // Insert a new row
        const insertRange = sheet.getRange(`${afterRow + 1}:${afterRow + 1}`);
        insertRange.insert(Excel.InsertShiftDirection.down);

        // Write the values to the new row
        const cols = values.length;
        const startCell = `A${afterRow + 1}`;
        const dataRange = sheet.getRange(startCell).getResizedRange(0, cols - 1);
        dataRange.values = [values];

        await ctx.sync();
    });
}

/**
 * Delete specified number of rows starting from a given row
 */
export async function deleteRows(startRow, count) {
    return Excel.run(async (ctx) => {
        const sheet = ctx.workbook.worksheets.getActiveWorksheet();
        const range = sheet.getRange(`${startRow}:${startRow + count - 1}`);
        range.delete(Excel.DeleteShiftDirection.up);

        await ctx.sync();
    });
}

/**
 * Remove duplicate rows from a range
 */
export async function removeDuplicates(address) {
    return Excel.run(async (ctx) => {
        const sheet = ctx.workbook.worksheets.getActiveWorksheet();
        const range = sheet.getRange(address);
        range.load(["values", "rowCount", "columnCount"]);

        await ctx.sync();

        const data = range.values;
        if (!data || data.length <= 1) return;

        const seen = new Set();
        const uniqueRows = [data[0]]; // Keep header

        for (let i = 1; i < data.length; i++) {
            const key = JSON.stringify(data[i]);
            if (!seen.has(key)) {
                seen.add(key);
                uniqueRows.push(data[i]);
            }
        }

        // Clear original range and write unique data
        range.clear();

        const outputRange = sheet.getRange(address.split(":")[0]);
        const targetRange = outputRange.getResizedRange(uniqueRows.length - 1, uniqueRows[0].length - 1);
        targetRange.values = uniqueRows;

        await ctx.sync();
    });
}

/**
 * Trim whitespace from cells in a range
 */
export async function trimWhitespace(address) {
    return Excel.run(async (ctx) => {
        const sheet = ctx.workbook.worksheets.getActiveWorksheet();
        const range = sheet.getRange(address);
        range.load("values");

        await ctx.sync();

        const values = range.values.map((row) =>
            row.map((cell) => (typeof cell === "string" ? cell.trim() : cell))
        );

        range.values = values;
        await ctx.sync();
    });
}

/**
 * Get the address of the first empty row in the used range
 */
export async function getNextEmptyRow() {
    return Excel.run(async (ctx) => {
        const sheet = ctx.workbook.worksheets.getActiveWorksheet();
        const usedRange = sheet.getUsedRangeOrNullObject();
        usedRange.load("rowCount");

        await ctx.sync();

        if (usedRange.isNullObject) {
            return 1;
        }

        return usedRange.rowCount + 1;
    });
}
