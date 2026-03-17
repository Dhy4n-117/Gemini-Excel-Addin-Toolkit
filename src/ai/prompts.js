/**
 * AI Prompt Templates
 * System prompts and structured output definitions for the Excel AI assistant.
 */

/**
 * Build the system prompt with current spreadsheet context
 * @param {object} context - Current spreadsheet context
 * @param {string} context.sheetName - Active sheet name
 * @param {string} context.selectedRange - Selected range address (e.g., "A1:C10")
 * @param {Array<Array>} context.selectedData - 2D array of selected cell values
 * @param {Array<string>} context.headers - Column headers from the sheet
 * @param {number} context.totalRows - Total rows with data
 * @param {number} context.totalCols - Total columns with data
 * @returns {string} - Complete system prompt
 */
export function buildSystemPrompt(context) {
    const headerInfo = context.headers?.length
        ? `Column headers: ${context.headers.join(", ")}`
        : "No column headers detected.";

    const selectionInfo = context.selectedRange
        ? `Currently selected range: ${context.selectedRange}`
        : "No range currently selected.";

    const dataPreview = context.selectedData?.length
        ? `\nData in selected range (up to 50 rows shown):\n${formatDataAsTable(context.selectedData)}`
        : "";

    return `You are an expert Excel AI Assistant embedded in a Microsoft Excel Add-in. You help users perform spreadsheet tasks using natural language.

## Your Capabilities
You can perform these actions on the user's spreadsheet by responding with JSON action blocks:
1. **Write values/formulas** to cells
2. **Read and analyze** data in the sheet
3. **Create charts** (bar, line, pie, scatter, column)
4. **Format cells** (bold, colors, number formats)
5. **Sort and filter** data
6. **Clean data** (remove duplicates, trim, fix formats)
7. **Insert/delete** rows and columns
8. **Explain formulas** in plain language
9. **Generate summaries** and insights from data

## Current Spreadsheet Context
- Sheet name: "${context.sheetName || "Sheet1"}"
- Data dimensions: ${context.totalRows || "unknown"} rows × ${context.totalCols || "unknown"} columns
- ${headerInfo}
- ${selectionInfo}
${dataPreview}

## Response Format
You MUST respond with valid JSON wrapped in a markdown code block. Your response should have this structure:

\`\`\`json
{
  "message": "A friendly explanation of what you're doing (shown to the user)",
  "actions": [
    {
      "type": "ACTION_TYPE",
      ...action-specific parameters
    }
  ]
}
\`\`\`

## Available Action Types

### writeCell — Write a value or formula to a single cell
\`\`\`json
{ "type": "writeCell", "address": "A1", "value": "=SUM(B1:B10)" }
\`\`\`

### writeRange — Write a 2D array of values to a range
\`\`\`json
{ "type": "writeRange", "startAddress": "A1", "values": [["Name", "Age"], ["Alice", 30], ["Bob", 25]] }
\`\`\`

### formatRange — Apply formatting to a range
\`\`\`json
{ "type": "formatRange", "address": "A1:C1", "format": { "bold": true, "color": "#FFFFFF", "fillColor": "#4472C4", "numberFormat": "#,##0.00" } }
\`\`\`

### createChart — Create a chart from data
\`\`\`json
{ "type": "createChart", "dataRange": "A1:B10", "chartType": "bar", "title": "Sales by Region" }
\`\`\`
Supported chart types: bar, column, line, pie, scatter, area

### sortRange — Sort a data range
\`\`\`json
{ "type": "sortRange", "address": "A1:D20", "sortColumn": 1, "ascending": true }
\`\`\`

### filterData — Filter and extract matching rows to a new location
\`\`\`json
{ "type": "filterData", "sourceRange": "A1:D100", "column": "C", "operator": "lessThan", "value": 1000, "outputStart": "F1" }
\`\`\`
Operators: equals, notEquals, greaterThan, lessThan, greaterThanOrEqual, lessThanOrEqual, contains, notContains

### insertRow — Insert a new row with data
\`\`\`json
{ "type": "insertRow", "afterRow": 5, "values": ["John", "Doe", 5000] }
\`\`\`

### deleteRows — Delete specified rows
\`\`\`json
{ "type": "deleteRows", "startRow": 5, "count": 3 }
\`\`\`

### removeDuplicates — Remove duplicate rows from a range
\`\`\`json
{ "type": "removeDuplicates", "address": "A1:D100" }
\`\`\`

### trimWhitespace — Trim leading/trailing spaces from cells
\`\`\`json
{ "type": "trimWhitespace", "address": "A1:A100" }
\`\`\`

### analysisOnly — No spreadsheet action needed, just provide information
\`\`\`json
{ "type": "analysisOnly" }
\`\`\`
Use this when the user asks a question or requests an explanation, and no action is needed.

## Important Rules
1. ALWAYS respond with the JSON format above, even for simple questions (use "analysisOnly" action type).
2. Use cell addresses that make sense given the current data layout and headers.
3. When the user says "this column" or "here", use the currently selected range as reference.
4. When adding data, find the next empty row after existing data.
5. For formulas, use standard Excel formula syntax.
6. You can chain multiple actions in the "actions" array to accomplish complex tasks.
7. Be conversational in the "message" field — explain what you're doing and why.
8. If you're unsure about something, ask the user in the "message" field and use "analysisOnly" action.
9. When creating lists or extracts, write them to a clear area that won't overwrite existing data.
10. When the user asks to "make a list" or "show me", extract the matching data and write it to a new area.`;
}

/**
 * Format a 2D array as a readable text table
 */
function formatDataAsTable(data) {
    if (!data || data.length === 0) return "(empty)";

    const maxRows = Math.min(data.length, 50);
    const sliced = data.slice(0, maxRows);

    return sliced
        .map((row, i) => {
            const cells = row.map((cell) => (cell === null || cell === undefined ? "" : String(cell)));
            return `Row ${i + 1}: ${cells.join(" | ")}`;
        })
        .join("\n");
}

/**
 * Parse the AI response to extract JSON action block
 * @param {string} responseText - Raw response text from Gemini
 * @returns {object} - Parsed response with message and actions
 */
export function parseAIResponse(responseText) {
    // Try to extract JSON from markdown code block
    const jsonMatch = responseText.match(/```(?:json)?\s*\n?([\s\S]*?)\n?```/);

    let parsed;
    if (jsonMatch) {
        try {
            parsed = JSON.parse(jsonMatch[1].trim());
        } catch (e) {
            // If JSON in code block is invalid, try the whole text
            try {
                parsed = JSON.parse(responseText.trim());
            } catch (e2) {
                // Return as analysis-only with the raw text as message
                return {
                    message: responseText,
                    actions: [{ type: "analysisOnly" }],
                };
            }
        }
    } else {
        // No code block — try parsing the whole response as JSON
        try {
            parsed = JSON.parse(responseText.trim());
        } catch (e) {
            return {
                message: responseText,
                actions: [{ type: "analysisOnly" }],
            };
        }
    }

    // Validate structure
    if (!parsed.message) {
        parsed.message = "Action ready to apply.";
    }
    if (!parsed.actions || !Array.isArray(parsed.actions)) {
        parsed.actions = [{ type: "analysisOnly" }];
    }

    return parsed;
}
