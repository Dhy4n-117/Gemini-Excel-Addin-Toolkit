/**
 * Excel AI Assistant — Main Taskpane Logic
 * Chat interface, action execution, and Office JS integration.
 */

import "./taskpane.css";
import { sendMessage, setApiKey, getApiKey, hasApiKey, testApiKey } from "../ai/gemini.js";
import { buildSystemPrompt, parseAIResponse } from "../ai/prompts.js";
import {
    getSpreadsheetContext,
    writeCell,
    writeRange,
    formatRange,
    sortRange,
    filterData,
    insertRow,
    deleteRows,
    removeDuplicates,
    trimWhitespace,
    getNextEmptyRow,
} from "../excel/operations.js";
import { createChart } from "../excel/charts.js";
import { formatMessage, formatTime, uid } from "../utils/helpers.js";

/* ===================================================================
   State
   =================================================================== */
let conversationHistory = [];
let isProcessing = false;
let initialized = false;

/* ===================================================================
   DOM Elements (assigned during initialize)
   =================================================================== */
let chatMessages, userInput, sendBtn, settingsBtn, settingsPanel,
    closeSettingsBtn, apiKeyInput, saveKeyBtn, toggleKeyVisibility,
    keyStatus, selectionIndicator, quickActions;

let domAssigned = false;

function assignDomElements() {
    if (domAssigned) return;
    
    chatMessages = document.getElementById("chatMessages");
    userInput = document.getElementById("userInput");
    sendBtn = document.getElementById("sendBtn");
    settingsBtn = document.getElementById("settingsBtn");
    settingsPanel = document.getElementById("settingsPanel");
    closeSettingsBtn = document.getElementById("closeSettingsBtn");
    apiKeyInput = document.getElementById("apiKeyInput");
    saveKeyBtn = document.getElementById("saveKeyBtn");
    toggleKeyVisibility = document.getElementById("toggleKeyVisibility");
    keyStatus = document.getElementById("keyStatus");
    selectionIndicator = document.getElementById("selectionIndicator");
    quickActions = document.getElementById("quickActions");
    
    domAssigned = true;
}

/* ===================================================================
   Initialize Office
   =================================================================== */
Office.onReady((info) => {
    assignDomElements();
    if (!initialized) {
        initialize();
    }
});

// Fallback: if Office.onReady doesn't fire, initialize on DOMContentLoaded
document.addEventListener("DOMContentLoaded", () => {
    assignDomElements();
    setTimeout(() => {
        if (!initialized) {
            console.warn("Office.onReady did not fire, initializing manually.");
            initialize();
        }
    }, 1500);
});

function initialize() {
    if (initialized) return;
    initialized = true;

    assignDomElements(); // Failsafe


    // Load saved API key
    if (hasApiKey()) {
        apiKeyInput.value = getApiKey();
    } else {
        // Show settings if no key configured
        settingsPanel.classList.remove("hidden");
    }

    // Global Event Delegation (prevents binding issues in Office Webview)
    document.addEventListener("click", (e) => {
        if (e.target.closest("#settingsBtn") || e.target.closest("#closeSettingsBtn")) {
            toggleSettings();
        } else if (e.target.closest("#saveKeyBtn")) {
            handleSaveKey();
        } else if (e.target.closest("#toggleKeyVisibility")) {
            toggleKeyVis();
        } else if (e.target.closest("#sendBtn")) {
            handleSend();
        } else {
            const quickBtn = e.target.closest(".quick-btn");
            if (quickBtn && !isProcessing) {
                userInput.value = quickBtn.dataset.action;
                handleSend();
            }
        }
    });

    // Input events
    if (userInput) {
        userInput.addEventListener("keydown", handleInputKeydown);
        userInput.addEventListener("input", autoResizeInput);
    }

    // Track selection changes
    startSelectionTracking();

    console.log("Excel AI Assistant initialized successfully.");
}

/* ===================================================================
   Selection Tracking
   =================================================================== */
function startSelectionTracking() {
    try {
        Excel.run(async (ctx) => {
            const sheet = ctx.workbook.worksheets.getActiveWorksheet();
            sheet.onSelectionChanged.add(handleSelectionChanged);
            await ctx.sync();
        });
    } catch (e) {
        console.warn("Selection tracking not available:", e);
    }

    // Initial selection state
    updateSelectionIndicator();
}

async function handleSelectionChanged() {
    updateSelectionIndicator();
}

async function updateSelectionIndicator() {
    try {
        await Excel.run(async (ctx) => {
            const range = ctx.workbook.getSelectedRange();
            range.load("address");
            await ctx.sync();

            selectionIndicator.textContent = `Selected: ${range.address}`;
            selectionIndicator.classList.add("active");
        });
    } catch (e) {
        selectionIndicator.textContent = "No selection";
        selectionIndicator.classList.remove("active");
    }
}

/* ===================================================================
   Settings
   =================================================================== */
function toggleSettings() {
    settingsPanel.classList.toggle("hidden");
}

function toggleKeyVis() {
    apiKeyInput.type = apiKeyInput.type === "password" ? "text" : "password";
}

async function handleSaveKey() {
    const key = apiKeyInput.value.trim();
    if (!key) {
        showKeyStatus("Please enter an API key.", "error");
        return;
    }

    saveKeyBtn.textContent = "Validating...";
    saveKeyBtn.disabled = true;

    setApiKey(key);

    const result = await testApiKey();
    if (result.valid) {
        showKeyStatus(`✓ API key valid! Using model: ${result.model || "gemini"}`, "success");
        setTimeout(() => {
            settingsPanel.classList.add("hidden");
        }, 1500);
    } else {
        showKeyStatus(`✗ ${result.error}`, "error");
    }

    saveKeyBtn.textContent = "Save Key";
    saveKeyBtn.disabled = false;
}

function showKeyStatus(msg, type) {
    keyStatus.textContent = msg;
    keyStatus.className = "key-status " + type;
}

/* ===================================================================
   Input Handling
   =================================================================== */
function handleInputKeydown(e) {
    if (e.key === "Enter" && !e.shiftKey) {
        e.preventDefault();
        handleSend();
    }
}

function autoResizeInput() {
    userInput.style.height = "auto";
    userInput.style.height = Math.min(userInput.scrollHeight, 120) + "px";
}

/* ===================================================================
   Send Message
   =================================================================== */
async function handleSend() {
    const text = userInput.value.trim();
    if (!text || isProcessing) return;

    if (!hasApiKey()) {
        settingsPanel.classList.remove("hidden");
        showKeyStatus("Please add your API key first.", "error");
        return;
    }

    isProcessing = true;
    sendBtn.disabled = true;
    userInput.value = "";
    userInput.style.height = "auto";

    // Add user message to chat
    addMessage("user", text);

    // Show typing indicator
    const typingId = showTyping();

    try {
        // Get current spreadsheet context
        let context = {};
        try {
            context = await getSpreadsheetContext();
        } catch (e) {
            console.warn("Could not get spreadsheet context:", e);
            context = { sheetName: "Sheet1", headers: [], selectedRange: null, selectedData: [], totalRows: 0, totalCols: 0 };
        }

        // Build system prompt with context
        const systemPrompt = buildSystemPrompt(context);

        // Add user message to conversation history
        conversationHistory.push({
            role: "user",
            parts: [{ text: text }],
        });

        // Send to Gemini
        const rawResponse = await sendMessage(conversationHistory, systemPrompt);

        // Parse the structured response
        const parsed = parseAIResponse(rawResponse);

        // Add AI response to conversation history
        conversationHistory.push({
            role: "model",
            parts: [{ text: rawResponse }],
        });

        // Remove typing indicator
        removeTyping(typingId);

        // Display the AI message with action buttons
        addAIMessage(parsed);
    } catch (error) {
        removeTyping(typingId);
        addMessage("error", `Error: ${error.message}`);
        console.error("AI Error:", error);
    }

    isProcessing = false;
    sendBtn.disabled = false;
    userInput.focus();
}

/* ===================================================================
   Chat UI
   =================================================================== */
function addMessage(type, text) {
    const messageDiv = document.createElement("div");
    messageDiv.className = `message ${type}-message`;
    messageDiv.id = `msg-${uid()}`;

    const avatarSvg =
        type === "user"
            ? `<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M20 21v-2a4 4 0 0 0-4-4H8a4 4 0 0 0-4 4v2"/><circle cx="12" cy="7" r="4"/></svg>`
            : `<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M12 2L15.09 8.26L22 9.27L17 14.14L18.18 21.02L12 17.77L5.82 21.02L7 14.14L2 9.27L8.91 8.26L12 2Z"/></svg>`;

    messageDiv.innerHTML = `
    <div class="message-avatar">${avatarSvg}</div>
    <div class="message-content">${formatMessage(text)}</div>
  `;

    chatMessages.appendChild(messageDiv);
    scrollToBottom();
}

function addAIMessage(parsed) {
    const messageDiv = document.createElement("div");
    messageDiv.className = "message ai-message";
    messageDiv.id = `msg-${uid()}`;

    const hasActions = parsed.actions && parsed.actions.length > 0 && parsed.actions[0].type !== "analysisOnly";

    let actionsHtml = "";
    if (hasActions) {
        const actionDescriptions = parsed.actions.map((a) => describeAction(a));
        actionsHtml = `
      <div class="action-buttons">
        <button class="action-btn apply-btn" data-actions='${JSON.stringify(parsed.actions).replace(/'/g, "&#39;")}'>
          ✨ Apply ${parsed.actions.length > 1 ? `(${parsed.actions.length} actions)` : ""}
        </button>
      </div>
    `;
    }

    messageDiv.innerHTML = `
    <div class="message-avatar">
      <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
        <path d="M12 2L15.09 8.26L22 9.27L17 14.14L18.18 21.02L12 17.77L5.82 21.02L7 14.14L2 9.27L8.91 8.26L12 2Z"/>
      </svg>
    </div>
    <div class="message-content">
      ${formatMessage(parsed.message)}
      ${actionsHtml}
    </div>
  `;

    // Attach apply handler
    const applyBtn = messageDiv.querySelector(".apply-btn");
    if (applyBtn) {
        applyBtn.addEventListener("click", async () => {
            if (applyBtn.classList.contains("applied")) return;

            applyBtn.textContent = "⏳ Applying...";
            applyBtn.disabled = true;

            try {
                const actions = JSON.parse(applyBtn.dataset.actions);
                await executeActions(actions);

                applyBtn.textContent = "✓ Applied";
                applyBtn.classList.add("applied");
            } catch (err) {
                applyBtn.textContent = "✗ Failed — click to retry";
                applyBtn.disabled = false;
                addMessage("error", `Failed to apply: ${err.message}`);
                console.error("Apply error:", err);
            }
        });
    }

    chatMessages.appendChild(messageDiv);
    scrollToBottom();
}

function describeAction(action) {
    switch (action.type) {
        case "writeCell":
            return `Write "${action.value}" to ${action.address}`;
        case "writeRange":
            return `Write data to ${action.startAddress}`;
        case "formatRange":
            return `Format ${action.address}`;
        case "createChart":
            return `Create ${action.chartType} chart`;
        case "sortRange":
            return `Sort data`;
        case "filterData":
            return `Filter and extract data`;
        case "insertRow":
            return `Insert new row`;
        case "deleteRows":
            return `Delete rows`;
        case "removeDuplicates":
            return `Remove duplicates`;
        case "trimWhitespace":
            return `Trim whitespace`;
        default:
            return action.type;
    }
}

function showTyping() {
    const id = `typing-${uid()}`;
    const div = document.createElement("div");
    div.className = "message ai-message";
    div.id = id;
    div.innerHTML = `
    <div class="message-avatar">
      <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
        <path d="M12 2L15.09 8.26L22 9.27L17 14.14L18.18 21.02L12 17.77L5.82 21.02L7 14.14L2 9.27L8.91 8.26L12 2Z"/>
      </svg>
    </div>
    <div class="message-content">
      <div class="typing-indicator">
        <div class="typing-dot"></div>
        <div class="typing-dot"></div>
        <div class="typing-dot"></div>
      </div>
    </div>
  `;
    chatMessages.appendChild(div);
    scrollToBottom();
    return id;
}

function removeTyping(id) {
    const el = document.getElementById(id);
    if (el) el.remove();
}

function scrollToBottom() {
    requestAnimationFrame(() => {
        chatMessages.scrollTop = chatMessages.scrollHeight;
    });
}

/* ===================================================================
   Execute AI Actions on the Spreadsheet
   =================================================================== */
async function executeActions(actions) {
    for (const action of actions) {
        switch (action.type) {
            case "writeCell":
                await writeCell(action.address, action.value);
                break;

            case "writeRange":
                await writeRange(action.startAddress, action.values);
                break;

            case "formatRange":
                await formatRange(action.address, action.format || {});
                break;

            case "createChart":
                await createChart(action.dataRange, action.chartType, action.title);
                break;

            case "sortRange":
                await sortRange(action.address, action.sortColumn || 0, action.ascending !== false);
                break;

            case "filterData":
                await filterData(action.sourceRange, action.column, action.operator, action.value, action.outputStart);
                break;

            case "insertRow": {
                let targetRow = action.afterRow;
                if (!targetRow) {
                    targetRow = await getNextEmptyRow();
                    targetRow = targetRow - 1; // insertRow adds after this row
                }
                await insertRow(targetRow, action.values);
                break;
            }

            case "deleteRows":
                await deleteRows(action.startRow, action.count || 1);
                break;

            case "removeDuplicates":
                await removeDuplicates(action.address);
                break;

            case "trimWhitespace":
                await trimWhitespace(action.address);
                break;

            case "analysisOnly":
                // No action needed
                break;

            default:
                console.warn("Unknown action type:", action.type);
        }
    }
}
