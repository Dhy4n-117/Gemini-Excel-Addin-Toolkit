/**
 * Gemini API Client
 * Handles communication with Google Gemini API for AI-powered Excel assistance.
 */

const GEMINI_API_BASE = "https://generativelanguage.googleapis.com/v1beta/models";
const MODEL_PREFERENCES = ["gemini-2.5-flash", "gemini-2.5", "gemini-2.0-flash", "gemini-1.5-flash"];
let activeModel = MODEL_PREFERENCES[0];

const STORAGE_KEY = "excel_ai_gemini_key";

/**
 * Store the API key in localStorage
 */
export function setApiKey(key) {
    localStorage.setItem(STORAGE_KEY, key.trim());
}

/**
 * Retrieve the API key from localStorage
 */
export function getApiKey() {
    return localStorage.getItem(STORAGE_KEY) || "";
}

/**
 * Check if an API key is configured
 */
export function hasApiKey() {
    return !!getApiKey();
}

/**
 * Remove the stored API key
 */
export function clearApiKey() {
    localStorage.removeItem(STORAGE_KEY);
}

/**
 * Send a message to Gemini and get a response
 * @param {Array} conversationHistory - Array of { role, parts } messages
 * @param {string} systemPrompt - System instruction for the model
 * @returns {Promise<string>} - The model's response text
 */
export async function sendMessage(conversationHistory, systemPrompt) {
    const apiKey = getApiKey();
    if (!apiKey) {
        throw new Error("No API key configured. Please add your Gemini API key in Settings.");
    }

    const MAX_RETRIES = 3;
    const RETRY_DELAYS = [3000, 6000, 12000]; // 3s, 6s, 12s

    for (let attempt = 0; attempt <= MAX_RETRIES; attempt++) {
        const url = `${GEMINI_API_BASE}/${activeModel}:generateContent?key=${apiKey}`;

        const requestBody = {
            contents: conversationHistory,
            systemInstruction: {
                parts: [{ text: systemPrompt }],
            },
            generationConfig: {
                temperature: 0.7,
                topP: 0.95,
                topK: 40,
                maxOutputTokens: 4096,
            },
        };

        try {
            const response = await fetch(url, {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify(requestBody),
            });

            if (!response.ok) {
                const errorData = await response.json().catch(() => ({}));
                const errorMessage = errorData?.error?.message || `API error (status ${response.status})`;

                if (response.status === 429) {
                    // Rate limited — retry with delay
                    if (attempt < MAX_RETRIES) {
                        const delay = RETRY_DELAYS[attempt];
                        console.log(`Rate limited (${errorMessage}). Retrying in ${delay / 1000}s (attempt ${attempt + 1}/${MAX_RETRIES})...`);
                        await new Promise((resolve) => setTimeout(resolve, delay));
                        continue;
                    } else {
                        throw new Error(`Google API Quota Error: ${errorMessage}. Try creating a new project with billing enabled, or wait 24h.`);
                    }
                } else if (response.status === 400) {
                    throw new Error(`Request error: ${errorMessage}`);
                } else if (response.status === 403) {
                    throw new Error("API key rejected. Please generate a new key from Google AI Studio.");
                } else if (response.status === 404) {
                    throw new Error(`Model "${activeModel}" not found. ${errorMessage}`);
                } else {
                    throw new Error(`API error ${response.status}: ${errorMessage}`);
                }
            }

            const data = await response.json();

            if (!data.candidates || data.candidates.length === 0) {
                throw new Error("No response generated. The AI may have filtered the response.");
            }

            const candidate = data.candidates[0];
            if (candidate.finishReason === "SAFETY") {
                throw new Error("Response was filtered for safety. Please rephrase your request.");
            }

            const text = candidate.content?.parts?.[0]?.text;
            if (!text) {
                throw new Error("Empty response from AI.");
            }

            return text;
        } catch (e) {
            // If it's a network error (not our thrown errors), retry
            if (e.name === "TypeError" && attempt < MAX_RETRIES) {
                const delay = RETRY_DELAYS[attempt];
                console.log(`Network error. Retrying in ${delay / 1000}s...`);
                await new Promise((resolve) => setTimeout(resolve, delay));
                continue;
            }
            throw e;
        }
    }
}

/**
 * Test the API key and find a working model
 * @returns {Promise<{valid: boolean, error: string}>}
 */
export async function testApiKey() {
    const apiKey = getApiKey();
    if (!apiKey) {
        return { valid: false, error: "No API key provided." };
    }

    // First, test if the key is valid by listing available models
    try {
        const listUrl = `https://generativelanguage.googleapis.com/v1beta/models?key=${apiKey}`;
        const listResponse = await fetch(listUrl);

        if (!listResponse.ok) {
            if (listResponse.status === 400 || listResponse.status === 403) {
                return { valid: false, error: "Invalid API key. Please check and try again." };
            }
            const errData = await listResponse.json().catch(() => ({}));
            return { valid: false, error: errData?.error?.message || `API error (${listResponse.status})` };
        }

        const listData = await listResponse.json();
        const availableModels = (listData.models || []).map((m) => m.name.replace("models/", ""));

        // Find the best available model from our preferences
        let foundModel = null;
        for (const pref of MODEL_PREFERENCES) {
            if (availableModels.some((m) => m === pref || m.startsWith(pref))) {
                foundModel = pref;
                break;
            }
        }

        if (foundModel) {
            activeModel = foundModel;
        } else if (availableModels.length > 0) {
            // Fallback to the first available flash model if our preferences fail
            activeModel = availableModels.find(m => m.includes("flash")) || availableModels[0];
        }

        return { valid: true, error: "", model: activeModel };
    } catch (e) {
        // Network error
        return { valid: false, error: `Connection failed: ${e.message}. Check your internet connection.` };
    }
}

