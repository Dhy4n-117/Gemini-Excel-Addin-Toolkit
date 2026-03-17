/**
 * Shared Utilities
 */

/**
 * Debounce a function call
 */
export function debounce(fn, delay) {
    let timer;
    return (...args) => {
        clearTimeout(timer);
        timer = setTimeout(() => fn(...args), delay);
    };
}

/**
 * Sanitize HTML to prevent XSS in chat messages
 */
export function sanitizeHtml(str) {
    const div = document.createElement("div");
    div.textContent = str;
    return div.innerHTML;
}

/**
 * Format a timestamp for the chat
 */
export function formatTime(date = new Date()) {
    return date.toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" });
}

/**
 * Simple markdown-like formatting for AI messages
 * Supports: **bold**, *italic*, `code`, ```code blocks```
 */
export function formatMessage(text) {
    let html = sanitizeHtml(text);

    // Code blocks (triple backtick)
    html = html.replace(/```(\w*)\n?([\s\S]*?)```/g, '<pre class="code-block"><code>$2</code></pre>');

    // Inline code
    html = html.replace(/`([^`]+)`/g, '<code class="inline-code">$1</code>');

    // Bold
    html = html.replace(/\*\*([^*]+)\*\*/g, "<strong>$1</strong>");

    // Italic
    html = html.replace(/\*([^*]+)\*/g, "<em>$1</em>");

    // Line breaks
    html = html.replace(/\n/g, "<br>");

    return html;
}

/**
 * Generate a unique ID
 */
export function uid() {
    return Date.now().toString(36) + Math.random().toString(36).substring(2);
}
