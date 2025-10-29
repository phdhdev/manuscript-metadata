/* global Word, Office */

Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
        document.getElementById("generateBtn").onclick = generateUniqueCode;
        document.getElementById("insertBtn").onclick = insertCode;
        
        // Update document stats on load
        updateDocumentStats();
    }
});

let currentCode = null;
let generatedCount = 0;

/**
 * Generate a random 6-digit code in format XXX-XXX
 */
function generateCode() {
    const part1 = Math.floor(Math.random() * 900) + 100; // 100-999
    const part2 = Math.floor(Math.random() * 900) + 100; // 100-999
    return `${part1}-${part2}`;
}

/**
 * Search the entire document for existing codes
 */
async function getAllCodesInDocument() {
    return await Word.run(async (context) => {
        const body = context.document.body;
        // Word wildcard pattern: [0-9] for digits
        const searchResults = body.search("[0-9]{3}-[0-9]{3}", { matchWildcards: true });
        
        searchResults.load("text");
        await context.sync();
        
        const codes = new Set();
        for (let i = 0; i < searchResults.items.length; i++) {
            codes.add(searchResults.items[i].text);
        }
        
        return codes;
    });
}

/**
 * Generate a unique code that doesn't exist in the document
 */
async function generateUniqueCode() {
    try {
        showStatus("Generating unique code...", "info");
        disableButtons(true);
        
        // Get all existing codes
        const existingCodes = await getAllCodesInDocument();
        
        // Generate a new code
        let newCode;
        let attempts = 0;
        const maxAttempts = 1000;
        
        do {
            newCode = generateCode();
            attempts++;
            
            if (attempts > maxAttempts) {
                throw new Error("Unable to generate unique code after 1000 attempts");
            }
        } while (existingCodes.has(newCode));
        
        // Display the code
        currentCode = newCode;
        document.getElementById("codeDisplay").textContent = newCode;
        document.getElementById("codeDisplay").classList.remove("empty");
        document.getElementById("insertBtn").disabled = false;
        
        generatedCount++;
        document.getElementById("generatedCount").textContent = generatedCount;
        
        showStatus(`✓ Generated unique code: ${newCode}`, "success");
        disableButtons(false);
        
    } catch (error) {
        showStatus(`Error: ${error.message}`, "error");
        disableButtons(false);
        console.error(error);
    }
}

/**
 * Insert the generated code at the cursor position with existing formatting
 */
async function insertCode() {
    if (!currentCode) {
        showStatus("Please generate a code first", "error");
        return;
    }
    
    try {
        showStatus("Inserting code...", "info");
        disableButtons(true);
        
        await Word.run(async (context) => {
            // Get the current selection/cursor position
            const selection = context.document.getSelection();
            
            // Load the font properties of the current selection
            selection.font.load(["name", "size", "color"]);
            await context.sync();
            
            // Store the current formatting
            const fontName = selection.font.name;
            const fontSize = selection.font.size;
            const fontColor = selection.font.color;
            
            // Insert the code at the cursor
            const insertedRange = selection.insertText(currentCode, Word.InsertLocation.replace);
            
            // Apply the original formatting to the inserted code
            insertedRange.font.name = fontName;
            insertedRange.font.size = fontSize;
            insertedRange.font.color = fontColor;
            
            await context.sync();
        });
        
        showStatus(`✓ Code ${currentCode} inserted successfully!`, "success");
        
        // Update stats
        await updateDocumentStats();
        
        // Reset for next generation
        currentCode = null;
        document.getElementById("codeDisplay").textContent = "Click Generate";
        document.getElementById("codeDisplay").classList.add("empty");
        document.getElementById("insertBtn").disabled = true;
        
        disableButtons(false);
        
    } catch (error) {
        showStatus(`Error inserting code: ${error.message}`, "error");
        disableButtons(false);
        console.error(error);
    }
}

/**
 * Update the document statistics
 */
async function updateDocumentStats() {
    try {
        const existingCodes = await getAllCodesInDocument();
        document.getElementById("totalCodes").textContent = existingCodes.size;
    } catch (error) {
        console.error("Error updating stats:", error);
    }
}

/**
 * Show status message
 */
function showStatus(message, type) {
    const statusDiv = document.getElementById("status");
    statusDiv.textContent = message;
    statusDiv.className = `status ${type}`;
    
    if (type === "success") {
        setTimeout(() => {
            statusDiv.style.display = "none";
        }, 3000);
    }
}

/**
 * Disable/enable buttons during operations
 */
function disableButtons(disabled) {
    document.getElementById("generateBtn").disabled = disabled;
    
    // Only disable insert button if we're disabling, or if there's no current code
    if (disabled) {
        document.getElementById("insertBtn").disabled = true;
    } else if (currentCode) {
        document.getElementById("insertBtn").disabled = false;
    }
}
