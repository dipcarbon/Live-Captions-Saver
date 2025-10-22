// --- Constants for DOM Elements and Data ---
const UI_ELEMENTS = {
    statusMessage: document.getElementById('status-message'),
    manualStartInfo: document.getElementById('manual-start-info'),
    copyButton: document.getElementById('copyButton'),
    copyDropdownButton: document.getElementById('copyDropdownButton'),
    copyOptions: document.getElementById('copyOptions'),
    saveButton: document.getElementById('saveButton'),
    saveDropdownButton: document.getElementById('saveDropdownButton'),
    saveOptions: document.getElementById('saveOptions'),
    speakerAliasList: document.getElementById('speaker-alias-list')
};


let currentDefaultFormat = 'md';
let extensionConfig = null;

// --- Error Handling ---
function safeExecute(fn, context = '', fallback = null) {
    try {
        return fn();
    } catch (error) {
        console.error(`[Teams Caption Saver] ${context}:`, error);
        return fallback;
    }
}

// --- Utility Functions ---
function escapeHtml(str) {
    const div = document.createElement('div');
    div.textContent = str;
    return div.innerHTML;
}

async function getActiveTeamsTab() {
    const tabs = await chrome.tabs.query({ active: true, currentWindow: true });
    const teamsTab = tabs.find(tab => tab.url?.startsWith("https://teams.microsoft.com"));
    return teamsTab || null;
}

async function formatTranscript(transcript, aliases, type = 'standard') {
    const processed = transcript.map(entry => ({
        ...entry,
        Name: aliases[entry.Name] || entry.Name
    }));

    if (type === 'ai') {
        const { aiInstructions: instructions } = await chrome.storage.sync.get('aiInstructions');
        const transcriptText = processed.map(entry => `[${entry.Time}] ${entry.Name}: ${entry.Text}`).join('\n\n');
        return instructions ? `${instructions}\n\n---\n\n${transcriptText}` : transcriptText;
    }

    return processed.map(entry => `[${entry.Time}] ${entry.Name}: ${entry.Text}`).join('\n');
}

// --- UI Update Functions ---
async function updateStatusUI({ capturing, captionCount, isInMeeting, attendeeCount }) {
    const { statusMessage } = UI_ELEMENTS;
    const { trackCaptions, trackAttendees } = await chrome.storage.sync.get(['trackCaptions', 'trackAttendees']);
    
    if (isInMeeting) {
        // In meeting - show appropriate status based on what's being tracked
        if (trackCaptions !== false && capturing) {
            let status = captionCount > 0 ? `Capturing! (${captionCount} lines recorded` : 'Capturing... (Waiting for speech';
            if (attendeeCount > 0) {
                status += `, ${attendeeCount} attendees`;
            }
            status += ')';
            statusMessage.textContent = status;
            statusMessage.style.color = captionCount > 0 ? '#28a745' : '#ffc107';
        } else if (trackCaptions === false && trackAttendees !== false && attendeeCount > 0) {
            // Only tracking attendees
            statusMessage.textContent = `Tracking attendees (${attendeeCount} participants)`;
            statusMessage.style.color = '#17a2b8';
        } else if (trackCaptions === false) {
            statusMessage.textContent = 'In a meeting (caption tracking disabled)';
            statusMessage.style.color = '#6c757d';
        } else {
            statusMessage.textContent = 'In a meeting, but captions are off.';
            statusMessage.style.color = '#dc3545';
        }
    } else {
        // Not in meeting - show saved data status
        let hasData = captionCount > 0 || attendeeCount > 0;
        if (hasData) {
            let status = 'Meeting ended. ';
            let parts = [];
            if (captionCount > 0) parts.push(`${captionCount} lines`);
            if (attendeeCount > 0) parts.push(`${attendeeCount} attendees`);
            status += parts.join(', ') + ' available.';
            statusMessage.textContent = status;
            statusMessage.style.color = '#17a2b8';
        } else {
            statusMessage.textContent = 'Not in a meeting.';
            statusMessage.style.color = '#6c757d';
        }
    }
}

function updateButtonStates(hasData) {
    const buttons = [
        UI_ELEMENTS.copyButton,
        UI_ELEMENTS.copyDropdownButton,
        UI_ELEMENTS.saveButton,
        UI_ELEMENTS.saveDropdownButton
    ].filter(Boolean);
    buttons.forEach(btn => btn.disabled = !hasData);
}

function updateSaveButtonText(format) {
    const formatLabels = {
        md: 'Save as Markdown',
        txt: 'Save as TXT'
    };
    UI_ELEMENTS.saveButton.textContent = formatLabels[format] || `Save as ${format.toUpperCase()}`;
}

async function renderSpeakerAliases(tab) {
    const { speakerAliasList } = UI_ELEMENTS;
    try {
        const response = await chrome.tabs.sendMessage(tab.id, { message: "get_unique_speakers" });
        if (!response?.speakers?.length) {
            speakerAliasList.innerHTML = '<p>No speakers detected yet.</p>';
            return;
        }

        const { speakerAliases = {} } = await chrome.storage.session.get('speakerAliases');
        speakerAliasList.innerHTML = ''; // Clear existing

        response.speakers.forEach(speaker => {
            const item = document.createElement('div');
            item.className = 'alias-item';
            item.innerHTML = `
                <label title="${escapeHtml(speaker)}">${escapeHtml(speaker)}</label>
                <input type="text" data-original-name="${escapeHtml(speaker)}" placeholder="Enter alias..." value="${escapeHtml(speakerAliases[speaker] || '')}">
            `;
            speakerAliasList.appendChild(item);
        });
    } catch (error) {
        console.error("Could not fetch or render speaker aliases:", error);
        speakerAliasList.innerHTML = '<p>Unable to load speakers. Please refresh the Teams tab and try again.</p>';
    }
}

// --- Settings Management ---
async function getExtensionConfig() {
    if (extensionConfig) {
        return extensionConfig;
    }

    const response = await fetch(chrome.runtime.getURL('config.json'));
    extensionConfig = await response.json();
    return extensionConfig;
}

async function loadSettings() {
    const config = await getExtensionConfig();
    const [{ autoEnableCaptions }, { defaultSaveFormat }] = await Promise.all([
        chrome.storage.sync.get('autoEnableCaptions'),
        chrome.storage.sync.get('defaultSaveFormat')
    ]);
    const allowedFormats = config.allowedSaveFormats || ['md', 'txt'];
    const shouldAutoEnable = (autoEnableCaptions ?? config.autoEnableCaptions) === true;
    const storedFormat = defaultSaveFormat && allowedFormats.includes(defaultSaveFormat) ? defaultSaveFormat : null;

    currentDefaultFormat = storedFormat || config.defaultSaveFormat || allowedFormats[0] || 'md';
    updateSaveButtonText(currentDefaultFormat);

    if (UI_ELEMENTS.manualStartInfo) {
        UI_ELEMENTS.manualStartInfo.style.display = shouldAutoEnable ? 'none' : 'block';
    }
}

// --- Event Handling ---
function setupEventListeners() {
    if (UI_ELEMENTS.speakerAliasList) {
        UI_ELEMENTS.speakerAliasList.addEventListener('change', async (e) => {
            if (e.target.tagName === 'INPUT') {
                const { originalName } = e.target.dataset;
                const newAlias = e.target.value.trim();
                const { speakerAliases = {} } = await chrome.storage.session.get('speakerAliases');
                speakerAliases[originalName] = newAlias;
                await chrome.storage.session.set({ speakerAliases });
            }
        });
    }

    UI_ELEMENTS.saveButton.addEventListener('click', async () => {
        const tab = await getActiveTeamsTab();
        if (tab) {
            chrome.tabs.sendMessage(tab.id, { message: "return_transcript", format: currentDefaultFormat });
        }
    });

    setupDropdown(UI_ELEMENTS.copyButton, UI_ELEMENTS.copyDropdownButton, UI_ELEMENTS.copyOptions, handleCopy);
    setupDropdown(null, UI_ELEMENTS.saveDropdownButton, UI_ELEMENTS.saveOptions, handleSave);

    document.addEventListener('click', () => {
        UI_ELEMENTS.copyOptions.style.display = 'none';
        UI_ELEMENTS.saveOptions.style.display = 'none';
    });
}

function setupDropdown(mainButton, dropdownButton, optionsContainer, actionHandler) {
    if (mainButton) {
        mainButton.addEventListener('click', () => optionsContainer.firstElementChild.click());
    }
    dropdownButton.addEventListener('click', (e) => {
        e.stopPropagation();
        optionsContainer.style.display = 'block';
    });
    optionsContainer.addEventListener('click', (e) => {
        e.preventDefault();
        e.stopPropagation();
        actionHandler(e.target);
        optionsContainer.style.display = 'none';
    });
}

async function handleCopy(target) {
    const copyType = target.dataset.copyType;
    if (!copyType) return;

    const tab = await getActiveTeamsTab();
    if (!tab) return;
    
    UI_ELEMENTS.statusMessage.textContent = "Preparing text to copy...";
    try {
        const response = await chrome.tabs.sendMessage(tab.id, { message: "get_transcript_for_copying" });
        if (response?.transcriptArray) {
            const { speakerAliases = {} } = await chrome.storage.session.get('speakerAliases');
            const formattedText = await formatTranscript(response.transcriptArray, speakerAliases, copyType);
            await navigator.clipboard.writeText(formattedText);
            UI_ELEMENTS.statusMessage.textContent = "Copied to clipboard!";
            UI_ELEMENTS.statusMessage.style.color = '#28a745';
        }
    } catch (error) {
        UI_ELEMENTS.statusMessage.textContent = "Copy failed.";
        UI_ELEMENTS.statusMessage.style.color = '#dc3545';
    }
}

async function handleSave(target) {
    const format = target.dataset.format;
    if (!format) return;

    const config = await getExtensionConfig();
    const allowedFormats = config.allowedSaveFormats || ['md', 'txt'];
    if (!allowedFormats.includes(format)) {
        return;
    }

    const tab = await getActiveTeamsTab();
    if (tab) {
        const formatLabels = {
            md: 'Markdown',
            txt: 'TXT'
        };
        UI_ELEMENTS.statusMessage.textContent = `Saving as ${formatLabels[format] || format.toUpperCase()}...`;
        chrome.tabs.sendMessage(tab.id, { message: "return_transcript", format });
    }
}

// --- Initialization ---
async function initializePopup() {
    await loadSettings();
    setupEventListeners();

    const tab = await getActiveTeamsTab();
    if (!tab) {
        UI_ELEMENTS.statusMessage.innerHTML = 'Please <a href="https://teams.microsoft.com" target="_blank">open a Teams tab</a> to use this extension.';
        UI_ELEMENTS.statusMessage.style.color = '#dc3545';
        return;
    }

    try {
        const status = await chrome.tabs.sendMessage(tab.id, { message: "get_status" });
        if (status) {
            await updateStatusUI(status);
            // Enable buttons if we have either captions or attendees
            const hasData = status.captionCount > 0 || (status.attendeeCount > 0 && status.isInMeeting === false);
            updateButtonStates(hasData);
            if (status.captionCount > 0) {
                renderSpeakerAliases(tab);
            }
        }
    } catch (error) {
        // This error is expected when content script isn't loaded yet
        if (error.message.includes("Could not establish connection")) {
            console.log("Content script not ready. This is normal if the Teams page was just opened.");
            UI_ELEMENTS.statusMessage.innerHTML = 'Please refresh your Teams tab (F5) to activate the extension.';
            UI_ELEMENTS.statusMessage.style.color = '#ffc107';
            
            // Try to inject the content script if it's not loaded
            try {
                await chrome.scripting.executeScript({
                    target: { tabId: tab.id },
                    files: ['content_script.js']
                });
                console.log("Content script injected successfully. Retrying connection...");
                // Retry after injection
                setTimeout(() => initializePopup(), 500);
            } catch (injectError) {
                console.log("Could not inject content script:", injectError.message);
                UI_ELEMENTS.statusMessage.textContent = "Please refresh your Teams tab to activate the extension.";
                UI_ELEMENTS.statusMessage.style.color = '#dc3545';
            }
        } else {
            console.error("Unexpected error:", error.message);
            UI_ELEMENTS.statusMessage.textContent = "Connection error. Please refresh your Teams tab and try again.";
            UI_ELEMENTS.statusMessage.style.color = '#dc3545';
        }
    }
}

// --- Keyboard Shortcuts ---
document.addEventListener('keydown', (e) => {
    // Ctrl/Cmd + S for save
    if ((e.ctrlKey || e.metaKey) && e.key === 's') {
        e.preventDefault();
        if (!UI_ELEMENTS.saveButton.disabled) {
            UI_ELEMENTS.saveButton.click();
        }
    }
    
    // Ctrl/Cmd + C for copy
    if ((e.ctrlKey || e.metaKey) && e.key === 'c' && !e.target.matches('input, textarea')) {
        e.preventDefault();
        if (!UI_ELEMENTS.copyButton.disabled) {
            UI_ELEMENTS.copyButton.click();
        }
    }
    
});

document.addEventListener('DOMContentLoaded', initializePopup);
