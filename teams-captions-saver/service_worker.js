// --- Utility Functions ---
function getSanitizedMeetingName(fullTitle) {
    if (!fullTitle) return "Meeting";
    const parts = fullTitle.split('|');
    // Handles titles like "Meeting Name | Microsoft Teams" or "Location | Meeting | Teams"
    const meetingName = parts.length > 2 ? parts[1] : parts[0];
    const cleanedName = meetingName.replace('Microsoft Teams', '').trim();
    // Replace characters forbidden in filenames
    return cleanedName.replace(/[<>:"/\\|?*\x00-\x1F]/g, '_') || "Meeting";
}


function applyAliasesToTranscript(transcriptArray, aliases = {}) {
    if (Object.keys(aliases).length === 0) {
        return transcriptArray;
    }
    return transcriptArray.map(entry => {
        const newName = aliases[entry.Name]?.trim();
        return {
            ...entry,
            Name: newName || entry.Name
        };
    });
}

function applyAliasesToAttendeeReport(attendeeReport, aliases = {}) {
    if (!attendeeReport || Object.keys(aliases).length === 0) {
        return attendeeReport;
    }

    // Create a new report with aliased names
    const aliasedReport = {
        ...attendeeReport,
        attendeeList: attendeeReport.attendeeList.map(name => {
            const aliasedName = aliases[name]?.trim();
            return aliasedName || name;
        }),
        currentAttendees: attendeeReport.currentAttendees.map(attendee => ({
            ...attendee,
            name: aliases[attendee.name]?.trim() || attendee.name
        })),
        attendeeHistory: attendeeReport.attendeeHistory.map(event => ({
            ...event,
            name: aliases[event.name]?.trim() || event.name
        }))
    };
    
    return aliasedReport;
}

const CONFIG_URL = chrome.runtime.getURL('config.json');
let cachedConfig = null;
const SCREENSHOT_KEY_PREFIX = 'screenshots_';
const MAX_SCREENSHOTS_PER_MEETING = 20;
const screenshotBuffers = new Map();

async function loadExtensionConfig() {
    if (cachedConfig) {
        return cachedConfig;
    }

    const response = await fetch(CONFIG_URL);
    cachedConfig = await response.json();
    return cachedConfig;
}

async function ensureDefaultSettings() {
    try {
        const config = await loadExtensionConfig();
        const defaults = {
            autoEnableCaptions: config.autoEnableCaptions,
            autoSaveOnEnd: config.autoSaveOnEnd,
            defaultSaveFormat: config.defaultSaveFormat,
            trackCaptions: config.trackCaptions,
            trackAttendees: config.trackAttendees,
            autoOpenAttendees: config.autoOpenAttendees,
            filenamePattern: config.filenamePattern,
            timestampFormat: config.timestampFormat,
            screenshotEnabled: config.screenshotEnabled,
            screenshotIntervalSeconds: config.screenshotIntervalSeconds,
            screenshotDiffThreshold: config.screenshotDiffThreshold
        };

        const stored = await chrome.storage.sync.get(Object.keys(defaults));
        const updates = {};

        for (const [key, value] of Object.entries(defaults)) {
            if (value === undefined) continue;
            if (stored[key] !== value) {
                updates[key] = value;
            }
        }

        if (config.allowedSaveFormats && !config.allowedSaveFormats.includes(defaults.defaultSaveFormat)) {
            console.warn('[Service Worker] Config default save format is not allowed. Falling back to first allowed format.');
            updates.defaultSaveFormat = config.allowedSaveFormats[0] || 'md';
        }

        if (Object.keys(updates).length > 0) {
            await chrome.storage.sync.set(updates);
        }
    } catch (error) {
        console.error('[Service Worker] Failed to ensure default settings:', error);
    }
}

ensureDefaultSettings();

function getScreenshotStorageKey(meetingId) {
    return `${SCREENSHOT_KEY_PREFIX}${meetingId}`;
}

async function hasTabsPermission() {
    if (!chrome.permissions?.contains) {
        return true;
    }
    try {
        return await chrome.permissions.contains({ permissions: ['tabs'] });
    } catch (error) {
        console.warn('[Service Worker] Unable to verify tabs permission:', error);
        return false;
    }
}

async function storeScreenshotForMeeting(meetingId, screenshot) {
    if (!meetingId || !screenshot?.dataUrl || !screenshot?.timestamp) {
        return;
    }

    const storageKey = getScreenshotStorageKey(meetingId);
    let buffer = screenshotBuffers.get(meetingId);

    if (!buffer) {
        const stored = await chrome.storage.local.get(storageKey);
        buffer = Array.isArray(stored[storageKey]) ? stored[storageKey] : [];
    }

    buffer.push(screenshot);
    if (buffer.length > MAX_SCREENSHOTS_PER_MEETING) {
        buffer = buffer.slice(-MAX_SCREENSHOTS_PER_MEETING);
    }

    screenshotBuffers.set(meetingId, buffer);
    await chrome.storage.local.set({ [storageKey]: buffer });
}

async function clearScreenshotsForMeeting(meetingId) {
    if (!meetingId) {
        return;
    }

    screenshotBuffers.delete(meetingId);
    const storageKey = getScreenshotStorageKey(meetingId);
    await chrome.storage.local.remove(storageKey);
}

// --- Formatting Functions ---
function formatAsTxt(transcript, attendeeReport) {
    let content = '';
    
    console.log('[Teams Caption Saver] formatAsTxt called with:', {
        transcriptLength: transcript?.length,
        hasAttendeeReport: !!attendeeReport,
        attendeeCount: attendeeReport?.totalUniqueAttendees || 0,
        attendeeList: attendeeReport?.attendeeList || []
    });
    
    // Add attendee information if available
    if (attendeeReport && attendeeReport.totalUniqueAttendees > 0) {
        content += '=== MEETING ATTENDEES ===\n';
        content += `Total Attendees: ${attendeeReport.totalUniqueAttendees}\n`;
        content += `Meeting Start: ${new Date(attendeeReport.meetingStartTime).toLocaleString()}\n`;
        content += '\nAttendee List:\n';
        attendeeReport.attendeeList.forEach(name => {
            content += `- ${name}\n`;
        });
        content += '\n=== TRANSCRIPT ===\n';
    }
    
    content += transcript.map(entry => `[${entry.Time}] ${entry.Name}: ${entry.Text}`).join('\n');
    return content;
}

function formatAsMarkdown(transcript, attendeeReport) {
    let content = '';
    
    // Add attendee information if available
    if (attendeeReport && attendeeReport.totalUniqueAttendees > 0) {
        content += '# Meeting Attendees\n\n';
        content += `**Total Attendees:** ${attendeeReport.totalUniqueAttendees}\n\n`;
        content += `**Meeting Start:** ${new Date(attendeeReport.meetingStartTime).toLocaleString()}\n\n`;
        content += '## Attendee List\n\n';
        attendeeReport.attendeeList.forEach(name => {
            content += `- ${name}\n`;
        });
        content += '\n---\n\n# Transcript\n\n';
    }
    
    let lastSpeaker = null;
    content += transcript.map(entry => {
        if (entry.Name !== lastSpeaker) {
            lastSpeaker = entry.Name;
            return `\n**${entry.Name}** (${entry.Time}):\n> ${entry.Text}`;
        }
        return `> ${entry.Text}`;
    }).join('\n').trim();
    
    return content;
}

// --- Core Actions ---
async function downloadFile(filename, content, mimeType, saveAs) {
    const url = `data:${mimeType};charset=utf-8,${encodeURIComponent(content)}`;
    chrome.downloads.download({
        url: url,
        filename: filename,
        saveAs: saveAs
    });
    
    // Notify viewer that transcript was saved
    try {
        const tabs = await chrome.tabs.query({});
        for (const tab of tabs) {
            if (tab.url && tab.url.includes('viewer.html')) {
                chrome.tabs.sendMessage(tab.id, { message: 'transcript_saved' });
            }
        }
    } catch (error) {
        // Silent fail if viewer is not open
    }
}

async function generateFilename(pattern, meetingTitle, format, attendeeReport) {
    const config = await loadExtensionConfig();
    const now = new Date();
    const dateStr = now.toISOString().split('T')[0]; // YYYY-MM-DD
    const timeStr = now.toTimeString().split(' ')[0].replace(/:/g, '-'); // HH-MM-SS
    const attendeeCount = attendeeReport ? attendeeReport.totalUniqueAttendees : 0;

    const replacements = {
        '{date}': dateStr,
        '{time}': timeStr,
        '{title}': getSanitizedMeetingName(meetingTitle),
        '{format}': format,
        '{attendees}': attendeeCount > 0 ? `${attendeeCount}_attendees` : ''
    };
    
    const defaultPattern = config.filenamePattern || '{date}_{title}';
    let filename = pattern || defaultPattern;
    for (const [key, value] of Object.entries(replacements)) {
        filename = filename.replace(new RegExp(key.replace(/[{}]/g, '\\$&'), 'g'), value);
    }
    
    // Clean up any double underscores or trailing underscores
    filename = filename.replace(/__+/g, '_').replace(/_+$/, '');
    
    return filename;
}

async function saveTranscript(meetingTitle, transcriptArray, aliases, format, recordingStartTime, saveAsPrompt, attendeeReport = null) {
    const processedTranscript = applyAliasesToTranscript(transcriptArray, aliases);
    const processedAttendeeReport = applyAliasesToAttendeeReport(attendeeReport, aliases);

    const config = await loadExtensionConfig();
    const allowedFormats = config.allowedSaveFormats || ['md', 'txt'];
    let selectedFormat = allowedFormats.includes(format) ? format : null;
    if (!selectedFormat) {
        const configDefault = config.defaultSaveFormat;
        if (configDefault && allowedFormats.includes(configDefault)) {
            selectedFormat = configDefault;
        } else {
            selectedFormat = allowedFormats[0] || 'md';
        }
    }

    // Get filename pattern from settings
    const { filenamePattern } = await chrome.storage.sync.get('filenamePattern');
    const filename = await generateFilename(filenamePattern, meetingTitle, selectedFormat, processedAttendeeReport);

    let content, extension, mimeType;

    switch (selectedFormat) {
        case 'md':
            content = formatAsMarkdown(processedTranscript, processedAttendeeReport);
            extension = 'md';
            mimeType = 'text/markdown';
            break;
        case 'txt':
        default:
            content = formatAsTxt(processedTranscript, processedAttendeeReport);
            extension = 'txt';
            mimeType = 'text/plain';
            break;
    }

    // Add extension to filename
    const fullFilename = `${filename}.${extension}`;
    downloadFile(fullFilename, content, mimeType, saveAsPrompt);
}

// --- State Management ---
let lastAutoSaveId = null;
let autoSaveInProgress = false;

async function createViewerTab(transcriptArray) {
    await chrome.storage.local.set({ captionsToView: transcriptArray });
    chrome.tabs.create({ url: chrome.runtime.getURL('viewer.html') });
}

function updateBadge(isCapturing) {
    if (isCapturing) {
        chrome.action.setBadgeText({ text: 'ON' });
        chrome.action.setBadgeBackgroundColor({ color: '#28a745' }); // Green
    } else {
        chrome.action.setBadgeText({ text: 'OFF' });
        chrome.action.setBadgeBackgroundColor({ color: '#6c757d' }); // Grey
    }
}

// --- Event Listeners ---
// Helper function to chunk arrays
function chunkArray(array, chunkSize) {
    const chunks = [];
    for (let i = 0; i < array.length; i += chunkSize) {
        chunks.push(array.slice(i, i + chunkSize));
    }
    return chunks;
}

// Helper function to calculate duration
function calculateDuration(transcriptArray) {
    if (!transcriptArray || transcriptArray.length === 0) return '0 min';
    
    try {
        const firstTime = new Date(transcriptArray[0].Time);
        const lastTime = new Date(transcriptArray[transcriptArray.length - 1].Time);
        
        // Check if dates are valid
        if (isNaN(firstTime.getTime()) || isNaN(lastTime.getTime())) {
            // Fallback: estimate based on caption count (avg 3 seconds per caption)
            const estimatedMinutes = Math.round((transcriptArray.length * 3) / 60);
            return `~${estimatedMinutes} min`;
        }
        
        const durationMs = lastTime - firstTime;
        const minutes = Math.round(durationMs / 60000);
        
        if (minutes < 60) {
            return `${minutes} min`;
        } else {
            const hours = Math.floor(minutes / 60);
            const mins = minutes % 60;
            return `${hours}h ${mins}m`;
        }
    } catch (error) {
        // If all else fails, show caption count
        return `${transcriptArray.length} captions`;
    }
}

chrome.runtime.onInstalled.addListener(() => {
    updateBadge(false);
    ensureDefaultSettings();
});

chrome.runtime.onStartup.addListener(() => {
    updateBadge(false);
    ensureDefaultSettings();
});

chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
    if (message.message === 'capture_screenshot') {
        (async () => {
            if (!chrome.tabs?.captureVisibleTab) {
                sendResponse({ error: 'capture_not_supported' });
                return;
            }

            const permissionGranted = await hasTabsPermission();
            if (!permissionGranted) {
                sendResponse({ error: 'missing_permission' });
                return;
            }

            try {
                const windowId = sender?.tab?.windowId;
                const dataUrl = await chrome.tabs.captureVisibleTab(windowId, { format: 'png' });
                sendResponse({ dataUrl });
            } catch (error) {
                console.error('[Service Worker] Failed to capture screenshot:', error);
                sendResponse({ error: error?.message || 'capture_failed' });
            }
        })();

        return true;
    }

    (async () => {
        const { speakerAliases } = await chrome.storage.session.get('speakerAliases');

        switch (message.message) {
            case 'save_session_history':
                // Save meeting to session history using chrome.storage directly
                try {
                    // Since we can't import in service worker, implement inline
                    const sessionId = `session_${Date.now()}`;
                    const transcriptArray = message.transcriptArray;
                    const meetingTitle = message.meetingTitle;
                    const attendeeReport = message.attendeeReport;

                    // Create session metadata
                    const metadata = {
                        id: sessionId,
                        title: meetingTitle || 'Untitled Meeting',
                        timestamp: new Date().toISOString(),
                        date: new Date().toLocaleDateString(),
                        time: new Date().toLocaleTimeString(),
                        captionCount: transcriptArray.length,
                        duration: calculateDuration(transcriptArray),
                        speakers: [...new Set(transcriptArray.map(c => c.Name))].slice(0, 10),
                        attendees: attendeeReport?.attendeeList?.slice(0, 20),
                        attendeeCount: attendeeReport?.totalUniqueAttendees || 0,
                        preview: transcriptArray.slice(0, 3).map(c => `${c.Name}: ${c.Text.substring(0, 50)}`).join(' | ')
                    };

                    // Save transcript in chunks to avoid size limits
                    const chunks = chunkArray(transcriptArray, 100); // 100 items per chunk
                    for (let i = 0; i < chunks.length; i++) {
                        await chrome.storage.local.set({
                            [`${sessionId}_chunk_${i}`]: chunks[i]
                        });
                    }
                    metadata.chunkCount = chunks.length;

                    // Save attendee report if exists
                    if (attendeeReport) {
                        await chrome.storage.local.set({
                            [`${sessionId}_attendees`]: attendeeReport
                        });
                    }

                    // Update session index
                    const { session_index = [] } = await chrome.storage.local.get('session_index');
                    session_index.push(metadata);

                    // Keep only last 10 sessions
                    if (session_index.length > 10) {
                        const toDelete = session_index.shift();
                        // Clean up old session data
                        const keysToDelete = [];
                        for (let i = 0; i < toDelete.chunkCount; i++) {
                            keysToDelete.push(`${toDelete.id}_chunk_${i}`);
                        }
                        keysToDelete.push(`${toDelete.id}_attendees`);
                        await chrome.storage.local.remove(keysToDelete);
                    }

                    // Sort by timestamp (newest first)
                    session_index.sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp));

                    await chrome.storage.local.set({ 'session_index': session_index });
                    console.log('[Service Worker] Session saved to history:', sessionId);

                } catch (error) {
                    console.error('[Service Worker] Failed to save session:', error);
                }
                break;

            case 'download_captions':
                console.log('[Teams Caption Saver] Download request received:', {
                    format: message.format,
                    transcriptCount: message.transcriptArray?.length,
                    hasAttendeeReport: !!message.attendeeReport,
                    attendeeCount: message.attendeeReport?.totalUniqueAttendees || 0
                });
                await saveTranscript(message.meetingTitle, message.transcriptArray, speakerAliases, message.format, message.recordingStartTime, true, message.attendeeReport);
                break;

            case 'save_on_leave':
                // Generate unique ID for this save request
                const saveId = `${message.meetingTitle}_${message.recordingStartTime}`;

                // Prevent duplicate saves
                if (autoSaveInProgress || lastAutoSaveId === saveId) {
                    console.log('Auto-save already in progress or completed for this meeting, skipping...');
                    break;
                }

                autoSaveInProgress = true;
                lastAutoSaveId = saveId;

                try {
                    const settings = await chrome.storage.sync.get(['autoSaveOnEnd', 'defaultSaveFormat']);
                    if (settings.autoSaveOnEnd && message.transcriptArray.length > 0) {
                        const config = await loadExtensionConfig();
                        const allowedFormats = config.allowedSaveFormats || ['md', 'txt'];
                        let formatToSave = settings.defaultSaveFormat;
                        if (!formatToSave || !allowedFormats.includes(formatToSave)) {
                            if (config.defaultSaveFormat && allowedFormats.includes(config.defaultSaveFormat)) {
                                formatToSave = config.defaultSaveFormat;
                            } else {
                                formatToSave = allowedFormats[0] || 'md';
                            }
                        }
                        console.log(`Auto-saving transcript in ${formatToSave.toUpperCase()} format.`);
                        await saveTranscript(message.meetingTitle, message.transcriptArray, speakerAliases, formatToSave, message.recordingStartTime, false, message.attendeeReport);
                        console.log('Auto-save completed successfully.');
                    }
                } catch (error) {
                    console.error('Auto-save failed:', error);
                    // Reset state on error to allow retry
                    lastAutoSaveId = null;
                } finally {
                    autoSaveInProgress = false;
                }
                break;

            case 'display_captions':
                await createViewerTab(message.transcriptArray);
                break;

            case 'store_screenshot':
                try {
                    await storeScreenshotForMeeting(message.meetingId, message.screenshot);
                } catch (error) {
                    console.error('[Service Worker] Failed to store screenshot:', error);
                }
                break;

            case 'update_badge_status':
                updateBadge(message.capturing);
                // Reset auto-save state when starting a new capture session
                if (message.capturing) {
                    lastAutoSaveId = null;
                    autoSaveInProgress = false;
                    console.log('New capture session started, auto-save state reset.');
                } else if (message.meetingId) {
                    await clearScreenshotsForMeeting(message.meetingId);
                }
                break;

            case 'error_logged':
                // Central error logging - could send to analytics service
                console.warn('[Teams Caption Saver] Error logged:', message.error);
                // Could implement error reporting here
                break;
        }
    })();

    return true; // Indicates that the response will be sent asynchronously
});