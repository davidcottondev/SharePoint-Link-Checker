// Background service worker for Edge extension
chrome.runtime.onInstalled.addListener(() => {
    console.log('SharePoint Link Checker extension installed');
});

// Handle side panel opening
chrome.sidePanel.setPanelBehavior({ openPanelOnActionClick: true });

// Listen for messages from content script
chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
    if (request.type === 'CHECK_LINK_STATUS') {
        checkLinkStatus(request.url)
            .then(status => sendResponse({ status }))
            .catch(error => sendResponse({ status: 'Error', error: error.message }));
        return true; // Keep the message channel open for async response
    }
});

// Function to check HTTP status of external links
async function checkLinkStatus(url) {
    try {
        const controller = new AbortController();
        const timeoutId = setTimeout(() => controller.abort(), 10000); // 10 second timeout
        
        const response = await fetch(url, {
            method: 'HEAD', // Use HEAD to avoid downloading the full content
            signal: controller.signal,
            cache: 'no-cache'
        });
        
        clearTimeout(timeoutId);
        return response.status;
    } catch (error) {
        if (error.name === 'AbortError') {
            return 'Timeout';
        }
        // Try with GET if HEAD fails
        try {
            const controller = new AbortController();
            const timeoutId = setTimeout(() => controller.abort(), 10000);
            
            const response = await fetch(url, {
                method: 'GET',
                signal: controller.signal,
                cache: 'no-cache'
            });
            
            clearTimeout(timeoutId);
            return response.status;
        } catch (getError) {
            return 'Error';
        }
    }
}

// Listen for tab updates to check if we're on a SharePoint site
chrome.tabs.onUpdated.addListener((tabId, changeInfo, tab) => {
    if (changeInfo.status === 'complete' && tab.url) {
        checkSharePointSite(tab.url, tabId);
    }
});

// Listen for tab activation
chrome.tabs.onActivated.addListener(async (activeInfo) => {
    const tab = await chrome.tabs.get(activeInfo.tabId);
    if (tab.url) {
        checkSharePointSite(tab.url, activeInfo.tabId);
    }
});

function checkSharePointSite(url, tabId) {
    const isSharePointSite = url.includes('sharepoint.com') || url.includes('sharepoint-df.com');
    
    // Send message to sidebar if it's open
    chrome.runtime.sendMessage({
        type: 'SHAREPOINT_CHECK',
        isSharePoint: isSharePointSite,
        url: url,
        tabId: tabId
    }).catch(() => {
        // Sidebar might not be open, ignore the error
    });
}
