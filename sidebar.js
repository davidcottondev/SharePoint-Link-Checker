// Sidebar JavaScript functionality
(function() {
    'use strict';

    let currentTabId = null;
    let isSharePointSite = false;
    let userSettings = null;
    let originalTabId = null; // Remember the tab we were originally opened on
    let hasInitialScanRun = false; // Track if we've done the initial scan

    // Default settings
    const defaultSettings = {
        linkTypes: {
            teams: true,
            onedrive: true,
            external: true,
            email: true
        },
        externalStatusCodes: {
            '2xx': true,
            '3xx': true,
            '4xx': true,
            '5xx': true,
            'network': true
        }
    };

    // DOM elements
    const loadingState = document.getElementById('loading-state');
    const notSharePointDiv = document.getElementById('not-sharepoint');
    const sharePointDetectedDiv = document.getElementById('sharepoint-detected');
    const siteTitle = document.getElementById('site-title');
    const pageTitle = document.getElementById('page-title');
    const pageUrl = document.getElementById('page-url');
    const teamCount = document.getElementById('team-count');
    const oneDriveCount = document.getElementById('onedrive-count');
    const externalCount = document.getElementById('external-count');
    const emailCount = document.getElementById('email-count');
    const teamCountBox = document.getElementById('team-count-box');
    const oneDriveCountBox = document.getElementById('onedrive-count-box');
    const externalCountBox = document.getElementById('external-count-box');
    const emailCountBox = document.getElementById('email-count-box');
    const teamLinksSection = document.getElementById('team-links-section');
    const teamLinksList = document.getElementById('team-links-list');
    const oneDriveLinksSection = document.getElementById('onedrive-links-section');
    const oneDriveLinksList = document.getElementById('onedrive-links-list');
    const externalLinksSection = document.getElementById('external-links-section');
    const externalLinksList = document.getElementById('external-links-list');
    const emailLinksSection = document.getElementById('email-links-section');
    const emailLinksList = document.getElementById('email-links-list');
    const httpLegend = document.getElementById('http-legend');
    const resultsSection = document.getElementById('results-section');
    const resultsContent = document.getElementById('results-content');
    const settingsBtn = document.getElementById('settings-btn');
    const refreshBtn = document.getElementById('refresh-btn');

    // Initialize the sidebar
    async function init() {
        await loadUserSettings();
        showLoading();
        await checkCurrentTab();
        setupEventListeners();
    }

    // Load user settings
    async function loadUserSettings() {
        try {
            const result = await chrome.storage.sync.get(['linkCheckerSettings']);
            userSettings = result.linkCheckerSettings || defaultSettings;
        } catch (error) {
            console.error('Error loading settings:', error);
            userSettings = defaultSettings;
        }
    }

    function showLoading() {
        hideAllStates();
        loadingState.style.display = 'block';
    }

    function hideAllStates() {
        loadingState.style.display = 'none';
        notSharePointDiv.style.display = 'none';
        sharePointDetectedDiv.style.display = 'none';
        resultsSection.style.display = 'none';
    }

    function showNotSharePoint() {
        hideAllStates();
        notSharePointDiv.style.display = 'block';
    }

    function showSharePointDetected(siteInfo) {
        hideAllStates();
        sharePointDetectedDiv.style.display = 'block';
        
        if (siteInfo) {
            siteTitle.textContent = siteInfo.siteTitle || 'Unknown Site';
            pageTitle.textContent = siteInfo.pageTitle || 'Unknown Page';
            pageUrl.textContent = siteInfo.cleanUrl || siteInfo.fullUrl || 'Unknown URL';
        }
        
        // Automatically run the scan when SharePoint site is detected (only on initial load)
        if (!hasInitialScanRun) {
            // Show that analysis is starting
            teamLinksSection.style.display = 'block';
            teamLinksList.innerHTML = '<div class="no-links-message">🔄 Analyzing page for team links...</div>';
            oneDriveLinksSection.style.display = 'block';
            oneDriveLinksList.innerHTML = '<div class="no-links-message">🔄 Analyzing page for OneDrive links...</div>';
            externalLinksSection.style.display = 'block';
            externalLinksList.innerHTML = '<div class="no-links-message">🔄 Analyzing page for external links...</div>';
            emailLinksSection.style.display = 'block';
            emailLinksList.innerHTML = '<div class="no-links-message">🔄 Analyzing page for email links...</div>';
            
            setTimeout(() => {
                scanForLinks();
                hasInitialScanRun = true;
            }, 500);
        } else {
            // Show message that user needs to refresh for updates
            teamLinksSection.style.display = 'block';
            teamLinksList.innerHTML = '<div class="no-links-message">📄 Links analyzed. Click 🔄 to refresh analysis.</div>';
            oneDriveLinksSection.style.display = 'block';
            oneDriveLinksList.innerHTML = '<div class="no-links-message">📄 Links analyzed. Click 🔄 to refresh analysis.</div>';
            externalLinksSection.style.display = 'block';
            externalLinksList.innerHTML = '<div class="no-links-message">📄 Links analyzed. Click 🔄 to refresh analysis.</div>';
            emailLinksSection.style.display = 'block';
            emailLinksList.innerHTML = '<div class="no-links-message">📄 Links analyzed. Click 🔄 to refresh analysis.</div>';
        }
    }

    async function checkCurrentTab() {
        try {
            const tabs = await chrome.tabs.query({ active: true, currentWindow: true });
            if (tabs.length > 0) {
                currentTabId = tabs[0].id;
                
                // Remember the original tab if not set
                if (originalTabId === null) {
                    originalTabId = currentTabId;
                }
                
                const url = tabs[0].url;
                
                try {
                    // Try to check if it's a SharePoint site via content script
                    const response = await chrome.tabs.sendMessage(currentTabId, { 
                        type: 'CHECK_SHAREPOINT_SITE' 
                    });
                    
                    if (response && response.isSharePoint) {
                        isSharePointSite = true;
                        showSharePointDetected(response.siteInfo);
                    } else {
                        isSharePointSite = false;
                        showNotSharePoint();
                    }
                } catch (contentScriptError) {
                    console.log('Content script not available, checking URL directly');
                    // Fallback to URL-based check
                    if (url && (url.includes('sharepoint.com') || url.includes('sharepoint-df.com'))) {
                        isSharePointSite = true;
                        const fallbackSiteInfo = {
                            siteTitle: 'SharePoint Site',
                            pageTitle: tabs[0].title || 'Unknown Page',
                            cleanUrl: url,
                            fullUrl: url
                        };
                        showSharePointDetected(fallbackSiteInfo);
                    } else {
                        isSharePointSite = false;
                        showNotSharePoint();
                    }
                }
            } else {
                showNotSharePoint();
            }
        } catch (error) {
            console.error('Error checking current tab:', error);
            showNotSharePoint();
        }
    }

    function setupEventListeners() {
        settingsBtn.addEventListener('click', openSettings);
        refreshBtn.addEventListener('click', refreshDetection);
    }

    // Open settings page
    function openSettings() {
        chrome.tabs.create({
            url: chrome.runtime.getURL('settings.html')
        });
    }

    // Refresh SharePoint detection
    async function refreshDetection() {
        // Reset the scan flag to allow a new scan
        hasInitialScanRun = false;
        // Reset the original tab to current tab
        originalTabId = null;
        
        showLoading();
        await checkCurrentTab();
    }

    async function scanForLinks() {
        if (!isSharePointSite || !currentTabId) {
            return;
        }

        // Reload settings in case they changed
        await loadUserSettings();

        try {
            // First, try to inject the content script if it's not already there
            await ensureContentScriptLoaded();
            
            // Get links based on user settings
            const promises = [];
            let teamResponse = null, oneDriveResponse = null, externalResponse = null, emailResponse = null;

            if (userSettings.linkTypes.teams) {
                promises.push(chrome.tabs.sendMessage(currentTabId, { type: 'GET_TEAM_LINKS' }));
            } else {
                promises.push(Promise.resolve(null));
            }

            if (userSettings.linkTypes.onedrive) {
                promises.push(chrome.tabs.sendMessage(currentTabId, { type: 'GET_ONEDRIVE_LINKS' }));
            } else {
                promises.push(Promise.resolve(null));
            }

            if (userSettings.linkTypes.external) {
                promises.push(chrome.tabs.sendMessage(currentTabId, { type: 'GET_EXTERNAL_LINKS' }));
            } else {
                promises.push(Promise.resolve(null));
            }

            if (userSettings.linkTypes.email) {
                promises.push(chrome.tabs.sendMessage(currentTabId, { type: 'GET_EMAIL_LINKS' }));
            } else {
                promises.push(Promise.resolve(null));
            }

            [teamResponse, oneDriveResponse, externalResponse, emailResponse] = await Promise.all(promises);
            
            // Display results based on settings
            if (userSettings.linkTypes.teams) {
                if (teamResponse && teamResponse.teamLinks) {
                    displayTeamLinks(teamResponse.teamLinks);
                } else {
                    displayTeamLinks([]);
                }
            } else {
                hideTeamLinks();
            }
            
            if (userSettings.linkTypes.onedrive) {
                if (oneDriveResponse && oneDriveResponse.oneDriveLinks) {
                    displayOneDriveLinks(oneDriveResponse.oneDriveLinks);
                } else {
                    displayOneDriveLinks([]);
                }
            } else {
                hideOneDriveLinks();
            }
            
            if (userSettings.linkTypes.external) {
                if (externalResponse && externalResponse.externalLinks) {
                    await displayExternalLinksWithStatus(externalResponse.externalLinks);
                } else {
                    displayExternalLinks([]);
                }
            } else {
                hideExternalLinks();
            }
            
            if (userSettings.linkTypes.email) {
                if (emailResponse && emailResponse.emailLinks) {
                    displayEmailLinks(emailResponse.emailLinks);
                } else {
                    displayEmailLinks([]);
                }
            } else {
                hideEmailLinks();
            }
        } catch (error) {
            console.error('Error scanning links:', error);
            
            // Try to inject content script and retry once
            try {
                await injectContentScript();
                const [teamResponse, oneDriveResponse, externalResponse, emailResponse] = await Promise.all([
                    chrome.tabs.sendMessage(currentTabId, { type: 'GET_TEAM_LINKS' }),
                    chrome.tabs.sendMessage(currentTabId, { type: 'GET_ONEDRIVE_LINKS' }),
                    chrome.tabs.sendMessage(currentTabId, { type: 'GET_EXTERNAL_LINKS' }),
                    chrome.tabs.sendMessage(currentTabId, { type: 'GET_EMAIL_LINKS' })
                ]);
                
                if (teamResponse && teamResponse.teamLinks) {
                    displayTeamLinks(teamResponse.teamLinks);
                } else {
                    displayTeamLinks([]);
                }
                
                if (oneDriveResponse && oneDriveResponse.oneDriveLinks) {
                    displayOneDriveLinks(oneDriveResponse.oneDriveLinks);
                } else {
                    displayOneDriveLinks([]);
                }
                
                if (externalResponse && externalResponse.externalLinks) {
                    await displayExternalLinksWithStatus(externalResponse.externalLinks);
                } else {
                    displayExternalLinks([]);
                }
                
                if (emailResponse && emailResponse.emailLinks) {
                    displayEmailLinks(emailResponse.emailLinks);
                } else {
                    displayEmailLinks([]);
                }
            } catch (retryError) {
                console.error('Retry failed:', retryError);
                showError('Failed to scan links. Please refresh the page and try again.');
            }
        }
    }

    async function ensureContentScriptLoaded() {
        try {
            // Try to ping the content script
            await chrome.tabs.sendMessage(currentTabId, { type: 'PING' });
        } catch (error) {
            // Content script not loaded, inject it
            await injectContentScript();
        }
    }

    async function injectContentScript() {
        try {
            await chrome.scripting.executeScript({
                target: { tabId: currentTabId },
                files: ['content.js']
            });
            
            // Wait a moment for the script to initialize
            await new Promise(resolve => setTimeout(resolve, 500));
        } catch (error) {
            console.error('Failed to inject content script:', error);
            throw error;
        }
    }

    function displayTeamLinks(teamLinks) {
        // Show the count box and update team count
        teamCountBox.style.display = 'flex';
        teamCount.textContent = teamLinks.length;
        
        // Show/hide team links section based on count
        if (teamLinks.length > 0) {
            teamLinksSection.style.display = 'block';
            
            // Clear previous links
            teamLinksList.innerHTML = '';
            
            // Add each team link
            teamLinks.forEach(link => {
                const linkElement = createLinkDisplayElement(link.text, link.url, link.isFooter, link.isNav, false);
                teamLinksList.appendChild(linkElement);
            });
        } else {
            teamLinksSection.style.display = 'block';
            teamLinksList.innerHTML = '<div class="no-links-message">No team links found on this page.</div>';
        }
    }

    function displayOneDriveLinks(oneDriveLinks) {
        // Show the count box and update OneDrive count
        oneDriveCountBox.style.display = 'flex';
        oneDriveCount.textContent = oneDriveLinks.length;
        
        // Show/hide OneDrive links section based on count
        if (oneDriveLinks.length > 0) {
            oneDriveLinksSection.style.display = 'block';
            
            // Clear previous links
            oneDriveLinksList.innerHTML = '';
            
            // Add each OneDrive link
            oneDriveLinks.forEach(link => {
                const linkElement = createLinkDisplayElement(link.text, link.url, link.isFooter, link.isNav, false);
                oneDriveLinksList.appendChild(linkElement);
            });
        } else {
            oneDriveLinksSection.style.display = 'block';
            oneDriveLinksList.innerHTML = '<div class="no-links-message">No OneDrive links found on this page.</div>';
        }
    }

    function displayExternalLinks(externalLinks) {
        // Show the count box and update external count
        externalCountBox.style.display = 'flex';
        externalCount.textContent = externalLinks.length;
        
        // Show/hide external links section based on count
        if (externalLinks.length > 0) {
            externalLinksSection.style.display = 'block';
            httpLegend.style.display = 'none'; // Hide legend for basic display
            
            // Clear previous links
            externalLinksList.innerHTML = '';
            
            // Add each external link
            externalLinks.forEach(link => {
                const linkElement = createLinkDisplayElement(link.text, link.url, link.isFooter, link.isNav, link.isMailto || false);
                externalLinksList.appendChild(linkElement);
            });
        } else {
            externalLinksSection.style.display = 'block';
            httpLegend.style.display = 'none';
            externalLinksList.innerHTML = '<div class="no-links-message">No external links found on this page.</div>';
        }
    }

    function displayEmailLinks(emailLinks) {
        // Show the count box and update email count
        emailCountBox.style.display = 'flex';
        emailCount.textContent = emailLinks.length;
        
        // Show/hide email links section based on count
        if (emailLinks.length > 0) {
            emailLinksSection.style.display = 'block';
            
            // Clear previous links
            emailLinksList.innerHTML = '';
            
            // Add each email link
            emailLinks.forEach(link => {
                const linkElement = createLinkDisplayElement(link.text, link.url, link.isFooter, link.isNav, true);
                emailLinksList.appendChild(linkElement);
            });
        } else {
            emailLinksSection.style.display = 'block';
            emailLinksList.innerHTML = '<div class="no-links-message">No email links found on this page.</div>';
        }
    }

    async function displayExternalLinksWithStatus(externalLinks) {
        // Show the count box
        externalCountBox.style.display = 'flex';
        
        if (externalLinks.length > 0) {
            externalLinksSection.style.display = 'block';
            httpLegend.style.display = 'none'; // Hide initially
            
            // Show loading message
            externalLinksList.innerHTML = '<div class="no-links-message">🔄 Checking HTTP status for external links...</div>';
            
            try {
                // Check HTTP status for all external links
                const response = await chrome.tabs.sendMessage(currentTabId, {
                    type: 'CHECK_EXTERNAL_LINKS_STATUS',
                    links: externalLinks
                });
                
                if (response && response.externalLinks) {
                    const linksWithStatus = response.externalLinks;
                    
                    // Filter links based on user settings to get actual display count
                    const filteredLinks = linksWithStatus.filter(link => {
                        const status = link.status;
                        
                        if (status === 0 || status === 'Error' || status === 'Timeout') {
                            return userSettings.externalStatusCodes.network;
                        } else if (status >= 500 && status < 600) {
                            return userSettings.externalStatusCodes['5xx'];
                        } else if (status >= 400 && status < 500) {
                            return userSettings.externalStatusCodes['4xx'];
                        } else if (status >= 300 && status < 400) {
                            return userSettings.externalStatusCodes['3xx'];
                        } else if (status >= 200 && status < 300) {
                            return userSettings.externalStatusCodes['2xx'];
                        }
                        
                        return userSettings.externalStatusCodes.network;
                    });
                    
                    // Update external count to reflect only displayed links
                    externalCount.textContent = filteredLinks.length;
                    
                    // Show legend if we have links with status
                    if (filteredLinks.length > 0) {
                        httpLegend.style.display = 'block';
                    }
                    
                    // Clear previous links
                    externalLinksList.innerHTML = '';
                    
                    // Group links by status and display
                    displayLinksGroupedByStatus(linksWithStatus);
                } else {
                    // Update external count for basic display
                    externalCount.textContent = externalLinks.length;
                    displayExternalLinks(externalLinks);
                }
            } catch (error) {
                console.error('Error checking external link status:', error);
                // Update external count for fallback basic display
                externalCount.textContent = externalLinks.length;
                displayExternalLinks(externalLinks);
            }
        } else {
            externalCount.textContent = '0';
            externalLinksSection.style.display = 'block';
            httpLegend.style.display = 'none';
            externalLinksList.innerHTML = '<div class="no-links-message">No external links found on this page.</div>';
        }
    }

    function displayLinksGroupedByStatus(links) {
        // Filter links based on user settings for status codes
        const filteredLinks = links.filter(link => {
            const status = link.status;
            
            if (status === 0 || status === 'Error' || status === 'Timeout') {
                return userSettings.externalStatusCodes.network;
            } else if (status >= 500 && status < 600) {
                return userSettings.externalStatusCodes['5xx'];
            } else if (status >= 400 && status < 500) {
                return userSettings.externalStatusCodes['4xx'];
            } else if (status >= 300 && status < 400) {
                return userSettings.externalStatusCodes['3xx'];
            } else if (status >= 200 && status < 300) {
                return userSettings.externalStatusCodes['2xx'];
            }
            
            // Default to showing unknown status codes if network errors are enabled
            return userSettings.externalStatusCodes.network;
        });

        // Group links by status category
        const groups = {
            'Network/Unknown Errors': [],
            'Server Errors (5xx)': [],
            'Client Errors (4xx)': [],
            'Redirects (3xx)': [],
            'Success (2xx)': []
        };

        filteredLinks.forEach(link => {
            const status = link.status;
            if (status === 0 || status === 'Error' || status === 'Timeout') {
                groups['Network/Unknown Errors'].push(link);
            } else if (status >= 500 && status < 600) {
                groups['Server Errors (5xx)'].push(link);
            } else if (status >= 400 && status < 500) {
                groups['Client Errors (4xx)'].push(link);
            } else if (status >= 300 && status < 400) {
                groups['Redirects (3xx)'].push(link);
            } else if (status >= 200 && status < 300) {
                groups['Success (2xx)'].push(link);
            } else {
                groups['Network/Unknown Errors'].push(link); // fallback
            }
        });

        // Display each group that has links and is enabled in settings
        const groupsToShow = Object.entries(groups).filter(([groupName, groupLinks]) => {
            if (groupLinks.length === 0) return false;
            
            // Check if this group type is enabled in settings
            if (groupName === 'Success (2xx)') return userSettings.externalStatusCodes['2xx'];
            if (groupName === 'Redirects (3xx)') return userSettings.externalStatusCodes['3xx'];
            if (groupName === 'Client Errors (4xx)') return userSettings.externalStatusCodes['4xx'];
            if (groupName === 'Server Errors (5xx)') return userSettings.externalStatusCodes['5xx'];
            if (groupName === 'Network/Unknown Errors') return userSettings.externalStatusCodes.network;
            
            return true; // fallback
        });

        if (groupsToShow.length === 0) {
            externalLinksList.innerHTML = '<div class="no-links-message">No external links match your current filter settings.</div>';
            return;
        }

        groupsToShow.forEach(([groupName, groupLinks]) => {
            const groupElement = createLinkGroup(groupName, groupLinks);
            externalLinksList.appendChild(groupElement);
        });
    }

    function createLinkGroup(groupName, links) {
        const groupDiv = document.createElement('div');
        groupDiv.className = 'link-group';

        // Group header
        const headerDiv = document.createElement('div');
        headerDiv.className = 'link-group-header ' + getStatusClassForGroup(groupName);
        
        const titleSpan = document.createElement('span');
        titleSpan.className = 'link-group-title';
        titleSpan.textContent = groupName;
        
        const countSpan = document.createElement('span');
        countSpan.className = 'link-group-count';
        countSpan.textContent = links.length;
        
        headerDiv.appendChild(titleSpan);
        headerDiv.appendChild(countSpan);

        // Group items
        const itemsDiv = document.createElement('div');
        itemsDiv.className = 'link-group-items';
        
        links.forEach(link => {
            const linkElement = createSimpleLinkElement(link);
            itemsDiv.appendChild(linkElement);
        });

        groupDiv.appendChild(headerDiv);
        groupDiv.appendChild(itemsDiv);
        
        return groupDiv;
    }

    function createSimpleLinkElement(link) {
        const div = document.createElement('div');
        div.className = 'link-item-simple';
        div.title = 'Hover to highlight link on page';
        div.setAttribute('tabindex', '0'); // Make focusable for accessibility
        
        const titleDiv = document.createElement('div');
        titleDiv.className = 'link-title';
        // Add appropriate emoji based on link location or type
        let displayText = link.text || 'No text';
        if (link.isMailto) {
            displayText = `📧 ${displayText}`;
        } else if (link.isFooter) {
            displayText = `🦶 ${displayText}`;
        } else if (link.isNav) {
            displayText = `🧭 ${displayText}`;
        }
        titleDiv.textContent = `${displayText} - ${link.statusText || 'Unknown'}`;
        
        const urlDiv = document.createElement('div');
        urlDiv.className = 'link-url-simple';
        urlDiv.textContent = link.url;
        
        // Add hover handlers for highlighting
        const highlightLink = () => {
            if (currentTabId) {
                chrome.tabs.sendMessage(currentTabId, {
                    type: 'HIGHLIGHT_LINK',
                    url: link.url
                }).catch(() => {
                    // Ignore errors if content script not available
                });
            }
        };
        
        const removeHighlight = () => {
            if (currentTabId) {
                chrome.tabs.sendMessage(currentTabId, {
                    type: 'REMOVE_HIGHLIGHT'
                }).catch(() => {
                    // Ignore errors if content script not available
                });
            }
        };
        
        // Add hover events for highlighting
        div.addEventListener('mouseenter', highlightLink);
        div.addEventListener('mouseleave', removeHighlight);
        div.addEventListener('focus', highlightLink);
        div.addEventListener('blur', removeHighlight);
        
        div.appendChild(titleDiv);
        div.appendChild(urlDiv);
        
        return div;
    }

    function getStatusClassForGroup(groupName) {
        if (groupName.includes('Network') || groupName.includes('Unknown')) {
            return 'status-network-error';
        } else if (groupName.includes('Server')) {
            return 'status-server-error';
        } else if (groupName.includes('Client')) {
            return 'status-client-error';
        } else if (groupName.includes('Redirects')) {
            return 'status-redirect';
        } else if (groupName.includes('Success')) {
            return 'status-success';
        }
        return 'status-network-error';
    }

    function getStatusClass(status) {
        if (status >= 200 && status < 300) {
            return 'status-success';
        } else if (status >= 300 && status < 400) {
            return 'status-redirect';
        } else if (status >= 400 && status < 500) {
            return 'status-client-error';
        } else if (status >= 500 && status < 600) {
            return 'status-server-error';
        } else {
            return 'status-network-error';
        }
    }

    function createLinkDisplayElement(text, url, isFooter, isNav, isMailto) {
        const div = document.createElement('div');
        div.className = 'link-item-display';
        div.title = 'Hover to highlight link on page';
        div.setAttribute('tabindex', '0'); // Make focusable for accessibility
        
        const textSpan = document.createElement('span');
        textSpan.className = 'link-text';
        // Add appropriate emoji based on link location or type
        let displayText = text || 'No text';
        if (isMailto) {
            displayText = `📧 ${displayText}`;
        } else if (isFooter) {
            displayText = `🦶 ${displayText}`;
        } else if (isNav) {
            displayText = `🧭 ${displayText}`;
        }
        textSpan.textContent = displayText;
        
        const urlSpan = document.createElement('span');
        urlSpan.className = 'link-url-display';
        urlSpan.textContent = url;
        
        // Add hover handlers for highlighting
        const highlightLink = () => {
            if (currentTabId) {
                chrome.tabs.sendMessage(currentTabId, {
                    type: 'HIGHLIGHT_LINK',
                    url: url
                }).catch(() => {
                    // Ignore errors if content script not available
                });
            }
        };
        
        const removeHighlight = () => {
            if (currentTabId) {
                chrome.tabs.sendMessage(currentTabId, {
                    type: 'REMOVE_HIGHLIGHT'
                }).catch(() => {
                    // Ignore errors if content script not available
                });
            }
        };
        
        // Add hover events for highlighting
        div.addEventListener('mouseenter', highlightLink);
        div.addEventListener('mouseleave', removeHighlight);
        div.addEventListener('focus', highlightLink);
        div.addEventListener('blur', removeHighlight);
        
        div.appendChild(textSpan);
        div.appendChild(urlSpan);
        
        return div;
    }

    async function displayLinkResults(links) {
        if (links.length === 0) {
            showError('No links found on this page.');
            return;
        }

        // Limit to first 20 links to avoid overwhelming the UI
        const linksToCheck = links.slice(0, 20);
        
        resultsContent.innerHTML = `
            <p class="ms-fontSize-12 ms-fontColor-neutralSecondary">
                Found ${links.length} links. Checking first ${linksToCheck.length}...
            </p>
        `;
        
        resultsSection.style.display = 'block';

        // Create placeholders for each link
        linksToCheck.forEach((link, index) => {
            const linkElement = createLinkElement(link, 'checking');
            linkElement.id = `link-${index}`;
            resultsContent.appendChild(linkElement);
        });

        // Check each link (simplified check - just verify the URL format)
        for (let i = 0; i < linksToCheck.length; i++) {
            const link = linksToCheck[i];
            const linkElement = document.getElementById(`link-${i}`);
            
            // Simple check - validate URL and check if it's an internal SharePoint link
            const status = validateLink(link.url);
            updateLinkElement(linkElement, status);
            
            // Add small delay to show progressive checking
            if (i < linksToCheck.length - 1) {
                await new Promise(resolve => setTimeout(resolve, 100));
            }
        }
    }

    function validateLink(url) {
        try {
            const urlObj = new URL(url);
            
            // Check if it's a SharePoint link
            if (urlObj.hostname.includes('sharepoint.com') || urlObj.hostname.includes('sharepoint-df.com')) {
                return { status: 'success', message: 'SharePoint link' };
            }
            
            // Check if it's an external link
            if (urlObj.protocol === 'http:' || urlObj.protocol === 'https:') {
                return { status: 'warning', message: 'External link' };
            }
            
            // Check for other protocols
            if (urlObj.protocol === 'mailto:') {
                return { status: 'success', message: 'Email link' };
            }
            
            if (urlObj.protocol === 'tel:') {
                return { status: 'success', message: 'Phone link' };
            }
            
            return { status: 'warning', message: 'Other protocol' };
        } catch (error) {
            return { status: 'error', message: 'Invalid URL' };
        }
    }

    function createLinkElement(link, status) {
        const div = document.createElement('div');
        div.className = 'link-item';
        
        const statusIcon = document.createElement('i');
        statusIcon.className = 'ms-Icon link-status';
        
        const urlSpan = document.createElement('span');
        urlSpan.className = 'link-url';
        urlSpan.textContent = link.url;
        urlSpan.title = link.text || link.url;
        
        div.appendChild(statusIcon);
        div.appendChild(urlSpan);
        
        updateLinkElement(div, { status: status });
        
        return div;
    }

    function updateLinkElement(element, result) {
        const statusIcon = element.querySelector('.link-status');
        
        statusIcon.className = 'ms-Icon link-status';
        
        switch (result.status) {
            case 'success':
                statusIcon.classList.add('success', 'ms-Icon--StatusCircleCheckmark');
                statusIcon.title = result.message || 'Link OK';
                break;
            case 'error':
                statusIcon.classList.add('error', 'ms-Icon--StatusCircleErrorX');
                statusIcon.title = result.message || 'Link Error';
                break;
            case 'warning':
                statusIcon.classList.add('warning', 'ms-Icon--Warning');
                statusIcon.title = result.message || 'Link Warning';
                break;
            case 'checking':
            default:
                statusIcon.classList.add('ms-Icon--Info');
                statusIcon.title = 'Checking...';
                break;
        }
    }

    function showError(message) {
        resultsContent.innerHTML = `
            <div style="text-align: center; padding: 20px;">
                <i class="ms-Icon ms-Icon--StatusCircleErrorX ms-fontSize-24" style="color: #d13438;"></i>
                <p class="ms-fontSize-14 ms-fontColor-neutralSecondary" style="margin-top: 8px;">
                    ${message}
                </p>
            </div>
        `;
        resultsSection.style.display = 'block';
    }

    // Hide functions for disabled link types
    function hideTeamLinks() {
        teamLinksSection.style.display = 'none';
        teamCountBox.style.display = 'none';
        teamCount.textContent = '0';
    }

    function hideOneDriveLinks() {
        oneDriveLinksSection.style.display = 'none';
        oneDriveCountBox.style.display = 'none';
        oneDriveCount.textContent = '0';
    }

    function hideExternalLinks() {
        externalLinksSection.style.display = 'none';
        externalCountBox.style.display = 'none';
        externalCount.textContent = '0';
        httpLegend.style.display = 'none';
    }

    function hideEmailLinks() {
        emailLinksSection.style.display = 'none';
        emailCountBox.style.display = 'none';
        emailCount.textContent = '0';
    }

    // Listen for messages from background script
    chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
        if (message.type === 'SHAREPOINT_CHECK') {
            isSharePointSite = message.isSharePoint;
            currentTabId = message.tabId;
            
            if (message.isSharePoint) {
                // Create site info from basic data
                const siteInfo = {
                    siteTitle: 'SharePoint Site',
                    pageTitle: 'Loading...',
                    cleanUrl: message.url,
                    fullUrl: message.url
                };
                showSharePointDetected(siteInfo);
            } else {
                showNotSharePoint();
            }
        } else if (message.type === 'CONTENT_SHAREPOINT_DETECTED') {
            isSharePointSite = true;
            currentTabId = message.tabId || currentTabId;
            showSharePointDetected(message.siteInfo);
        }
    });

    // Listen for visibility changes (removed auto-refresh - now manual only)
    document.addEventListener('visibilitychange', () => {
        // Extension became visible - no automatic refresh, user must manually refresh
        if (!document.hidden) {
            console.log('Extension became visible - use refresh button to update');
        }
    });

    // Initialize when the DOM is loaded
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', init);
    } else {
        init();
    }
})();
