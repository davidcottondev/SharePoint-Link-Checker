// Content script to interact with SharePoint pages
(function() {
    'use strict';

    // Check if this is a SharePoint site
    function isSharePointSite() {
        return window.location.hostname.includes('sharepoint.com') || 
               window.location.hostname.includes('sharepoint-df.com') ||
               document.querySelector('meta[name="ms.sharepointpnpversion"]') !== null ||
               document.querySelector('script[src*="sharepoint"]') !== null ||
               window._spPageContextInfo !== undefined;
    }

    // Scroll to bottom to trigger lazy loading, then return to top
    async function scrollToLoadAllContent() {
        return new Promise((resolve) => {
            // Look for SharePoint's content scroll region first
            const contentScrollRegion = document.querySelector('[data-automation-id="contentScrollRegion"]');
            const scrollElement = contentScrollRegion || window;
            const isElementScroll = !!contentScrollRegion;
            
            // Store original scroll position
            const originalScrollTop = isElementScroll ? 
                contentScrollRegion.scrollTop : 
                (window.pageYOffset || document.documentElement.scrollTop);
            
            // Scroll to bottom gradually to trigger lazy loading
            const scrollStep = () => {
                const currentScroll = isElementScroll ? 
                    contentScrollRegion.scrollTop : 
                    (window.pageYOffset || document.documentElement.scrollTop);
                    
                const scrollHeight = isElementScroll ? 
                    contentScrollRegion.scrollHeight : 
                    Math.max(
                        document.body.scrollHeight,
                        document.body.offsetHeight,
                        document.documentElement.clientHeight,
                        document.documentElement.scrollHeight,
                        document.documentElement.offsetHeight
                    );
                    
                const clientHeight = isElementScroll ? 
                    contentScrollRegion.clientHeight : 
                    window.innerHeight;
                
                // If we haven't reached the bottom, continue scrolling
                if (currentScroll + clientHeight < scrollHeight) {
                    const scrollAmount = Math.min(500, scrollHeight - currentScroll - clientHeight + 100);
                    if (isElementScroll) {
                        contentScrollRegion.scrollBy(0, scrollAmount);
                    } else {
                        window.scrollBy(0, scrollAmount);
                    }
                    setTimeout(scrollStep, 100); // Small delay to allow content to load
                } else {
                    // Wait a bit more for any final lazy loading, then return to original position
                    setTimeout(() => {
                        if (isElementScroll) {
                            contentScrollRegion.scrollTo(0, originalScrollTop);
                        } else {
                            window.scrollTo(0, originalScrollTop);
                        }
                        resolve();
                    }, 500);
                }
            };
            
            // Start scrolling
            scrollStep();
        });
    }

    // Check if a link element is in the footer section or navigation
    function getLinkLocation(element) {
        // Walk up the DOM tree to check if the element is inside a footer or nav
        let current = element;
        while (current && current !== document.body) {
            // Check if current element is a footer element or has contentinfo role
            if (current.tagName === 'FOOTER' || 
                current.getAttribute('role') === 'contentinfo') {
                return { isFooter: true, isNav: false };
            }
            
            current = current.parentElement;
        }
        
        // Check if the link element contains SharePoint navigation text spans
        const navTextSpan = element.querySelector('span.ms-HorizontalNavItem-linkText, span.ms-Nav-linkText');
        if (navTextSpan) {
            return { isFooter: false, isNav: true };
        }
        
        return { isFooter: false, isNav: false };
    }

    // Get all links on the page
    function getAllLinks() {
        const links = Array.from(document.querySelectorAll('a[href]'));
        return links.map(link => {
            const location = getLinkLocation(link);
            return {
                url: link.href,
                text: link.textContent.trim(),
                element: link,
                isFooter: location.isFooter,
                isNav: location.isNav
            };
        }).filter(link => link.url && !link.url.startsWith('javascript:'));
    }

    // Get team site links
    function getTeamLinks() {
        const allLinks = getAllLinks();
        const teamLinks = allLinks.filter(link => {
            const url = link.url.toLowerCase();
            return (url.includes('teams.microsoft.com') || 
                   (url.includes('sharepoint.com') && url.includes('/teams/')));
        });
        
        return teamLinks.map(link => ({
            text: link.text || 'No text',
            url: link.url,
            isFooter: link.isFooter,
            isNav: link.isNav
        }));
    }

    // Get OneDrive links
    function getOneDriveLinks() {
        const allLinks = getAllLinks();
        const oneDriveLinks = allLinks.filter(link => {
            const url = link.url.toLowerCase();
            return (url.includes('-my.sharepoint.com') || url.includes('/personal/'));
        });
        
        return oneDriveLinks.map(link => ({
            text: link.text || 'No text',
            url: link.url,
            isFooter: link.isFooter,
            isNav: link.isNav
        }));
    }

    // Get email links (mailto links)
    function getEmailLinks() {
        const allLinks = getAllLinks();
        const emailLinks = allLinks.filter(link => {
            try {
                const url = new URL(link.url);
                return url.protocol === 'mailto:';
            } catch (error) {
                return false;
            }
        });
        
        return emailLinks.map(link => ({
            text: link.text || 'No text',
            url: link.url,
            status: 'mailto',
            statusText: 'Email link',
            isFooter: link.isFooter,
            isNav: link.isNav,
            isMailto: true
        }));
    }

    // Get external links (excluding Microsoft products and email links)
    function getExternalLinks() {
        // Microsoft product domains to exclude
        const microsoftDomains = [
            'microsoft.com',
            'sharepoint.com',
            'sharepoint-df.com',
            'office.com',
            'office365.com',
            'teams.microsoft.com',
            'outlook.com',
            'outlook.office.com',
            'onedrive.com',
            'live.com',
            'hotmail.com',
            'msn.com',
            'bing.com',
            'azure.com',
            'azurewebsites.net',
            'microsoftonline.com',
            'graph.microsoft.com',
            'powerapps.com',
            'powerbi.com',
            'dynamics.com',
            'xbox.com',
            'skype.com',
            'linkedin.com',
            'github.com',
            'visualstudio.com',
            'vscode.dev',
            'aka.ms',
            'microsoftstore.com'
        ];

        const allLinks = getAllLinks();
        const currentHostname = window.location.hostname.toLowerCase();
        
        const externalLinks = allLinks.filter(link => {
            try {
                const url = new URL(link.url);
                const hostname = url.hostname.toLowerCase();
                
                // Exclude mailto links (they are handled separately)
                if (url.protocol === 'mailto:') {
                    return false;
                }
                
                // Skip internal links (same domain)
                if (hostname === currentHostname) {
                    return false;
                }
                
                // Skip Microsoft product domains
                const isMicrosoftProduct = microsoftDomains.some(domain => 
                    hostname === domain || hostname.endsWith('.' + domain)
                );
                
                if (isMicrosoftProduct) {
                    return false;
                }
                
                // Skip tel and other non-http protocols
                if (!url.protocol.startsWith('http')) {
                    return false;
                }
                
                return true;
            } catch (error) {
                // Invalid URL, skip it
                return false;
            }
        });
        
        return externalLinks.map(link => ({
            text: link.text || 'No text',
            url: link.url,
            status: null, // Will be populated by checkExternalLinksStatus
            statusText: 'Checking...',
            isFooter: link.isFooter,
            isNav: link.isNav,
            isMailto: false
        }));
    }

    // Check HTTP status for external links using background script
    async function checkExternalLinksStatus(externalLinks) {
        const results = await Promise.allSettled(
            externalLinks.map(async (link) => {
                try {
                    const response = await chrome.runtime.sendMessage({
                        type: 'CHECK_LINK_STATUS',
                        url: link.url
                    });
                    
                    return {
                        ...link,
                        status: response.status,
                        statusText: getStatusText(response.status)
                    };
                } catch (error) {
                    return {
                        ...link,
                        status: 0,
                        statusText: 'Network Error'
                    };
                }
            })
        );
        
        return results.map(result => 
            result.status === 'fulfilled' ? result.value : {
                ...externalLinks[results.indexOf(result)],
                status: 0,
                statusText: 'Request Failed'
            }
        );
    }

    // Get HTTP status text
    function getStatusText(status) {
        // Handle string statuses from background script
        if (typeof status === 'string') {
            if (status === 'Error') return 'Network Error';
            if (status === 'Timeout') return 'Request Timeout';
            return status;
        }
        
        const statusMap = {
            200: '200 OK',
            201: '201 Created',
            301: '301 Moved Permanently',
            302: '302 Found',
            304: '304 Not Modified',
            400: '400 Bad Request',
            401: '401 Unauthorized',
            403: '403 Forbidden',
            404: '404 Not Found',
            410: '410 Gone',
            429: '429 Too Many Requests',
            500: '500 Internal Server Error',
            502: '502 Bad Gateway',
            503: '503 Service Unavailable',
            504: '504 Gateway Timeout',
            0: 'Network Error'
        };
        
        return statusMap[status] || `${status} Unknown`;
    }

    // Sort external links by HTTP status severity (broken links first)
    function sortLinksBySeverity(links) {
        const severityOrder = {
            // Network/Unknown errors (highest priority - show first)
            0: 1,
            'Error': 1,
            'Timeout': 1,
            // Server Error (5xx)
            500: 2, 502: 2, 503: 2, 504: 2,
            // Client Error (4xx)  
            400: 3, 401: 3, 403: 3, 404: 3, 410: 3, 429: 3,
            // Redirection (3xx)
            301: 4, 302: 4, 304: 4, 307: 4, 308: 4,
            // Success (2xx) (lowest priority - show last)
            200: 5, 201: 5, 202: 5, 204: 5,
            // Email links (same priority as success)
            'mailto': 5
        };
        
        return links.sort((a, b) => {
            const severityA = severityOrder[a.status] || 6; // Unknown status gets lowest priority
            const severityB = severityOrder[b.status] || 6;
            
            if (severityA !== severityB) {
                return severityA - severityB; // Lower severity number = higher priority (show first)
            }
            
            // If same severity, sort by status code
            if (typeof a.status === 'number' && typeof b.status === 'number') {
                return a.status - b.status;
            }
            
            // If one or both are strings, sort alphabetically
            return String(a.status).localeCompare(String(b.status));
        });
    }

    // Get SharePoint site information
    function getSharePointSiteInfo() {
        const siteTitle = document.getElementById('SiteHeaderTitle')?.textContent?.trim() || 'Unknown Site';
        const pageTitle = document.title || 'Unknown Page';
        
        // Clean URL - remove everything after .aspx
        let cleanUrl = window.location.href;
        const aspxIndex = cleanUrl.indexOf('.aspx');
        if (aspxIndex !== -1) {
            cleanUrl = cleanUrl.substring(0, aspxIndex + 5); // Include .aspx
        }
        
        return {
            siteTitle,
            pageTitle,
            cleanUrl,
            fullUrl: window.location.href
        };
    }

    // Highlight functionality for link elements
    let currentHighlightedElement = null;
    
    function highlightElement(url) {
        // Remove previous highlight
        removeHighlight();
        
        // Try multiple strategies to find the matching link element
        let linkElement = null;
        
        // Strategy 1: Exact URL match
        linkElement = document.querySelector(`a[href="${url}"]`);
        
        // Strategy 2: Try without fragment (hash)
        if (!linkElement) {
            const urlWithoutFragment = url.split('#')[0];
            linkElement = document.querySelector(`a[href="${urlWithoutFragment}"]`);
        }
        
        // Strategy 3: Try without query parameters
        if (!linkElement) {
            const urlWithoutQuery = url.split('?')[0];
            linkElement = document.querySelector(`a[href="${urlWithoutQuery}"]`);
        }
        
        // Strategy 4: Find all links and match by comparing URL objects
        if (!linkElement) {
            const allLinks = document.querySelectorAll('a[href]');
            for (const link of allLinks) {
                try {
                    const linkUrl = new URL(link.href);
                    const targetUrl = new URL(url);
                    
                    // Compare pathname and hostname
                    if (linkUrl.hostname === targetUrl.hostname && linkUrl.pathname === targetUrl.pathname) {
                        linkElement = link;
                        break;
                    }
                } catch (error) {
                    // Skip invalid URLs
                    continue;
                }
            }
        }
        
        // Strategy 5: Last resort - find by text content if URL contains text
        if (!linkElement) {
            const allLinks = document.querySelectorAll('a[href]');
            for (const link of allLinks) {
                if (link.href === url) {
                    linkElement = link;
                    break;
                }
            }
        }
        
        if (linkElement) {
            // Store reference to highlighted element
            currentHighlightedElement = linkElement;
            
            // Add highlight styles - bright red and thicker
            linkElement.style.outline = '5px solid #ff0000';
            linkElement.style.outlineOffset = '3px';
            linkElement.style.backgroundColor = 'rgba(255, 0, 0, 0.15)';
            linkElement.style.borderRadius = '4px';
            linkElement.style.transition = 'all 0.2s ease-in-out';
            linkElement.style.zIndex = '9999';
            
            // Scroll element into view if not visible
            linkElement.scrollIntoView({ 
                behavior: 'smooth', 
                block: 'center',
                inline: 'nearest'
            });
            
            console.log('Highlighted element:', linkElement, 'for URL:', url);
        } else {
            console.log('Could not find element for URL:', url);
        }
    }
    
    function removeHighlight() {
        if (currentHighlightedElement) {
            // Remove highlight styles
            currentHighlightedElement.style.outline = '';
            currentHighlightedElement.style.outlineOffset = '';
            currentHighlightedElement.style.backgroundColor = '';
            currentHighlightedElement.style.borderRadius = '';
            currentHighlightedElement.style.transition = '';
            currentHighlightedElement = null;
        }
    }

    // Listen for messages from the sidebar
    chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
        if (message.type === 'PING') {
            sendResponse({ status: 'ready' });
        } else if (message.type === 'CHECK_SHAREPOINT_SITE') {
            const siteInfo = isSharePointSite() ? getSharePointSiteInfo() : null;
            sendResponse({
                isSharePoint: isSharePointSite(),
                url: window.location.href,
                title: document.title,
                siteInfo: siteInfo
            });
        } else if (message.type === 'HIGHLIGHT_LINK') {
            highlightElement(message.url);
            sendResponse({ success: true });
        } else if (message.type === 'REMOVE_HIGHLIGHT') {
            removeHighlight();
            sendResponse({ success: true });
        } else if (message.type === 'GET_LINKS') {
            // Scroll to load all content first, then get links
            scrollToLoadAllContent().then(() => {
                const links = getAllLinks();
                sendResponse({ links: links });
            });
            return true; // Keep the message channel open for async response
        } else if (message.type === 'GET_TEAM_LINKS') {
            // Scroll to load all content first, then get team links
            scrollToLoadAllContent().then(() => {
                const teamLinks = getTeamLinks();
                sendResponse({ teamLinks: teamLinks });
            });
            return true; // Keep the message channel open for async response
        } else if (message.type === 'GET_ONEDRIVE_LINKS') {
            // Scroll to load all content first, then get OneDrive links
            scrollToLoadAllContent().then(() => {
                const oneDriveLinks = getOneDriveLinks();
                sendResponse({ oneDriveLinks: oneDriveLinks });
            });
            return true; // Keep the message channel open for async response
        } else if (message.type === 'GET_EXTERNAL_LINKS') {
            // Scroll to load all content first, then get external links
            scrollToLoadAllContent().then(() => {
                const externalLinks = getExternalLinks();
                sendResponse({ externalLinks: externalLinks });
            });
            return true; // Keep the message channel open for async response
        } else if (message.type === 'GET_EMAIL_LINKS') {
            // Scroll to load all content first, then get email links
            scrollToLoadAllContent().then(() => {
                const emailLinks = getEmailLinks();
                sendResponse({ emailLinks: emailLinks });
            });
            return true; // Keep the message channel open for async response
        } else if (message.type === 'CHECK_EXTERNAL_LINKS_STATUS') {
            checkExternalLinksStatus(message.links).then(checkedLinks => {
                const sortedLinks = sortLinksBySeverity(checkedLinks);
                sendResponse({ externalLinks: sortedLinks });
            });
            return true; // Keep the message channel open for async response
        }
    });

    // Send initial site check when content script loads
    if (isSharePointSite()) {
        const siteInfo = getSharePointSiteInfo();
        chrome.runtime.sendMessage({
            type: 'CONTENT_SHAREPOINT_DETECTED',
            url: window.location.href,
            title: document.title,
            siteInfo: siteInfo
        }).catch(() => {
            // Extension might not be ready, ignore
        });
    }
})();
