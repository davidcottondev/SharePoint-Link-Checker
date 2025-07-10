// Settings page JavaScript functionality
(function() {
    'use strict';

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
    const saveBtn = document.getElementById('save-btn');
    const teamsEnabled = document.getElementById('teams-enabled');
    const onedriveEnabled = document.getElementById('onedrive-enabled');
    const externalEnabled = document.getElementById('external-enabled');
    const emailEnabled = document.getElementById('email-enabled');
    const externalSettings = document.getElementById('external-settings');
    const status2xx = document.getElementById('status-2xx');
    const status3xx = document.getElementById('status-3xx');
    const status4xx = document.getElementById('status-4xx');
    const status5xx = document.getElementById('status-5xx');
    const statusNetwork = document.getElementById('status-network');

    // Initialize settings page
    async function init() {
        await loadSettings();
        setupEventListeners();
        updateExternalSettingsState();
    }

    // Load settings from storage
    async function loadSettings() {
        try {
            const result = await chrome.storage.sync.get(['linkCheckerSettings']);
            const settings = result.linkCheckerSettings || defaultSettings;
            
            // Set link type checkboxes
            teamsEnabled.checked = settings.linkTypes?.teams ?? true;
            onedriveEnabled.checked = settings.linkTypes?.onedrive ?? true;
            externalEnabled.checked = settings.linkTypes?.external ?? true;
            emailEnabled.checked = settings.linkTypes?.email ?? true;
            
            // Set status code checkboxes
            status2xx.checked = settings.externalStatusCodes?.['2xx'] ?? true;
            status3xx.checked = settings.externalStatusCodes?.['3xx'] ?? true;
            status4xx.checked = settings.externalStatusCodes?.['4xx'] ?? true;
            status5xx.checked = settings.externalStatusCodes?.['5xx'] ?? true;
            statusNetwork.checked = settings.externalStatusCodes?.['network'] ?? true;
            
        } catch (error) {
            console.error('Error loading settings:', error);
            // Use defaults if loading fails
        }
    }

    // Save settings to storage
    async function saveSettings() {
        const settings = {
            linkTypes: {
                teams: teamsEnabled.checked,
                onedrive: onedriveEnabled.checked,
                external: externalEnabled.checked,
                email: emailEnabled.checked
            },
            externalStatusCodes: {
                '2xx': status2xx.checked,
                '3xx': status3xx.checked,
                '4xx': status4xx.checked,
                '5xx': status5xx.checked,
                'network': statusNetwork.checked
            }
        };

        try {
            await chrome.storage.sync.set({ linkCheckerSettings: settings });
            
            // Show success feedback briefly
            const originalText = saveBtn.innerHTML;
            saveBtn.innerHTML = '<span class="ms-Button-label">✅ Settings Saved!</span>';
            saveBtn.disabled = true;
            
            // Close the settings page after a brief delay
            setTimeout(() => {
                window.close();
            }, 1000);
            
        } catch (error) {
            console.error('Error saving settings:', error);
            
            // Show error feedback
            const originalText = saveBtn.innerHTML;
            saveBtn.innerHTML = '<span class="ms-Button-label">❌ Save Failed</span>';
            saveBtn.disabled = true;
            
            setTimeout(() => {
                saveBtn.innerHTML = originalText;
                saveBtn.disabled = false;
            }, 2000);
        }
    }

    // Update external settings visibility
    function updateExternalSettingsState() {
        if (externalEnabled.checked) {
            externalSettings.classList.remove('disabled');
        } else {
            externalSettings.classList.add('disabled');
        }
    }

    // Setup event listeners
    function setupEventListeners() {
        // Save button
        saveBtn.addEventListener('click', saveSettings);

        // External link checkbox
        externalEnabled.addEventListener('change', updateExternalSettingsState);

        // Auto-save indicator on change
        const allCheckboxes = [
            teamsEnabled, onedriveEnabled, externalEnabled, emailEnabled,
            status2xx, status3xx, status4xx, status5xx, statusNetwork
        ];
        
        allCheckboxes.forEach(checkbox => {
            checkbox.addEventListener('change', () => {
                // Visual feedback that settings have changed
                saveBtn.style.backgroundColor = '#0078d4';
                saveBtn.innerHTML = '<span class="ms-Button-label">💾 Save Settings & Return *</span>';
            });
        });
    }

    // Initialize when DOM is loaded
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', init);
    } else {
        init();
    }
})();
