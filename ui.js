// ui.js - UI Module for SharePoint & OneDrive Scanner v3.0
// Handles DOM manipulation, rendering, event handlers, tab management, and UI updates

// TAB MANAGEMENT
function initializeTabs() {
    const tabButtons = document.querySelectorAll('.tab-button');
    const tabContents = document.querySelectorAll('.tab-content');

    tabButtons.forEach(button => {
        button.addEventListener('click', () => {
            const targetTab = button.getAttribute('data-tab');
            
            tabButtons.forEach(btn => btn.classList.remove('active'));
            button.classList.add('active');
            
            tabContents.forEach(content => {
                content.classList.remove('active');
                if (content.id === targetTab + '-tab') {
                    content.classList.add('active');
                }
            });

            // Clear results when switching tabs
            if (window.configModule && window.configModule.clearResults) {
                window.configModule.clearResults();
            }
            
            if (window.configModule && window.configModule.showToast) {
                window.configModule.showToast(`Switched to ${targetTab === 'sharepoint' ? 'SharePoint' : 'OneDrive'} - results cleared`);
            }
        });
    });
}

// ENHANCED SCAN CONTROLS INITIALIZATION
function initializeScanControls() {
    // SharePoint scan controls
    const spSharingControls = document.querySelectorAll('#sharepoint-tab .toggle-btn[data-filter]');
    const spScopeControls = document.querySelectorAll('#sharepoint-tab .toggle-btn[data-scope]');
    
    // OneDrive scan controls
    const odSharingControls = document.querySelectorAll('#onedrive-tab .toggle-btn[data-filter]');
    const odScopeControls = document.querySelectorAll('#onedrive-tab .toggle-btn[data-scope]');
    
    // Setup sharing filter controls
    [...spSharingControls, ...odSharingControls].forEach(btn => {
        btn.addEventListener('click', () => {
            const filterType = btn.dataset.filter;
            const parentTab = btn.closest('.tab-content');
            
            // Update active state within the same tab
            parentTab.querySelectorAll('.toggle-btn[data-filter]').forEach(b => b.classList.remove('active'));
            btn.classList.add('active');
            
            if (window.configModule && window.configModule.updateScanSettings) {
                window.configModule.updateScanSettings({ sharingFilter: filterType });
            }
            
            console.log(`Sharing filter updated to: ${filterType}`);
            
            if (window.configModule && window.configModule.showToast) {
                window.configModule.showToast(`Sharing filter: ${filterType}`);
            }
        });
    });
    
    // Setup content scope controls
    [...spScopeControls, ...odScopeControls].forEach(btn => {
        btn.addEventListener('click', () => {
            const scopeType = btn.dataset.scope;
            const parentTab = btn.closest('.tab-content');
            
            // Update active state within the same tab
            parentTab.querySelectorAll('.toggle-btn[data-scope]').forEach(b => b.classList.remove('active'));
            btn.classList.add('active');
            
            if (window.configModule && window.configModule.updateScanSettings) {
                window.configModule.updateScanSettings({ contentScope: scopeType });
            }
            
            console.log(`Content scope updated to: ${scopeType}`);
            
            if (window.configModule && window.configModule.showToast) {
                window.configModule.showToast(`Content scope: ${scopeType === 'folders' ? 'Folders Only' : 'All Content'}`);
            }
        });
    });
}

// SITES RENDERING AND MANAGEMENT
function renderSites() {
    const container = document.getElementById('sites-container');
    const sitesCount = document.getElementById('sites-count');
    
    if (!window.configModule) {
        console.error('configModule not available');
        return;
    }
    
    const sites = window.configModule.sites;
    const selectedSiteIds = window.configModule.selectedSiteIds;
    
    container.innerHTML = '';
    
    if (sites.length === 0) {
        container.innerHTML = '<div class="empty-state"><p>No sites found</p></div>';
        sitesCount.innerText = 'No sites loaded';
        return;
    }

    sitesCount.innerText = `${sites.length} sites found`;
    sitesCount.className = 'status-badge status-approved';

    sites.forEach(site => {
        const siteItem = document.createElement('div');
        siteItem.className = 'site-item';

        const checkbox = document.createElement('input');
        checkbox.type = 'checkbox';
        checkbox.addEventListener('change', (e) => {
            if (e.target.checked) {
                selectedSiteIds.add(site.id);
            } else {
                selectedSiteIds.delete(site.id);
            }
            updateScanButton();
        });

        const siteInfo = document.createElement('div');
        siteInfo.className = 'site-info';
        
        const siteName = document.createElement('h3');
        siteName.className = 'site-name';
        siteName.innerText = site.name || site.displayName || 'Unnamed Site';
        
        const siteUrl = document.createElement('p');
        siteUrl.className = 'site-url';
        siteUrl.innerText = site.webUrl || '';

        siteInfo.appendChild(siteName);
        siteInfo.appendChild(siteUrl);
        siteItem.appendChild(checkbox);
        siteItem.appendChild(siteInfo);
        container.appendChild(siteItem);
    });

    document.getElementById('select-all-sites').disabled = false;
    document.getElementById('deselect-all-sites').disabled = false;
}

function renderUsers() {
    const container = document.getElementById('users-container');
    const usersCount = document.getElementById('users-count');
    
    if (!window.configModule) {
        console.error('configModule not available');
        return;
    }
    
    const users = window.configModule.users;
    const selectedUserIds = window.configModule.selectedUserIds;
    
    container.innerHTML = '';
    
    if (users.length === 0) {
        container.innerHTML = '<div class="empty-state"><p>No users found</p></div>';
        usersCount.innerText = 'No users loaded';
        return;
    }

    usersCount.innerText = `${users.length} users found`;
    usersCount.className = 'status-badge status-approved';

    users.forEach(user => {
        const userItem = document.createElement('div');
        userItem.className = 'user-item';

        const checkbox = document.createElement('input');
        checkbox.type = 'checkbox';
        checkbox.addEventListener('change', (e) => {
            if (e.target.checked) {
                selectedUserIds.add(user.id);
            } else {
                selectedUserIds.delete(user.id);
            }
            updateUserButtons();
        });

        const userInfo = document.createElement('div');
        userInfo.className = 'user-info';
        
        const userName = document.createElement('h3');
        userName.className = 'user-name';
        userName.innerText = user.displayName || user.userPrincipalName || 'Unknown User';
        
        const userEmail = document.createElement('p');
        userEmail.className = 'user-email';
        userEmail.innerText = user.userPrincipalName || user.mail || '';

        userInfo.appendChild(userName);
        userInfo.appendChild(userEmail);
        userItem.appendChild(checkbox);
        userItem.appendChild(userInfo);
        container.appendChild(userItem);
    });

    document.getElementById('select-all-users').disabled = false;
    document.getElementById('deselect-all-users').disabled = false;
}

function updateScanButton() {
    const scanBtn = document.getElementById('scan-sharepoint-btn');
    const selectedSiteIds = window.configModule ? window.configModule.selectedSiteIds : new Set();
    const scanning = window.configModule ? window.configModule.scanning : false;
    
    if (scanBtn) {
        scanBtn.disabled = selectedSiteIds.size === 0 || scanning;
    }
}

function updateUserButtons() {
    const scanBtn = document.getElementById('scan-onedrive-btn');
    const selectedUserIds = window.configModule ? window.configModule.selectedUserIds : new Set();
    const scanning = window.configModule ? window.configModule.scanning : false;
    
    if (scanBtn) {
        scanBtn.disabled = selectedUserIds.size === 0 || scanning;
    }
}

function updateCheckboxes() {
    const sites = window.configModule ? window.configModule.sites : [];
    const selectedSiteIds = window.configModule ? window.configModule.selectedSiteIds : new Set();
    
    document.querySelectorAll('.site-item input[type="checkbox"]').forEach((checkbox, index) => {
        if (sites[index]) {
            checkbox.checked = selectedSiteIds.has(sites[index].id);
        }
    });
}

function updateUserCheckboxes() {
    const users = window.configModule ? window.configModule.users : [];
    const selectedUserIds = window.configModule ? window.configModule.selectedUserIds : new Set();
    
    document.querySelectorAll('.user-item input[type="checkbox"]').forEach((checkbox, index) => {
        if (users[index]) {
            checkbox.checked = selectedUserIds.has(users[index].id);
        }
    });
}

// DISCOVERY EVENT HANDLERS
function initializeDiscoveryHandlers() {
    // SharePoint site discovery
    const findSitesBtn = document.getElementById('find-sites');
    if (findSitesBtn) {
        findSitesBtn.addEventListener('click', async function() {
            this.disabled = true;
            this.innerText = 'Loading...';

            try {
                if (!window.apiModule || !window.apiModule.discoverSharePointSites) {
                    throw new Error('API module not available');
                }
                
                const sites = await window.apiModule.discoverSharePointSites();
                
                if (window.configModule) {
                    window.configModule.sites = sites;
                }
                
                renderSites();
                
                if (window.configModule && window.configModule.showToast) {
                    window.configModule.showToast(`Found ${sites.length} sites`);
                }

            } catch (error) {
                console.error('Error fetching sites:', error);
                alert('Failed to fetch sites: ' + error.message);
            } finally {
                this.disabled = false;
                this.innerText = 'Discover Sites';
            }
        });
    }

    // OneDrive user discovery
    const discoverUsersBtn = document.getElementById('discover-users');
    if (discoverUsersBtn) {
        discoverUsersBtn.addEventListener('click', async function() {
            this.disabled = true;
            this.innerText = 'Loading...';

            try {
                if (!window.apiModule || !window.apiModule.discoverOneDriveUsers) {
                    throw new Error('API module not available');
                }
                
                const users = await window.apiModule.discoverOneDriveUsers();
                
                if (window.configModule) {
                    window.configModule.users = users;
                }
                
                renderUsers();
                
                if (window.configModule && window.configModule.showToast) {
                    window.configModule.showToast(`Found ${users.length} users`);
                }

            } catch (error) {
                console.error('Error discovering users:', error);
                alert('Failed to discover users: ' + error.message);
            } finally {
                this.disabled = false;
                this.innerText = 'Discover Users';
            }
        });
    }
}

// SELECTION HANDLERS
function initializeSelectionHandlers() {
    // Sites selection handlers
    const selectAllSitesBtn = document.getElementById('select-all-sites');
    if (selectAllSitesBtn) {
        selectAllSitesBtn.addEventListener('click', function() {
            const sites = window.configModule ? window.configModule.sites : [];
            const selectedSiteIds = window.configModule ? window.configModule.selectedSiteIds : new Set();
            
            selectedSiteIds.clear();
            sites.forEach(site => selectedSiteIds.add(site.id));
            updateCheckboxes();
            updateScanButton();
            
            if (window.configModule && window.configModule.showToast) {
                window.configModule.showToast(`Selected all ${sites.length} sites`);
            }
        });
    }

    const deselectAllSitesBtn = document.getElementById('deselect-all-sites');
    if (deselectAllSitesBtn) {
        deselectAllSitesBtn.addEventListener('click', function() {
            const selectedSiteIds = window.configModule ? window.configModule.selectedSiteIds : new Set();
            
            selectedSiteIds.clear();
            updateCheckboxes();
            updateScanButton();
            
            if (window.configModule && window.configModule.showToast) {
                window.configModule.showToast('Deselected all sites');
            }
        });
    }

    // Users selection handlers
    const selectAllUsersBtn = document.getElementById('select-all-users');
    if (selectAllUsersBtn) {
        selectAllUsersBtn.addEventListener('click', function() {
            const users = window.configModule ? window.configModule.users : [];
            const selectedUserIds = window.configModule ? window.configModule.selectedUserIds : new Set();
            
            selectedUserIds.clear();
            users.forEach(user => selectedUserIds.add(user.id));
            updateUserCheckboxes();
            updateUserButtons();
            
            if (window.configModule && window.configModule.showToast) {
                window.configModule.showToast(`Selected all ${users.length} users`);
            }
        });
    }

    const deselectAllUsersBtn = document.getElementById('deselect-all-users');
    if (deselectAllUsersBtn) {
        deselectAllUsersBtn.addEventListener('click', function() {
            const selectedUserIds = window.configModule ? window.configModule.selectedUserIds : new Set();
            
            selectedUserIds.clear();
            updateUserCheckboxes();
            updateUserButtons();
            
            if (window.configModule && window.configModule.showToast) {
                window.configModule.showToast('Deselected all users');
            }
        });
    }
}

// SCAN CONTROL HANDLERS
function initializeScanHandlers() {
    // SharePoint scan button
    const scanSharePointBtn = document.getElementById('scan-sharepoint-btn');
    if (scanSharePointBtn) {
        scanSharePointBtn.addEventListener('click', function() {
            if (window.scanningModule && window.scanningModule.scanSharePointSites) {
                window.scanningModule.scanSharePointSites();
            } else {
                console.error('Scanning module not available');
                alert('Scanning module not loaded. Please refresh the page.');
            }
        });
    }

    // OneDrive scan button
    const scanOneDriveBtn = document.getElementById('scan-onedrive-btn');
    if (scanOneDriveBtn) {
        scanOneDriveBtn.addEventListener('click', function() {
            if (window.scanningModule && window.scanningModule.scanOneDriveUsers) {
                window.scanningModule.scanOneDriveUsers();
            } else {
                console.error('Scanning module not available');
                alert('Scanning module not loaded. Please refresh the page.');
            }
        });
    }

    // Stop scan buttons
    const stopSharePointBtn = document.getElementById('stop-sharepoint-btn');
    if (stopSharePointBtn) {
        stopSharePointBtn.addEventListener('click', function() {
            console.log('🛑 STOP SHAREPOINT SCAN requested by user');
            
            const controller = window.configModule ? window.configModule.controller : null;
            if (controller) {
                controller.stop = true;
            }
            
            if (window.configModule && window.configModule.showToast) {
                window.configModule.showToast('⏹️ Stopping SharePoint scan...', 2000);
            }
            
            // Update UI immediately
            this.disabled = true;
            this.innerText = 'Stopping...';
            
            // Reset scan button
            const scanBtn = document.getElementById('scan-sharepoint-btn');
            if (scanBtn) {
                scanBtn.innerText = 'Scan Selected Sites';
            }
            
            console.log('✅ SharePoint scan stop signal sent');
        });
    }

    const stopOneDriveBtn = document.getElementById('stop-onedrive-btn');
    if (stopOneDriveBtn) {
        stopOneDriveBtn.addEventListener('click', function() {
            console.log('🛑 STOP ONEDRIVE SCAN requested by user');
            
            const controller = window.configModule ? window.configModule.controller : null;
            if (controller) {
                controller.stop = true;
            }
            
            if (window.configModule && window.configModule.showToast) {
                window.configModule.showToast('⏹️ Stopping OneDrive scan...', 2000);
            }
            
            // Update UI immediately
            this.disabled = true;
            this.innerText = 'Stopping...';
            
            // Reset scan button
            const scanBtn = document.getElementById('scan-onedrive-btn');
            if (scanBtn) {
                scanBtn.innerText = 'Scan Selected Users';
            }
            
            console.log('✅ OneDrive scan stop signal sent');
        });
    }
}

// PROGRESS UI MANAGEMENT
function initializeProgressUI() {
    // Progress sections are managed by the scanning module
    // This function ensures progress elements exist and are properly initialized
    const progressSections = [
        'sharepoint-progress-section',
        'onedrive-progress-section'
    ];
    
    progressSections.forEach(sectionId => {
        const section = document.getElementById(sectionId);
        if (section) {
            section.style.display = 'none'; // Hidden by default
        }
    });
}

// BUTTON STATE MANAGEMENT
function updateButtonStates(scanning = false) {
    const scanButtons = [
        'scan-sharepoint-btn',
        'scan-onedrive-btn'
    ];
    
    const stopButtons = [
        'stop-sharepoint-btn',
        'stop-onedrive-btn'
    ];
    
    scanButtons.forEach(btnId => {
        const btn = document.getElementById(btnId);
        if (btn) {
            if (scanning) {
                btn.disabled = true;
                btn.innerText = btn.id.includes('sharepoint') ? 'Scanning...' : 'Scanning...';
            } else {
                const selectedIds = btn.id.includes('sharepoint') ? 
                    (window.configModule ? window.configModule.selectedSiteIds : new Set()) :
                    (window.configModule ? window.configModule.selectedUserIds : new Set());
                
                btn.disabled = selectedIds.size === 0;
                btn.innerText = btn.id.includes('sharepoint') ? 'Scan Selected Sites' : 'Scan Selected Users';
            }
        }
    });
    
    stopButtons.forEach(btnId => {
        const btn = document.getElementById(btnId);
        if (btn) {
            btn.disabled = !scanning;
            btn.innerText = 'Stop Scan';
        }
    });
}

// CSV PREVIEW DISPLAY
function displayCSVPreview(data) {
    const previewSection = document.getElementById('csv-preview-section');
    const previewContainer = document.getElementById('csv-preview');
    
    if (!data || data.length === 0) {
        if (previewSection) {
            previewSection.style.display = 'none';
        }
        return;
    }
    
    const headers = Object.keys(data[0]);
    
    let tableHTML = '<table><thead><tr>';
    headers.forEach(header => {
        tableHTML += `<th>${header}</th>`;
    });
    tableHTML += '</tr></thead><tbody>';
    
    data.slice(0, 10).forEach(row => { // Show first 10 rows
        tableHTML += '<tr>';
        headers.forEach(header => {
            tableHTML += `<td>${row[header] || ''}</td>`;
        });
        tableHTML += '</tr>';
    });
    
    if (data.length > 10) {
        tableHTML += `<tr><td colspan="${headers.length}" style="text-align: center; font-style: italic;">... and ${data.length - 10} more rows</td></tr>`;
    }
    
    tableHTML += '</tbody></table>';
    
    if (previewContainer) {
        previewContainer.innerHTML = tableHTML;
    }
    
    if (previewSection) {
        previewSection.style.display = 'block';
    }
}

// INITIALIZE ALL UI COMPONENTS
function initializeUI() {
    console.log('🎨 Initializing UI components...');
    
    try {
        initializeTabs();
        initializeScanControls();
        initializeDiscoveryHandlers();
        initializeSelectionHandlers();
        initializeScanHandlers();
        initializeProgressUI();
        
        console.log('✅ UI components initialized successfully');
    } catch (error) {
        console.error('❌ Error initializing UI components:', error);
    }
}

// Export functions for use in other modules
window.uiModule = {
    // Tab Management
    initializeTabs,
    
    // Scan Controls
    initializeScanControls,
    
    // Rendering
    renderSites,
    renderUsers,
    
    // Button Updates
    updateScanButton,
    updateUserButtons,
    updateCheckboxes,
    updateUserCheckboxes,
    updateButtonStates,
    
    // Event Handlers
    initializeDiscoveryHandlers,
    initializeSelectionHandlers,
    initializeScanHandlers,
    
    // Progress UI
    initializeProgressUI,
    
    // CSV Display
    displayCSVPreview,
    
    // Main Initialization
    initializeUI
};
