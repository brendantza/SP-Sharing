// config.js - Configuration and Utilities Module for SharePoint & OneDrive Scanner v3.0
// Contains constants, utility functions, permission classification, and helper methods

// APPLICATION CONSTANTS
const SKIP_FOLDERS = [
    'Forms', 'SiteAssets', '_catalogs', 'Style Library', 'SitePages', 
    'Lists', 'PublishingImages', 'SiteCollectionImages', 'MasterPageGallery',
    '_themes', '_layouts', '_vti_', 'wpresources', 'ClientSideAssets'
];

// SHAREPOINT PRESERVATION HOLD LIBRARY PATTERNS
// These patterns identify SharePoint preservation hold libraries used for legal hold and compliance
const PRESERVATION_HOLD_PATTERNS = [
    'Preservation Hold Library',
    'Preservation Hold',
    'Hold Library',
    'PreservationHoldLibrary',
    'Legal Hold',
    'Compliance Hold',
    'eDiscovery Hold'
];

// DEFAULT SHAREPOINT GROUPS - Groups typically created automatically by SharePoint
const DEFAULT_SHAREPOINT_GROUPS = [
    'Company Administrator', 'company administrator',
    'Team Site Owners', 'team site owners', 'site owners',
    'Team site members', 'team site members', 'site members', 'members',
    'excel services viewers', 'Excel Services Viewers',
    'team site visitors', 'Team site visitors', 'site visitors', 'visitors',
    'Style Resource Readers', 'style resource readers',
    'Hierarchy Managers', 'hierarchy managers',
    'Quick Deploy Users', 'quick deploy users',
    'Restricted Readers', 'restricted readers',
    'Viewers', 'viewers',
    'Web Analytics Data Viewers', 'web analytics data viewers',
    'Site Collection Administrators', 'site collection administrators',
    'Site Collection Auditors', 'site collection auditors'
];

const APP_CONFIG = {
    name: "SharePoint & OneDrive Scanner Enhanced",
    version: "3.0.0",
    userAgent: "NONISV|YourCompany|SharePointOneDriveScanner/3.0.0",
    maxRetries: 3,
    batchSize: 5,
    throttleDelay: 500
};

// ENHANCED SCAN SETTINGS - UPDATED DEFAULTS FOR COMPREHENSIVE SCANNING
let scanSettings = {
    sharingFilter: 'all', // external, internal, all - DEFAULT: all sharing
    contentScope: 'all'    // folders, all - DEFAULT: all content (files + folders)
};

// APPLICATION STATE
let sites = [];
let users = [];
let selectedSiteIds = new Set();
let selectedUserIds = new Set();
let results = [];
let scanning = false;
let tenantDomains = new Set();
let controller = { stop: false };
let bulkCsvData = [];

// UTILITY FUNCTIONS
function showToast(msg, timeout = 3000) {
    const toast = document.getElementById('toast');
    if (toast) {
        toast.innerText = msg;
        toast.style.display = 'block';
        setTimeout(() => toast.style.display = 'none', timeout);
    }
}

function shouldSkipFolder(folderName) {
    if (!folderName) return true;
    if (folderName.startsWith('_') || folderName.startsWith('.')) return true;
    return SKIP_FOLDERS.some(skip => 
        folderName.toLowerCase().includes(skip.toLowerCase())
    );
}

// PERMISSION CLASSIFICATION FUNCTIONS
function isExternalUser(email, tenantDomains) {
    if (!email) return false;
    const emailDomain = email.toLowerCase().split('@')[1];
    if (!emailDomain) return false;
    
    for (const domain of tenantDomains) {
        if (emailDomain === domain.toLowerCase()) {
            return false;
        }
    }
    return true;
}

function isInternalUser(email, tenantDomains) {
    if (!email) return false;
    const emailDomain = email.toLowerCase().split('@')[1];
    if (!emailDomain) return false;
    
    for (const domain of tenantDomains) {
        if (emailDomain === domain.toLowerCase()) {
            return true;
        }
    }
    return false;
}

// ENHANCED PERMISSION CLASSIFICATION - FIXED TO ENFORCE CUSTOM INSTRUCTIONS AND HANDLE GROUPS
function classifyPermission(permission, tenantDomains) {
    console.log('🔍 CLASSIFYING PERMISSION:', permission);
    console.log('🎯 TENANT DOMAINS FOR COMPARISON:', Array.from(tenantDomains));
    
    let isExternal = false;
    let isInternal = false;
    let debugInfo = [];
    
    // Anonymous links are always external
    if (permission.link && permission.link.scope === 'anonymous') {
        console.log('✅ CLASSIFICATION: external (anonymous link)');
        return 'external';
    }
    
    // Organization links are internal
    if (permission.link && permission.link.scope === 'organization') {
        console.log('✅ CLASSIFICATION: internal (organization link)');
        return 'internal';
    }
    
    // Check for group permission in grantedToV2 - GROUPS ARE ALWAYS INTERNAL
    if (permission.grantedToV2 && permission.grantedToV2.group) {
        const group = permission.grantedToV2.group;
        debugInfo.push(`grantedToV2.group: ${group.displayName} -> INTERNAL (organizational group)`);
        console.log('✅ CLASSIFICATION: internal (group permission)');
        return 'internal';
    }
    
    // Check grantedTo user
    if (permission.grantedTo && permission.grantedTo.user && permission.grantedTo.user.email) {
        const email = permission.grantedTo.user.email;
        const isExt = isExternalUser(email, tenantDomains);
        debugInfo.push(`grantedTo: ${email} -> ${isExt ? 'EXTERNAL' : 'INTERNAL'}`);
        
        if (isExt) {
            isExternal = true;
        } else {
            isInternal = true;
        }
    }
    
    // Check grantedToIdentitiesV2
    if (Array.isArray(permission.grantedToIdentitiesV2)) {
        for (const g of permission.grantedToIdentitiesV2) {
            if (g.user && g.user.email) {
                const email = g.user.email;
                const isExt = isExternalUser(email, tenantDomains);
                debugInfo.push(`grantedToIdentitiesV2: ${email} -> ${isExt ? 'EXTERNAL' : 'INTERNAL'}`);
                
                if (isExt) {
                    isExternal = true;
                } else {
                    isInternal = true;
                }
            } else if (g.group) {
                // Groups in grantedToIdentitiesV2 are also internal
                debugInfo.push(`grantedToIdentitiesV2.group: ${g.group.displayName} -> INTERNAL (organizational group)`);
                isInternal = true;
            }
        }
    }
    
    // ❌ REMOVED INCORRECT ASSUMPTION: 'users' scope links can be shared with external users
    // The previous code incorrectly assumed permission.link.scope === 'users' meant internal only
    
    const result = isExternal && !isInternal ? 'external' : 
                  isInternal && !isExternal ? 'internal' : 
                  isExternal && isInternal ? 'mixed' : 'unknown';
    
    console.log(`🎯 CLASSIFICATION RESULT: ${result.toUpperCase()}`);
    if (debugInfo.length > 0) {
        console.log('📋 DEBUG INFO:', debugInfo);
    }
    
    return result;
}

// ENHANCED FILTERING BASED ON SCAN SETTINGS
function shouldIncludePermission(permission, tenantDomains, sharingFilter) {
    console.log('🔍 ANALYZING PERMISSION FOR INCLUSION:', JSON.stringify(permission, null, 2));
    
    // Check if this is a group permission first - INCLUDING SITE GROUPS
    const hasRegularGroup = permission.grantedToV2 && permission.grantedToV2.group;
    const hasSiteGroup = permission.grantedToV2 && permission.grantedToV2.siteGroup;
    const hasGroupPermission = hasRegularGroup || hasSiteGroup;
    
    console.log('🔍 GROUP PERMISSION CHECK:', {
        hasGrantedToV2: !!permission.grantedToV2,
        hasRegularGroup: hasRegularGroup,
        hasSiteGroup: hasSiteGroup,
        hasGroupPermission: hasGroupPermission,
        grantedToV2: permission.grantedToV2
    });
    
    // ⚠️ CRITICAL: Per custom instructions, direct grants to individual users (not groups) should NOT be shown
    // But group permissions (including site groups) should be included as they represent sharing configurations
    const isDirectUserGrant = (
        permission.grantedTo && 
        permission.grantedTo.user && 
        !permission.link && 
        !hasGroupPermission
    );
    
    console.log('🔍 DIRECT USER GRANT CHECK:', {
        hasGrantedTo: !!permission.grantedTo,
        hasGrantedToUser: !!(permission.grantedTo && permission.grantedTo.user),
        hasLink: !!permission.link,
        isDirectUserGrant: isDirectUserGrant
    });
    
    if (isDirectUserGrant) {
        console.log('🚫 EXCLUDING DIRECT USER GRANT per custom instructions:', permission);
        return false; // Exclude direct user grants only
    }
    
    // Include group permissions and link-based sharing
    if (hasGroupPermission) {
        console.log('✅ INCLUDING GROUP PERMISSION:', permission);
        console.log('📋 GROUP DETAILS:', permission.grantedToV2.group);
    }
    
    if (permission.link) {
        console.log('✅ INCLUDING LINK-BASED PERMISSION:', permission);
        console.log('🔗 LINK DETAILS:', permission.link);
    }
    
    const classification = classifyPermission(permission, tenantDomains);
    console.log('🎯 PERMISSION CLASSIFICATION:', classification);
    
    let includeBasedOnFilter = false;
    switch (sharingFilter) {
        case 'external':
            includeBasedOnFilter = classification === 'external' || classification === 'mixed';
            break;
        case 'internal':
            includeBasedOnFilter = classification === 'internal' || classification === 'mixed';
            break;
        case 'all':
            includeBasedOnFilter = true;
            break;
        default:
            includeBasedOnFilter = classification === 'external' || classification === 'mixed';
    }
    
    console.log('🎚️ FILTER DECISION:', {
        sharingFilter: sharingFilter,
        classification: classification,
        includeBasedOnFilter: includeBasedOnFilter,
        finalDecision: includeBasedOnFilter
    });
    
    return includeBasedOnFilter;
}

// PERMISSION EXTRACTION UTILITIES - FIXED FOR CORRECT API PROPERTIES
function extractUserFromPermission(p, tenantDomains) {
    let who = '';
    
    if (p.link) {
        if (p.link.scope === 'anonymous') {
            who = 'Anyone (Anonymous Link)';
        } else if (p.link.scope === 'organization') {
            who = 'Organization Link';
        } else {
            who = `Link (${p.link.scope || 'unknown scope'})`;
        }
    }
    
    // Check for group permission in grantedToV2 - INCLUDING SITE GROUPS
    if (p.grantedToV2 && p.grantedToV2.group) {
        const group = p.grantedToV2.group;
        who = group.displayName || group.email || '(group)';
        console.log('🔍 EXTRACTING GROUP NAME:', who, 'from', group);
        return who;
    }
    
    // Check for site group permission in grantedToV2.siteGroup
    if (p.grantedToV2 && p.grantedToV2.siteGroup) {
        const siteGroup = p.grantedToV2.siteGroup;
        who = siteGroup.displayName || siteGroup.loginName || '(site group)';
        console.log('🔍 EXTRACTING SITE GROUP NAME:', who, 'from', siteGroup);
        return who;
    }
    
    // Check for direct user grant (but this should be excluded by shouldIncludePermission)
    if (p.grantedTo && p.grantedTo.user && p.grantedTo.user.email) {
        const displayName = p.grantedTo.user.displayName;
        let email = p.grantedTo.user.email;
        
        if (displayName && displayName !== email) {
            who = `${displayName} (${email})`;
        } else {
            who = email;
        }
    }
    
    // Handle grantedToIdentitiesV2 if it exists (for older API versions or different structures)
    if (Array.isArray(p.grantedToIdentitiesV2) && p.grantedToIdentitiesV2.length > 0) {
        const parts = [];
        for (const g of p.grantedToIdentitiesV2) {
            if (g.user) {
                const displayName = g.user.displayName;
                let email = g.user.email;
                
                let userDisplay = '';
                if (displayName && email && displayName !== email) {
                    userDisplay = `${displayName} (${email})`;
                } else if (email) {
                    userDisplay = email;
                } else if (displayName) {
                    userDisplay = displayName;
                } else {
                    userDisplay = '(user)';
                }
                parts.push(userDisplay);
            } else if (g.group) {
                parts.push(g.group.displayName || g.group.email || '(group)');
            }
        }
        
        if (parts.length > 0) {
            if (who.includes('Link') && !who.includes('Anonymous')) {
                who = parts.join(', ');
            } else if (!who || who === '(direct grant)') {
                who = parts.join(', ');
            }
        }
    }
    
    if (!who) who = '(direct grant)';
    return who;
}

function extractExpirationDate(permission) {
    if (permission.expirationDateTime) {
        try {
            return new Date(permission.expirationDateTime).toLocaleDateString();
        } catch (error) {
            return permission.expirationDateTime;
        }
    }
    
    if (permission.link && permission.link.expirationDateTime) {
        try {
            return new Date(permission.link.expirationDateTime).toLocaleDateString();
        } catch (error) {
            return permission.link.expirationDateTime;
        }
    }
    
    return 'No expiration';
}

// STATE MANAGEMENT FUNCTIONS
function clearResults() {
    results = [];
    const resultsContainer = document.getElementById('results-container');
    if (resultsContainer) {
        resultsContainer.innerHTML = '<div class="empty-state"><p>No scan results yet. Configure scan options and run a scan to discover sharing.</p></div>';
    }
    
    const resultsActions = document.getElementById('results-actions');
    if (resultsActions) {
        resultsActions.style.display = 'none';
    }
    
    const sharingFilters = document.getElementById('sharing-filters');
    if (sharingFilters) {
        sharingFilters.style.display = 'none';
    }
    
    const bulkControls = document.getElementById('bulk-controls');
    if (bulkControls) {
        bulkControls.style.display = 'none';
    }
    
    // Update result count
    const resultCount = document.getElementById('result-count');
    if (resultCount) {
        resultCount.innerText = '0 found';
        resultCount.className = 'status-badge status-info';
    }
    
    // Disable export button
    const exportBtn = document.getElementById('export-btn');
    if (exportBtn) {
        exportBtn.disabled = true;
    }
}

function clearSitesAndUsers() {
    sites = [];
    users = [];
    selectedSiteIds.clear();
    selectedUserIds.clear();
    
    // Clear sites UI
    const sitesContainer = document.getElementById('sites-container');
    if (sitesContainer) {
        sitesContainer.innerHTML = '<div class="empty-state"><p>Click "Discover Sites" to load your SharePoint sites</p></div>';
    }
    
    const sitesCount = document.getElementById('sites-count');
    if (sitesCount) {
        sitesCount.innerText = 'No sites loaded';
        sitesCount.className = 'status-badge status-info';
    }
    
    // Clear users UI
    const usersContainer = document.getElementById('users-container');
    if (usersContainer) {
        usersContainer.innerHTML = '<div class="empty-state"><p>Click "Discover Users" to load users with OneDrive access</p></div>';
    }
    
    const usersCount = document.getElementById('users-count');
    if (usersCount) {
        usersCount.innerText = 'No users loaded';
        usersCount.className = 'status-badge status-info';
    }
    
    // Disable buttons
    const scanSharePointBtn = document.getElementById('scan-sharepoint-btn');
    if (scanSharePointBtn) {
        scanSharePointBtn.disabled = true;
    }
    
    const scanOneDriveBtn = document.getElementById('scan-onedrive-btn');
    if (scanOneDriveBtn) {
        scanOneDriveBtn.disabled = true;
    }
    
    const selectAllSitesBtn = document.getElementById('select-all-sites');
    if (selectAllSitesBtn) {
        selectAllSitesBtn.disabled = true;
    }
    
    const deselectAllSitesBtn = document.getElementById('deselect-all-sites');
    if (deselectAllSitesBtn) {
        deselectAllSitesBtn.disabled = true;
    }
    
    const selectAllUsersBtn = document.getElementById('select-all-users');
    if (selectAllUsersBtn) {
        selectAllUsersBtn.disabled = true;
    }
    
    const deselectAllUsersBtn = document.getElementById('deselect-all-users');
    if (deselectAllUsersBtn) {
        deselectAllUsersBtn.disabled = true;
    }
}

function updateScanSettings(newSettings) {
    scanSettings = { ...scanSettings, ...newSettings };
    console.log('Scan settings updated:', scanSettings);
}

function resetScanController() {
    controller = { stop: false };
}

// PATH FORMATTING UTILITIES
function formatItemPath(parentPath, itemName, driveName = 'Documents', scanType = 'sharepoint') {
    let itemPath = '';
    
    if (scanType === 'onedrive') {
        if (parentPath) {
            let cleanPath = parentPath;
            cleanPath = cleanPath.replace('/drive/root:', '');
            cleanPath = cleanPath.replace(/^\/drives\/[^\/]+/, '');
            itemPath = cleanPath ? `${cleanPath}/${itemName}` : `/${itemName}`;
        } else {
            itemPath = `/${itemName}`;
        }
    } else {
        const drivePrefix = driveName || 'Documents';
        if (parentPath) {
            let cleanPath = parentPath;
            cleanPath = cleanPath.replace('/drive/root:', '');
            cleanPath = cleanPath.replace(/^\/drives\/[^\/]+/, '');
            if (cleanPath && cleanPath !== '/') {
                itemPath = `/${drivePrefix}${cleanPath}/${itemName}`;
            } else {
                itemPath = `/${drivePrefix}/${itemName}`;
            }
        } else {
            itemPath = `/${drivePrefix}/${itemName}`;
        }
    }
    
    // Clean up path
    itemPath = itemPath.replace(/\/+/g, '/');
    if (!itemPath.startsWith('/')) itemPath = '/' + itemPath;
    
    return itemPath;
}

// VALIDATION UTILITIES
function validateTenantAndClientIds(tenantId, clientId) {
    if (!tenantId || !clientId) {
        return { isValid: false, message: 'Both Tenant ID and Client ID are required' };
    }
    
    // Basic GUID format validation
    const guidPattern = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
    
    if (!guidPattern.test(tenantId)) {
        return { isValid: false, message: 'Tenant ID must be a valid GUID format' };
    }
    
    if (!guidPattern.test(clientId)) {
        return { isValid: false, message: 'Client ID must be a valid GUID format' };
    }
    
    return { isValid: true, message: 'Valid credentials' };
}

// PROGRESS TRACKING UTILITIES
function updateProgressBar(progressBarId, percentage) {
    const progressBar = document.getElementById(progressBarId);
    if (progressBar) {
        progressBar.style.width = `${Math.min(100, Math.max(0, percentage))}%`;
    }
}

function updateProgressText(progressTextId, text) {
    const progressText = document.getElementById(progressTextId);
    if (progressText) {
        progressText.innerText = text;
    }
}

function showProgressSection(progressSectionId) {
    const progressSection = document.getElementById(progressSectionId);
    if (progressSection) {
        progressSection.style.display = 'block';
    }
}

function hideProgressSection(progressSectionId) {
    const progressSection = document.getElementById(progressSectionId);
    if (progressSection) {
        progressSection.style.display = 'none';
    }
}

// CSV UTILITIES
function validateCSVData(data) {
    if (!Array.isArray(data) || data.length === 0) {
        return { isValid: false, message: 'CSV data is empty or invalid' };
    }
    
    const requiredColumns = ['ItemID', 'Action'];
    const firstRow = data[0];
    
    for (const column of requiredColumns) {
        if (!(column in firstRow)) {
            return { isValid: false, message: `Missing required column: ${column}` };
        }
    }
    
    // Validate action values
    const validActions = ['add', 'remove', 'modify'];
    const invalidActions = data.filter(row => 
        row.Action && !validActions.includes(row.Action.toLowerCase())
    );
    
    if (invalidActions.length > 0) {
        return { 
            isValid: false, 
            message: `Invalid action values found. Valid actions: ${validActions.join(', ')}` 
        };
    }
    
    return { isValid: true, message: 'CSV data is valid' };
}

// SHAREPOINT GROUPS UTILITIES
function isDefaultSharePointGroup(groupName) {
    if (!groupName) return false;
    
    const groupNameLower = groupName.toLowerCase();
    return DEFAULT_SHAREPOINT_GROUPS.some(defaultGroup => 
        groupNameLower.includes(defaultGroup.toLowerCase()) || 
        defaultGroup.toLowerCase().includes(groupNameLower)
    );
}

function shouldShowSharePointGroups() {
    const checkbox = document.getElementById('show-sharepoint-groups');
    return checkbox ? checkbox.checked : true; // Default to true if checkbox not found
}

// PRESERVATION HOLD LIBRARY DETECTION
function isPreservationHoldLibrary(libraryName) {
    if (!libraryName) return false;
    
    const libraryNameLower = libraryName.toLowerCase();
    return PRESERVATION_HOLD_PATTERNS.some(pattern => 
        libraryNameLower.includes(pattern.toLowerCase())
    );
}

function shouldExcludePreservationHolds() {
    const checkbox = document.getElementById('exclude-preservation-holds');
    const isChecked = checkbox ? checkbox.checked : false;
    console.log('🔍 CHECKBOX STATE:', {
        checkboxExists: !!checkbox,
        isChecked: isChecked
    });
    return isChecked; // Default to false (include preservation holds)
}

function shouldSkipPreservationHoldLibrary(libraryName) {
    const checkboxEnabled = shouldExcludePreservationHolds();
    const isHoldLibrary = isPreservationHoldLibrary(libraryName);
    
    // Debug logging for preservation hold detection
    console.log('🔍 PRESERVATION HOLD DEBUG:', {
        libraryName: libraryName,
        checkboxEnabled: checkboxEnabled,
        isHoldLibrary: isHoldLibrary,
        patterns: PRESERVATION_HOLD_PATTERNS
    });
    
    // Only skip if the checkbox is checked (exclusion enabled)
    if (!checkboxEnabled) {
        if (isHoldLibrary) {
            console.log('✅ PRESERVATION HOLD DETECTED but INCLUSION enabled:', libraryName);
        }
        return false; // Don't skip if exclusion is disabled
    }
    
    if (isHoldLibrary) {
        console.log('🚫 SKIPPING PRESERVATION HOLD LIBRARY:', libraryName);
        return true;
    }
    
    return false;
}

// EXPORT FUNCTIONS AND VARIABLES
window.configModule = {
    // Constants
    SKIP_FOLDERS,
    APP_CONFIG,
    
    // State Variables
    get scanSettings() { return scanSettings; },
    get sites() { return sites; },
    get users() { return users; },
    get selectedSiteIds() { return selectedSiteIds; },
    get selectedUserIds() { return selectedUserIds; },
    get results() { return results; },
    get scanning() { return scanning; },
    get tenantDomains() { return tenantDomains; },
    get controller() { return controller; },
    get bulkCsvData() { return bulkCsvData; },
    
    // State Setters
    set sites(value) { sites = value; },
    set users(value) { users = value; },
    set results(value) { results = value; },
    set scanning(value) { scanning = value; },
    set tenantDomains(value) { tenantDomains = value; },
    set bulkCsvData(value) { bulkCsvData = value; },
    
    // Utility Functions
    showToast,
    shouldSkipFolder,
    
    // Permission Classification
    isExternalUser,
    isInternalUser,
    classifyPermission,
    shouldIncludePermission,
    extractUserFromPermission,
    extractExpirationDate,
    
    // State Management
    clearResults,
    clearSitesAndUsers,
    updateScanSettings,
    resetScanController,
    
    // Path Utilities
    formatItemPath,
    
    // Validation
    validateTenantAndClientIds,
    validateCSVData,
    
    // Progress Tracking
    updateProgressBar,
    updateProgressText,
    showProgressSection,
    hideProgressSection,
    
    // SharePoint Groups
    isDefaultSharePointGroup,
    shouldShowSharePointGroups,
    
    // Preservation Hold Libraries
    isPreservationHoldLibrary,
    shouldExcludePreservationHolds,
    shouldSkipPreservationHoldLibrary
};
