// results.js - Results Module for SharePoint & OneDrive Scanner v3.0
// Handles result processing, filtering, real-time display management, and result UI components

// EXTRACT OWNERS FROM RESULT - ENHANCED TO WORK WITH EXISTING DATA AND NEW DUAL PERMISSION STRUCTURE
function extractOwnersFromResult(result) {
    console.log('🔍 OWNERS: Function called for:', result ? result.itemName : 'NO RESULT');
    
    const configModule = window.configModule;
    if (!result) {
        console.log('❌ OWNERS: Missing result object');
        return 'n/a';
    }
    
    if (!configModule) {
        console.log('❌ OWNERS: configModule not available, trying direct access...');
        // Try to access config functions directly from window if module not loaded
        if (window.extractUserFromPermission) {
            console.log('🔧 OWNERS: Using direct window functions as fallback');
            return extractOwnersDirectly(result);
        } else {
            console.log('❌ OWNERS: No config functions available');
            return 'n/a';
        }
    }
    
    // 🔥 ENHANCED FIX: Try allPermissions first, then fall back to more aggressive analysis of filtered permissions
    let permissionsToCheck = result.allPermissions || result.permissions || [];
    let isUsingFilteredData = !result.allPermissions;
    
    console.log(`🔍 OWNERS: Checking ${permissionsToCheck.length} permissions for ${result.itemName} (${isUsingFilteredData ? 'FILTERED' : 'COMPLETE'} data)`);
    
    const owners = [];
    const ownerRoleKeywords = [
        'owner', 'owners', 'full control', 'fullcontrol', 'edit', 'write', 'manage', 'control'
    ];
    
    // If using filtered data, also look for owner-like groups that might have been included
    const ownerGroupKeywords = [
        'owner', 'owners', 'administrator', 'admin'
    ];
    
    permissionsToCheck.forEach((permission, index) => {
        console.log(`🔍 OWNERS: --- Permission ${index + 1} ---`);
        console.log(`🔍 OWNERS: Permission type:`, {
            hasGrantedTo: !!permission.grantedTo,
            hasUser: !!(permission.grantedTo && permission.grantedTo.user),
            hasGroup: !!(permission.grantedToV2 && permission.grantedToV2.group),
            hasSiteGroup: !!(permission.grantedToV2 && permission.grantedToV2.siteGroup),
            hasLink: !!permission.link,
            roles: permission.roles,
            isFiltered: isUsingFilteredData
        });
        
        const roles = permission.roles || [];
        
        // Check if any role indicates owner-level access
        let hasOwnerRole = false;
        for (const role of roles) {
            const roleLower = role.toLowerCase();
            for (const keyword of ownerRoleKeywords) {
                if (roleLower.includes(keyword)) {
                    console.log(`🔍 OWNERS: ✅ OWNER ROLE FOUND: "${role}" contains "${keyword}"`);
                    hasOwnerRole = true;
                    break;
                }
            }
            if (hasOwnerRole) break;
        }
        
        // 🔥 ENHANCED: If using filtered data and no owner role found, check for owner-like groups
        let hasOwnerGroup = false;
        if (isUsingFilteredData && !hasOwnerRole) {
            // Check group names for owner indicators
            if (permission.grantedToV2 && permission.grantedToV2.group) {
                const groupName = (permission.grantedToV2.group.displayName || '').toLowerCase();
                for (const keyword of ownerGroupKeywords) {
                    if (groupName.includes(keyword)) {
                        console.log(`🔍 OWNERS: ✅ OWNER GROUP FOUND: "${permission.grantedToV2.group.displayName}" contains "${keyword}"`);
                        hasOwnerGroup = true;
                        break;
                    }
                }
            }
            
            // Check site group names for owner indicators
            if (!hasOwnerGroup && permission.grantedToV2 && permission.grantedToV2.siteGroup) {
                const siteGroupName = (permission.grantedToV2.siteGroup.displayName || permission.grantedToV2.siteGroup.loginName || '').toLowerCase();
                for (const keyword of ownerGroupKeywords) {
                    if (siteGroupName.includes(keyword)) {
                        console.log(`🔍 OWNERS: ✅ OWNER SITE GROUP FOUND: "${permission.grantedToV2.siteGroup.displayName}" contains "${keyword}"`);
                        hasOwnerGroup = true;
                        break;
                    }
                }
            }
        }
        
        if (hasOwnerRole || hasOwnerGroup) {
            // Extract user/group name for owner role
            let ownerName = null;
            
            // Direct user grant (most common for owners)
            if (permission.grantedTo && permission.grantedTo.user) {
                const user = permission.grantedTo.user;
                ownerName = user.displayName || user.email || user.userPrincipalName;
                console.log(`🔍 OWNERS: Direct user grant owner: "${ownerName}"`);
            }
            // Group permission (like "Site Owners" group)
            else if (permission.grantedToV2 && permission.grantedToV2.group) {
                const group = permission.grantedToV2.group;
                ownerName = group.displayName || group.email;
                console.log(`🔍 OWNERS: Group owner: "${ownerName}"`);
            }
            // Site group permission
            else if (permission.grantedToV2 && permission.grantedToV2.siteGroup) {
                const siteGroup = permission.grantedToV2.siteGroup;
                ownerName = siteGroup.displayName || siteGroup.loginName;
                console.log(`🔍 OWNERS: Site group owner: "${ownerName}"`);
            }
            // Multi-user permissions in grantedToIdentitiesV2
            else if (Array.isArray(permission.grantedToIdentitiesV2)) {
                const userNames = [];
                for (const identity of permission.grantedToIdentitiesV2) {
                    if (identity.user) {
                        userNames.push(identity.user.displayName || identity.user.email);
                    } else if (identity.group) {
                        userNames.push(identity.group.displayName || identity.group.email);
                    }
                }
                ownerName = userNames.join(', ');
                console.log(`🔍 OWNERS: Multi-identity owners: "${ownerName}"`);
            }
            
            // Add to owners list if valid and not duplicate
            if (ownerName && 
                ownerName !== 'Anyone (Anonymous Link)' && 
                !owners.includes(ownerName)) {
                owners.push(ownerName);
                console.log(`🔍 OWNERS: ✅ ADDED OWNER: "${ownerName}" (${hasOwnerGroup ? 'group-based' : 'role-based'})`);
            } else if (ownerName) {
                console.log(`🔍 OWNERS: ⚠️ SKIPPED DUPLICATE: "${ownerName}"`);
            }
        }
    });
    
    const result_text = owners.length > 0 ? owners.join(', ') : 'n/a';
    console.log(`🔍 OWNERS: 🎯 FINAL RESULT: "${result_text}" (found ${owners.length} owners from ${isUsingFilteredData ? 'FILTERED' : 'COMPLETE'} data)`);
    
    return result_text;
}

// FALLBACK FUNCTION: Extract owners without configModule dependency
function extractOwnersDirectly(result) {
    console.log('🔍 OWNERS FALLBACK: Processing result directly:', result.itemName);
    
    // Use allPermissions if available, otherwise filtered permissions
    let permissionsToCheck = result.allPermissions || result.permissions || [];
    let isUsingFilteredData = !result.allPermissions;
    
    console.log(`🔍 OWNERS FALLBACK: Checking ${permissionsToCheck.length} permissions (${isUsingFilteredData ? 'FILTERED' : 'COMPLETE'} data)`);
    
    const owners = [];
    const ownerRoleKeywords = [
        'owner', 'owners', 'full control', 'fullcontrol', 'edit', 'write', 'manage', 'control'
    ];
    
    permissionsToCheck.forEach((permission, index) => {
        console.log(`🔍 OWNERS FALLBACK: --- Permission ${index + 1} ---`);
        console.log(`🔍 OWNERS FALLBACK: Permission structure:`, {
            hasGrantedTo: !!permission.grantedTo,
            hasUser: !!(permission.grantedTo && permission.grantedTo.user),
            hasGroup: !!(permission.grantedToV2 && permission.grantedToV2.group),
            hasSiteGroup: !!(permission.grantedToV2 && permission.grantedToV2.siteGroup),
            roles: permission.roles
        });
        
        const roles = permission.roles || [];
        
        // Check if any role indicates owner-level access
        let hasOwnerRole = false;
        for (const role of roles) {
            const roleLower = role.toLowerCase();
            for (const keyword of ownerRoleKeywords) {
                if (roleLower.includes(keyword)) {
                    console.log(`🔍 OWNERS FALLBACK: ✅ OWNER ROLE FOUND: "${role}" contains "${keyword}"`);
                    hasOwnerRole = true;
                    break;
                }
            }
            if (hasOwnerRole) break;
        }
        
        if (hasOwnerRole) {
            // Extract user/group name for owner role
            let ownerName = null;
            
            // Direct user grant
            if (permission.grantedTo && permission.grantedTo.user) {
                const user = permission.grantedTo.user;
                ownerName = user.displayName || user.email || user.userPrincipalName;
                console.log(`🔍 OWNERS FALLBACK: Direct user grant owner: "${ownerName}"`);
            }
            // Group permission
            else if (permission.grantedToV2 && permission.grantedToV2.group) {
                const group = permission.grantedToV2.group;
                ownerName = group.displayName || group.email;
                console.log(`🔍 OWNERS FALLBACK: Group owner: "${ownerName}"`);
            }
            // Site group permission
            else if (permission.grantedToV2 && permission.grantedToV2.siteGroup) {
                const siteGroup = permission.grantedToV2.siteGroup;
                ownerName = siteGroup.displayName || siteGroup.loginName;
                console.log(`🔍 OWNERS FALLBACK: Site group owner: "${ownerName}"`);
            }
            // Multi-user permissions
            else if (Array.isArray(permission.grantedToIdentitiesV2)) {
                const userNames = [];
                for (const identity of permission.grantedToIdentitiesV2) {
                    if (identity.user) {
                        userNames.push(identity.user.displayName || identity.user.email);
                    } else if (identity.group) {
                        userNames.push(identity.group.displayName || identity.group.email);
                    }
                }
                ownerName = userNames.join(', ');
                console.log(`🔍 OWNERS FALLBACK: Multi-identity owners: "${ownerName}"`);
            }
            
            // Add to owners list if valid and not duplicate
            if (ownerName && 
                ownerName !== 'Anyone (Anonymous Link)' && 
                !owners.includes(ownerName)) {
                owners.push(ownerName);
                console.log(`🔍 OWNERS FALLBACK: ✅ ADDED OWNER: "${ownerName}"`);
            } else if (ownerName) {
                console.log(`🔍 OWNERS FALLBACK: ⚠️ SKIPPED DUPLICATE: "${ownerName}"`);
            }
        }
    });
    
    const result_text = owners.length > 0 ? owners.join(', ') : 'n/a';
    console.log(`🔍 OWNERS FALLBACK: 🎯 FINAL RESULT: "${result_text}" (found ${owners.length} owners)`);
    
    return result_text;
}

// RESULTS FILTERING FUNCTIONALITY
function initializeResultsFiltering() {
    const filterButtons = document.querySelectorAll('#sharing-filters .filter-btn');
    
    filterButtons.forEach(btn => {
        btn.addEventListener('click', () => {
            const filterType = btn.dataset.filter;
            
            // Update active state
            filterButtons.forEach(b => b.classList.remove('active'));
            btn.classList.add('active');
            
            // Apply filter to displayed results
            applyResultsFilter(filterType);
            
            console.log(`Results filter applied: ${filterType}`);
            
            if (window.configModule && window.configModule.showToast) {
                window.configModule.showToast(`Showing ${filterType === 'external' ? 'external sharing' : filterType === 'internal' ? 'internal sharing' : 'all sharing'} results`);
            }
        });
    });
}

function applyResultsFilter(filterType) {
    const configModule = window.configModule;
    if (!configModule) return;
    
    console.log(`🔧 PERMISSION-LEVEL FILTERING: Applying ${filterType} filter to individual permissions within results`);
    
    if (currentView === 'table') {
        // Re-render table view with filtered permissions
        displayResultsAsTable();
    } else if (currentView === 'hierarchy') {
        // Re-render hierarchy view with filtered permissions
        displayResultsAsHierarchy();
    } else {
        // Re-render card view with filtered permissions
        displayResultsAsCards();
    }
    
    // Count total filtered permissions across all results
    let totalFilteredPermissions = 0;
    configModule.results.forEach(result => {
        const filteredPermissions = getFilteredPermissions(result.permissions, filterType);
        totalFilteredPermissions += filteredPermissions.length;
    });
    
    // Update result count display
    const resultCount = document.getElementById('result-count');
    if (resultCount) {
        if (filterType === 'all') {
            resultCount.innerText = `${configModule.results.length} items found`;
        } else {
            resultCount.innerText = `${configModule.results.length} items (${totalFilteredPermissions} ${filterType} permissions shown)`;
        }
    }
    
    console.log(`✅ PERMISSION-LEVEL FILTERING: Showing ${totalFilteredPermissions} ${filterType} permissions across ${configModule.results.length} items`);
}

function shouldShowResult(resultIndex, filterType) {
    const configModule = window.configModule;
    if (!configModule) return false;
    
    if (filterType === 'all') return true;
    if (resultIndex >= configModule.results.length) return false;
    
    const result = configModule.results[resultIndex];
    if (!result || !result.permissions) return false;
    
    // Check if any permission matches the filter
    for (const permission of result.permissions) {
        const classification = configModule.classifyPermission(permission, configModule.tenantDomains);
        
        switch (filterType) {
            case 'external':
                if (classification === 'external' || classification === 'mixed') {
                    return true;
                }
                break;
            case 'internal':
                if (classification === 'internal' || classification === 'mixed') {
                    return true;
                }
                break;
        }
    }
    
    return false;
}

// ENHANCED REAL-TIME RESULTS DISPLAY WITH ITEMID - FIXED VERSION WITH REAL-TIME FILTERING
function addResultToDisplay(result) {
    try {
        const resultsContainer = document.getElementById('results-container');
        if (!resultsContainer) {
            console.error('❌ Results container not found');
            return;
        }
        
        // Ensure results UI components are visible and initialized
        ensureResultsUIInitialized(resultsContainer);
        
        let resultsList = document.getElementById('results-list');
        if (!resultsList) {
            console.error('❌ Could not create results list container after initialization');
            return;
        }
        
        // 🔥 CRITICAL FIX: Check current results filter state before displaying
        const currentFilter = getCurrentResultsFilter();
        const configModule = window.configModule;
        const resultIndex = configModule ? configModule.results.length - 1 : 0; // This result was just added to results array
        const shouldShow = shouldShowResultBasedOnFilter(result, currentFilter);
        
        console.log(`🔍 REAL-TIME: Adding result #${resultIndex + 1}: ${result.itemName} (${result.itemType}) - Filter: ${currentFilter}, Show: ${shouldShow}`);
        
        // Create and insert the element (will be hidden if filter doesn't match)
        const resultDiv = createAndInsertResultElement(result, resultsList, shouldShow);
        
        // Force immediate DOM update and rendering
        resultsList.offsetHeight; // Force reflow
        
        // Update result count display with filter consideration
        updateResultsDisplayWithFilter(currentFilter);
        
        console.log(`✅ Result ${shouldShow ? 'displayed' : 'hidden'} based on filter (${currentFilter}): ${result.itemName} (Total: ${configModule ? configModule.results.length : 0})`);
        
    } catch (error) {
        console.error('❌ Error in addResultToDisplay:', error);
    }
}

// NEW FUNCTION: Get current active results filter
function getCurrentResultsFilter() {
    const activeFilterBtn = document.querySelector('#sharing-filters .filter-btn.active');
    return activeFilterBtn ? activeFilterBtn.dataset.filter : 'all';
}

// NEW FUNCTION: Check if result should show based on current filter
function shouldShowResultBasedOnFilter(result, filterType) {
    const configModule = window.configModule;
    if (!configModule) return false;
    
    if (filterType === 'all') return true;
    if (!result || !result.permissions) return false;
    
    // Check if any permission matches the filter
    for (const permission of result.permissions) {
        const classification = configModule.classifyPermission(permission, configModule.tenantDomains);
        
        switch (filterType) {
            case 'external':
                if (classification === 'external' || classification === 'mixed') {
                    return true;
                }
                break;
            case 'internal':
                if (classification === 'internal' || classification === 'mixed') {
                    return true;
                }
                break;
        }
    }
    
    return false;
}

// NEW FUNCTION: Get filtered permissions based on filter type AND SharePoint groups setting
function getFilteredPermissions(permissions, filterType) {
    const configModule = window.configModule;
    if (!configModule) {
        return permissions;
    }
    
    console.log(`🔍 SHAREPOINT GROUPS DEBUG: Starting filter with ${permissions.length} permissions, filter: ${filterType}`);
    console.log(`🔍 SHAREPOINT GROUPS DEBUG: shouldShowSharePointGroups() = ${configModule.shouldShowSharePointGroups()}`);
    
    let filteredPermissions = permissions;
    
    // First apply sharing filter (external/internal/all)
    if (filterType !== 'all') {
        filteredPermissions = filteredPermissions.filter(permission => {
            const classification = configModule.classifyPermission(permission, configModule.tenantDomains);
            
            switch (filterType) {
                case 'external':
                    return classification === 'external' || classification === 'mixed';
                case 'internal':
                    return classification === 'internal' || classification === 'mixed';
                default:
                    return true;
            }
        });
    }
    
    console.log(`🔍 SHAREPOINT GROUPS DEBUG: After sharing filter: ${filteredPermissions.length} permissions remaining`);
    
    // Then apply SharePoint groups filter if checkbox is unchecked
    if (!configModule.shouldShowSharePointGroups()) {
        console.log(`🚨 SHAREPOINT GROUPS DEBUG: Checkbox is UNCHECKED - should filter out default groups`);
        
        const beforeFilterCount = filteredPermissions.length;
        filteredPermissions = filteredPermissions.filter(permission => {
            
            // Check if this is a regular group permission
            if (permission.grantedToV2 && permission.grantedToV2.group) {
                const groupName = permission.grantedToV2.group.displayName || permission.grantedToV2.group.email || '';
                const isDefaultGroup = configModule.isDefaultSharePointGroup(groupName);
                
                console.log(`🔍 SHAREPOINT GROUP FILTER: ${groupName} -> Default Group: ${isDefaultGroup}, Show Groups: ${configModule.shouldShowSharePointGroups()}`);
                
                if (isDefaultGroup) {
                    console.log(`🚫 SHAREPOINT GROUP FILTER: EXCLUDING regular group: "${groupName}"`);
                    return false; // Exclude default SharePoint groups
                }
            }
            
            // 🔥 CRITICAL FIX: Also check for site group permissions (Team Site Owners, etc.)
            if (permission.grantedToV2 && permission.grantedToV2.siteGroup) {
                const siteGroupName = permission.grantedToV2.siteGroup.displayName || permission.grantedToV2.siteGroup.loginName || '';
                const isDefaultSiteGroup = configModule.isDefaultSharePointGroup(siteGroupName);
                
                console.log(`🔍 SHAREPOINT SITE GROUP FILTER: "${siteGroupName}" -> Default Site Group: ${isDefaultSiteGroup}, Show Groups: ${configModule.shouldShowSharePointGroups()}`);
                
                if (isDefaultSiteGroup) {
                    console.log(`🚫 SHAREPOINT SITE GROUP FILTER: EXCLUDING site group: "${siteGroupName}"`);
                    return false; // Exclude default SharePoint site groups
                }
            }
            
            // Also check in grantedToIdentitiesV2 for groups
            if (Array.isArray(permission.grantedToIdentitiesV2)) {
                const hasDefaultGroup = permission.grantedToIdentitiesV2.some(g => {
                    if (g.group) {
                        const groupName = g.group.displayName || g.group.email || '';
                        const isDefault = configModule.isDefaultSharePointGroup(groupName);
                        console.log(`🔍 SHAREPOINT GROUP FILTER (grantedToIdentitiesV2): ${groupName} -> Default: ${isDefault}`);
                        return isDefault;
                    }
                    return false;
                });
                
                if (hasDefaultGroup) {
                    console.log(`🚫 SHAREPOINT GROUP FILTER (grantedToIdentitiesV2): EXCLUDING permission with default group`);
                    return false; // Exclude permissions with default groups
                }
            }
            
            console.log(`✅ SHAREPOINT GROUP FILTER: KEEPING permission (not a default group)`);
            return true; // Keep non-group permissions and custom groups
        });
        
        console.log(`🔍 SHAREPOINT GROUPS DEBUG: After groups filter: ${filteredPermissions.length} permissions remaining (filtered out ${beforeFilterCount - filteredPermissions.length})`);
    } else {
        console.log(`✅ SHAREPOINT GROUPS DEBUG: Checkbox is CHECKED - showing all groups`);
    }
    
    console.log(`🔍 SHAREPOINT GROUPS DEBUG: Final result: ${filteredPermissions.length} permissions to display`);
    return filteredPermissions;
}

// NEW FUNCTION: Update results display count with filter consideration
function updateResultsDisplayWithFilter(filterType) {
    const resultCount = document.getElementById('result-count');
    const configModule = window.configModule;
    
    if (!configModule || !resultCount) return;
    
    if (filterType === 'all') {
        resultCount.innerText = `${configModule.results.length} found`;
    } else {
        // Count visible results
        const resultsList = document.getElementById('results-list');
        if (resultsList) {
            const visibleResults = Array.from(resultsList.children).filter(div => 
                div.style.display !== 'none'
            ).length;
            resultCount.innerText = `${visibleResults} of ${configModule.results.length} shown (${filterType})`;
        } else {
            resultCount.innerText = `${configModule.results.length} found`;
        }
    }
    
    if (configModule.results.length > 0) {
        resultCount.className = 'status-badge status-approved';
        const exportBtn = document.getElementById('export-btn');
        if (exportBtn) {
            exportBtn.disabled = false;
        }
    }
}

// NEW FUNCTION: Ensure results UI is properly initialized
function ensureResultsUIInitialized(resultsContainer) {
    let resultsList = document.getElementById('results-list');
    
    if (!resultsList) {
        console.log('🔧 Initializing results list container for real-time display');
        resultsContainer.innerHTML = '<div style="max-height: 400px; overflow-y: auto; border: 1px solid var(--border); border-radius: 8px; background: white;" id="results-list"></div>';
        
        // Show UI controls
        const actionsEl = document.getElementById('results-actions');
        const filtersEl = document.getElementById('sharing-filters');
        const bulkEl = document.getElementById('bulk-controls');
        
        if (actionsEl) actionsEl.style.display = 'block';
        if (filtersEl) filtersEl.style.display = 'flex';
        if (bulkEl) bulkEl.style.display = 'flex';
        
        console.log('✅ Results UI initialized and controls made visible');
    }
}

function createAndInsertResultElement(result, resultsList, shouldShow = true) {
    const configModule = window.configModule;
    if (!configModule) return null;
    
    try {
        const resultDiv = document.createElement('div');
        resultDiv.style.cssText = 'border: 1px solid var(--border); border-radius: 8px; padding: 16px; margin-bottom: 12px; background: white; animation: fadeIn 0.3s ease-in;';
        
        // 🔥 CRITICAL FIX: Apply filter visibility immediately when creating the element
        if (!shouldShow) {
            resultDiv.style.display = 'none';
            console.log(`🔍 REAL-TIME FILTER: Result hidden immediately: ${result.itemName}`);
        }
        
        const headerDiv = document.createElement('div');
        headerDiv.style.cssText = 'display: grid; grid-template-columns: 1fr auto auto; gap: 16px; align-items: flex-start; margin-bottom: 16px;';
        
        const infoDiv = document.createElement('div');
        infoDiv.style.cssText = 'flex: 1;';
        
        const title = document.createElement('h3');
        title.style.cssText = 'margin: 0 0 4px 0; font-size: 16px; color: var(--text); font-weight: 600;';
        
        let displayPath = result.itemPath;
        if (!displayPath || displayPath === 'undefined') {
            displayPath = `/${result.itemName}`;
        }
        displayPath = displayPath.replace(/\/+/g, '/');
        if (!displayPath.startsWith('/')) displayPath = '/' + displayPath;
        
        // Fix for file icon display - files should show appropriate icons based on source and type
        let displayIcon;
        if (result.itemType === 'file') {
            // For files: SharePoint files show just file icon, OneDrive files show source + file icon
            displayIcon = result.scanType === 'onedrive' ? '☁️📄' : '📄';
        } else {
            // For folders: SharePoint folders show just folder icon, OneDrive folders show source + folder icon
            displayIcon = result.scanType === 'onedrive' ? '☁️📁' : '📁';
        }
        title.innerText = `${displayIcon} ${result.siteName}${displayPath}`;
        
        const sourceUrl = document.createElement('p');
        sourceUrl.style.cssText = 'margin: 0; font-size: 12px; color: var(--text-muted);';
        sourceUrl.innerText = result.siteUrl || (result.scanType === 'onedrive' ? 'Personal OneDrive' : 'SharePoint Site');
        
        infoDiv.appendChild(title);
        infoDiv.appendChild(sourceUrl);
        
        // ItemID Column - Enhanced positioning and visibility (moved even further left)
        const itemIdDiv = document.createElement('div');
        itemIdDiv.style.cssText = 'display: flex; flex-direction: column; align-items: flex-start; padding-right: 32px; margin-right: 16px;';
        
        const itemIdLabel = document.createElement('span');
        itemIdLabel.style.cssText = 'font-size: 10px; color: var(--text-muted); text-transform: uppercase; font-weight: 600; margin-bottom: 2px;';
        itemIdLabel.innerText = 'ItemID';
        
        const itemIdValue = document.createElement('span');
        itemIdValue.style.cssText = 'font-size: 11px; color: var(--text-muted); font-family: monospace; max-width: 200px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;';
        itemIdValue.innerText = result.itemId;
        itemIdValue.title = result.itemId; // Full ID on hover
        
        itemIdDiv.appendChild(itemIdLabel);
        itemIdDiv.appendChild(itemIdValue);
        
        // Owners Column - New column showing item owners
        const ownersDiv = document.createElement('div');
        ownersDiv.style.cssText = 'display: flex; flex-direction: column; align-items: flex-start; padding-right: 16px; margin-right: 8px;';
        
        const ownersLabel = document.createElement('span');
        ownersLabel.style.cssText = 'font-size: 10px; color: var(--text-muted); text-transform: uppercase; font-weight: 600; margin-bottom: 2px;';
        ownersLabel.innerText = 'Owners';
        
        const ownersValue = document.createElement('span');
        ownersValue.style.cssText = 'font-size: 11px; color: var(--text); max-width: 150px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;';
        const owners = extractOwnersFromResult(result);
        ownersValue.innerText = owners;
        ownersValue.title = owners; // Full owners list on hover
        
        ownersDiv.appendChild(ownersLabel);
        ownersDiv.appendChild(ownersValue);
        
        headerDiv.appendChild(infoDiv);
        headerDiv.appendChild(itemIdDiv);
        headerDiv.appendChild(ownersDiv);
        
        // NOTE: Folder-level action buttons removed - now using user-level buttons in permission table
        
        // Enhanced results table with ItemID column
        const table = document.createElement('table');
        table.className = 'results-table';
        
        const thead = document.createElement('thead');
        thead.innerHTML = `
            <tr>
                <th>Who Has Access</th>
                <th>Permission Level</th>
                <th>Sharing Type</th>
                <th>Link Expiration</th>
                <th>Actions</th>
            </tr>
        `;
        table.appendChild(thead);
        
        const tbody = document.createElement('tbody');

        // 🔥 CRITICAL FIX: Filter permissions based on current active filter before displaying
        const currentFilter = getCurrentResultsFilter();
        const filteredPermissions = getFilteredPermissions(result.permissions, currentFilter);
        
        console.log(`🔍 PERMISSION FILTERING: ${result.itemName} - Showing ${filteredPermissions.length} of ${result.permissions.length} permissions (filter: ${currentFilter})`);

        filteredPermissions.forEach((p, index) => {
            const tr = document.createElement('tr');
            
            // Check if this is a direct grant for enhanced display
            const isDirectGrantPermission = configModule.isDirectGrant(p);
            const directGrantDetails = isDirectGrantPermission ? configModule.extractDirectGrantDetails(p, configModule.tenantDomains) : null;
            const directGrantDisplay = directGrantDetails ? configModule.formatDirectGrantDisplay(directGrantDetails) : null;
            
            const who = configModule.extractUserFromPermission(p, configModule.tenantDomains);
            const roles = (p.roles || []).join(', ') || 'Not specified';
            const classification = configModule.classifyPermission(p, configModule.tenantDomains);
            const exp = configModule.extractExpirationDate(p);

            const tdWho = document.createElement('td');
            if (directGrantDisplay) {
                // Enhanced display for direct grants
                const whoContainer = document.createElement('div');
                whoContainer.innerHTML = `
                    <div style="font-weight: 600; color: var(--text);">${directGrantDisplay.primaryText}</div>
                    <div style="font-size: 11px; color: var(--text-muted); margin-top: 2px;">${directGrantDisplay.secondaryText}</div>
                    ${directGrantDisplay.riskFactors.length > 0 ? 
                        `<div style="font-size: 10px; color: var(--danger); margin-top: 2px;">⚠️ ${directGrantDisplay.riskFactors.join(', ')}</div>` : 
                        ''}
                `;
                tdWho.appendChild(whoContainer);
            } else {
                tdWho.innerText = who;
            }
            
            const tdRoles = document.createElement('td');
            if (directGrantDisplay && directGrantDisplay.inheritedFrom) {
                const rolesContainer = document.createElement('div');
                rolesContainer.innerHTML = `
                    <div>${roles}</div>
                    <div style="font-size: 10px; color: var(--text-muted); margin-top: 2px;">
                        📋 Inherited from: ${directGrantDisplay.inheritedFrom.name}
                    </div>
                `;
                tdRoles.appendChild(rolesContainer);
            } else {
                tdRoles.innerText = roles;
            }
            
            const tdType = document.createElement('td');
            const typeContainer = document.createElement('div');
            
            const typeBadge = document.createElement('span');
            typeBadge.style.cssText = 'padding: 2px 6px; border-radius: 3px; font-size: 10px; font-weight: 600; margin-right: 4px;';
            
            if (classification === 'external') {
                typeBadge.className = 'external-badge';
                typeBadge.innerText = 'EXTERNAL';
            } else if (classification === 'internal') {
                typeBadge.className = 'internal-badge';
                typeBadge.innerText = 'INTERNAL';
            } else {
                typeBadge.style.background = 'var(--warning)';
                typeBadge.style.color = 'white';
                typeBadge.innerText = classification.toUpperCase();
            }
            
            typeContainer.appendChild(typeBadge);
            
            // Add risk level badge for direct grants
            if (directGrantDisplay && directGrantDisplay.riskLevel) {
                const riskBadge = document.createElement('span');
                riskBadge.style.cssText = 'padding: 2px 6px; border-radius: 3px; font-size: 9px; font-weight: 600; margin-left: 2px;';
                
                switch (directGrantDisplay.riskLevel) {
                    case 'CRITICAL':
                        riskBadge.style.background = '#dc2626';
                        riskBadge.style.color = 'white';
                        riskBadge.innerText = '🔥 CRITICAL';
                        break;
                    case 'HIGH':
                        riskBadge.style.background = 'var(--danger)';
                        riskBadge.style.color = 'white';
                        riskBadge.innerText = '⚠️ HIGH';
                        break;
                    case 'MEDIUM':
                        riskBadge.style.background = 'var(--warning)';
                        riskBadge.style.color = 'white';
                        riskBadge.innerText = '⚡ MEDIUM';
                        break;
                }
                typeContainer.appendChild(riskBadge);
            }
            
            tdType.appendChild(typeContainer);
            
            const tdExp = document.createElement('td');
            if (directGrantDisplay && directGrantDisplay.hasApplication) {
                const expContainer = document.createElement('div');
                expContainer.innerHTML = `
                    <div>${exp}</div>
                    <div style="font-size: 10px; color: var(--primary); margin-top: 2px;">
                        🔗 Via: ${directGrantDisplay.applicationName || 'Application'}
                    </div>
                `;
                tdExp.appendChild(expContainer);
            } else {
                tdExp.innerText = exp;
            }
            
            // ENHANCED: User-level action buttons with improved design and functionality
            const tdActions = document.createElement('td');
            tdActions.style.cssText = 'text-align: center; padding: 8px; vertical-align: middle;';
            
            const actionsContainer = document.createElement('div');
            actionsContainer.style.cssText = 'display: flex; gap: 6px; justify-content: center; align-items: center; flex-wrap: wrap;';
            
            const hasLink = p.link;
            
            if (hasLink) {
                // Enhanced Set Expiration button for individual permission
                const expBtn = document.createElement('button');
                expBtn.className = 'user-action-btn user-action-exp';
                expBtn.style.cssText = `
                    background: var(--purple); 
                    color: white; 
                    border: none; 
                    padding: 6px 10px; 
                    border-radius: 4px; 
                    font-size: 11px; 
                    font-weight: 600;
                    cursor: pointer;
                    transition: all 0.2s ease;
                    box-shadow: 0 2px 4px rgba(0,0,0,0.1);
                    min-width: 45px;
                `;
                expBtn.textContent = '📅 EXP';
                expBtn.title = `Set expiration date for ${who}`;
                expBtn.onclick = (e) => {
                    e.stopPropagation();
                    if (window.permissionsModule?.showExpirationDialog) {
                        // Pass individual permission for targeted action
                        window.permissionsModule.showExpirationDialog(result, resultDiv, p);
                    }
                };
                
                // Add hover effect
                expBtn.onmouseenter = () => {
                    expBtn.style.transform = 'translateY(-1px)';
                    expBtn.style.boxShadow = '0 4px 8px rgba(0,0,0,0.15)';
                };
                expBtn.onmouseleave = () => {
                    expBtn.style.transform = 'translateY(0)';
                    expBtn.style.boxShadow = '0 2px 4px rgba(0,0,0,0.1)';
                };
                
                actionsContainer.appendChild(expBtn);
                
                // Enhanced Disable Link button for individual permission
                const linkBtn = document.createElement('button');
                linkBtn.className = 'user-action-btn user-action-link';
                linkBtn.style.cssText = `
                    background: var(--orange); 
                    color: white; 
                    border: none; 
                    padding: 6px 10px; 
                    border-radius: 4px; 
                    font-size: 11px; 
                    font-weight: 600;
                    cursor: pointer;
                    transition: all 0.2s ease;
                    box-shadow: 0 2px 4px rgba(0,0,0,0.1);
                    min-width: 45px;
                `;
                linkBtn.textContent = '🔗 LINK';
                linkBtn.title = `Remove sharing link for ${who}`;
                linkBtn.onclick = (e) => {
                    e.stopPropagation();
                    if (confirm(`❓ Remove sharing link for ${who}?\n\nThis will disable the sharing link but keep other permissions.`)) {
                        if (window.permissionsModule?.disableLinks) {
                            // Pass individual permission for targeted action
                            window.permissionsModule.disableLinks(result, resultDiv, p);
                        }
                    }
                };
                
                // Add hover effect
                linkBtn.onmouseenter = () => {
                    linkBtn.style.transform = 'translateY(-1px)';
                    linkBtn.style.boxShadow = '0 4px 8px rgba(0,0,0,0.15)';
                };
                linkBtn.onmouseleave = () => {
                    linkBtn.style.transform = 'translateY(0)';
                    linkBtn.style.boxShadow = '0 2px 4px rgba(0,0,0,0.1)';
                };
                
                actionsContainer.appendChild(linkBtn);
            }
            
            // Enhanced Remove User button for individual permission
            const userBtn = document.createElement('button');
            userBtn.className = 'user-action-btn user-action-user';
            userBtn.style.cssText = `
                background: var(--danger); 
                color: white; 
                border: none; 
                padding: 6px 10px; 
                border-radius: 4px; 
                font-size: 11px; 
                font-weight: 600;
                cursor: pointer;
                transition: all 0.2s ease;
                box-shadow: 0 2px 4px rgba(0,0,0,0.1);
                min-width: 45px;
            `;
            userBtn.textContent = '🚫 USER';
            userBtn.title = `Remove all permissions for ${who}`;
            userBtn.onclick = (e) => {
                e.stopPropagation();
                if (confirm(`⚠️ Remove ALL sharing permissions for ${who}?\n\nThis will completely remove this user's access to the item.`)) {
                    if (window.permissionsModule?.disableAllSharing) {
                        // Pass individual permission for targeted action
                        window.permissionsModule.disableAllSharing(result, resultDiv, p);
                    }
                }
            };
            
            // Add hover effect
            userBtn.onmouseenter = () => {
                userBtn.style.transform = 'translateY(-1px)';
                userBtn.style.boxShadow = '0 4px 8px rgba(0,0,0,0.15)';
            };
            userBtn.onmouseleave = () => {
                userBtn.style.transform = 'translateY(0)';
                userBtn.style.boxShadow = '0 2px 4px rgba(0,0,0,0.1)';
            };
            
            actionsContainer.appendChild(userBtn);
            
            tdActions.appendChild(actionsContainer);
            
            tr.appendChild(tdWho);
            tr.appendChild(tdRoles);
            tr.appendChild(tdType);
            tr.appendChild(tdExp);
            tr.appendChild(tdActions);
            tbody.appendChild(tr);
        });

        table.appendChild(tbody);
        resultDiv.appendChild(headerDiv);
        resultDiv.appendChild(table);
        resultsList.appendChild(resultDiv);
        
        // Scroll to show new result
        resultsList.scrollTop = resultsList.scrollHeight;
        
        console.log(`✅ Result element created: ${result.itemName}`);
        return resultDiv;
        
    } catch (error) {
        console.error('Error creating result element:', error);
        return null;
    }
}

function updateResultsDisplay() {
    const resultCount = document.getElementById('result-count');
    const configModule = window.configModule;
    
    if (!configModule || !resultCount) return;
    
    resultCount.innerText = `${configModule.results.length} found`;
    if (configModule.results.length > 0) {
        resultCount.className = 'status-badge status-approved';
        const exportBtn = document.getElementById('export-btn');
        if (exportBtn) {
            exportBtn.disabled = false;
        }
    }
}


// RESULTS REFRESH FUNCTIONALITY FOR PERMISSION OPERATIONS
async function refreshItemPermissions(result, resultIndex) {
    const configModule = window.configModule;
    const apiModule = window.apiModule;
    
    if (!configModule || !apiModule) {
        console.error('Required modules not available for refreshing permissions');
        return;
    }
    
    try {
        const permissionsUrl = `https://graph.microsoft.com/v1.0/drives/${result.driveId}/items/${result.itemId}/permissions`;
        
        const updatedPermissions = await apiModule.requestQueue.add(async () => {
            return await apiModule.graphGetAll(permissionsUrl);
        });
        
        // Filter based on current scan settings
        const filteredPermissions = updatedPermissions.filter(p => 
            configModule.shouldIncludePermission(p, configModule.tenantDomains, configModule.scanSettings.sharingFilter)
        );
        
        // Update the result in memory
        if (resultIndex >= 0 && resultIndex < configModule.results.length) {
            configModule.results[resultIndex].permissions = filteredPermissions;
        }
        
        console.log(`Refreshed permissions for ${result.itemName}: ${filteredPermissions.length} relevant permissions`);
    } catch (error) {
        console.warn(`Failed to refresh permissions for ${result.itemName}:`, error);
    }
}

// TABLE VIEW FUNCTIONALITY
let currentView = 'card'; // Track current view mode

function initializeViewToggle() {
    const viewToggleButtons = document.querySelectorAll('#view-toggle .filter-btn');
    
    viewToggleButtons.forEach(btn => {
        btn.addEventListener('click', () => {
            const viewType = btn.dataset.view;
            
            // Update active state
            viewToggleButtons.forEach(b => b.classList.remove('active'));
            btn.classList.add('active');
            
            // Switch view
            switchView(viewType);
            
            console.log(`View switched to: ${viewType}`);
            
            if (window.configModule && window.configModule.showToast) {
                const viewName = viewType === 'card' ? 'Card' : 
                                viewType === 'table' ? 'Table' : 
                                viewType === 'hierarchy' ? 'Tree' : viewType;
                window.configModule.showToast(`Switched to ${viewName} View`);
            }
        });
    });
}

function switchView(viewType) {
    currentView = viewType;
    const configModule = window.configModule;
    
    if (!configModule || !configModule.results.length) {
        console.log('No results to switch view for');
        return;
    }
    
    console.log(`🔄 Switching to ${viewType} view with ${configModule.results.length} results`);
    
    if (viewType === 'table') {
        displayResultsAsTable();
    } else if (viewType === 'hierarchy') {
        displayResultsAsHierarchy();
    } else {
        displayResultsAsCards();
    }
}

function displayResultsAsTable() {
    const configModule = window.configModule;
    const resultsContainer = document.getElementById('results-container');
    
    if (!configModule || !resultsContainer) return;
    
    console.log('🔧 Converting to table view...');
    
    // Create table container
    const tableContainer = document.createElement('div');
    tableContainer.className = 'table-container';
    tableContainer.id = 'results-table-container';
    
    const table = document.createElement('table');
    table.className = 'results-table-view';
    table.id = 'results-table';
    
    // Create table header
    const thead = document.createElement('thead');
    thead.innerHTML = `
        <tr>
            <th style="width: 20%;">Item</th>
            <th style="width: 12%;">ItemID</th>
            <th style="width: 12%;">Owners</th>
            <th style="width: 26%;">Permissions</th>
            <th style="width: 12%;">Classifications</th>
            <th style="width: 18%;">Actions</th>
        </tr>
    `;
    table.appendChild(thead);
    
    // Create table body
    const tbody = document.createElement('tbody');
    table.appendChild(tbody);
    
    // Populate table with existing results
    const currentFilter = getCurrentResultsFilter();
    configModule.results.forEach((result, index) => {
        const shouldShow = shouldShowResult(index, currentFilter);
        if (shouldShow) {
            const row = createTableRow(result, index);
            tbody.appendChild(row);
        }
    });
    
    tableContainer.appendChild(table);
    resultsContainer.innerHTML = '';
    resultsContainer.appendChild(tableContainer);
    
    console.log(`✅ Table view created with ${tbody.children.length} visible rows`);
}

function displayResultsAsCards() {
    const configModule = window.configModule;
    const resultsContainer = document.getElementById('results-container');
    
    if (!configModule || !resultsContainer) return;
    
    console.log('🔧 Converting to card view...');
    
    // Create card container
    resultsContainer.innerHTML = '<div style="max-height: 400px; overflow-y: auto; border: 1px solid var(--border); border-radius: 8px; background: white;" id="results-list"></div>';
    
    const resultsList = document.getElementById('results-list');
    
    // Populate with existing results
    const currentFilter = getCurrentResultsFilter();
    configModule.results.forEach((result, index) => {
        const shouldShow = shouldShowResult(index, currentFilter);
        const resultDiv = createAndInsertResultElement(result, resultsList, shouldShow);
    });
    
    console.log(`✅ Card view restored with ${configModule.results.length} results`);
}

function createTableRow(result, resultIndex) {
    const configModule = window.configModule;
    if (!configModule) return null;
    
    const tr = document.createElement('tr');
    tr.dataset.resultIndex = resultIndex;
    
    // Item column
    const tdItem = document.createElement('td');
    
    let displayPath = result.itemPath;
    if (!displayPath || displayPath === 'undefined') {
        displayPath = `/${result.itemName}`;
    }
    displayPath = displayPath.replace(/\/+/g, '/');
    if (!displayPath.startsWith('/')) displayPath = '/' + displayPath;
    
    let displayIcon;
    if (result.itemType === 'file') {
        displayIcon = result.scanType === 'onedrive' ? '☁️📄' : '📄';
    } else {
        displayIcon = result.scanType === 'onedrive' ? '☁️📁' : '📁';
    }
    
    tdItem.innerHTML = `
        <div class="table-item-name">${displayIcon} ${result.siteName}${displayPath}</div>
        <div class="table-item-path">${result.siteUrl || (result.scanType === 'onedrive' ? 'Personal OneDrive' : 'SharePoint Site')}</div>
    `;
    
    // ItemID column
    const tdItemId = document.createElement('td');
    tdItemId.innerHTML = `<div class="table-item-id">${result.itemId}</div>`;
    
    // Owners column - Enhanced debugging for table view
    const tdOwners = document.createElement('td');
    console.log('🔍 TABLE OWNERS DEBUG: Processing result for table:', {
        itemName: result.itemName,
        hasPermissions: !!result.permissions,
        permissionsCount: result.permissions ? result.permissions.length : 0,
        hasAllPermissions: !!result.allPermissions,
        allPermissionsCount: result.allPermissions ? result.allPermissions.length : 0,
        scanType: result.scanType
    });
    
    // Log some sample permissions for debugging
    if (result.allPermissions && result.allPermissions.length > 0) {
        console.log('🔍 TABLE OWNERS DEBUG: Sample allPermissions:', result.allPermissions.slice(0, 2).map(p => ({
            roles: p.roles,
            hasGrantedTo: !!p.grantedTo,
            hasUser: !!(p.grantedTo && p.grantedTo.user),
            hasGroup: !!(p.grantedToV2 && p.grantedToV2.group),
            hasSiteGroup: !!(p.grantedToV2 && p.grantedToV2.siteGroup)
        })));
    }
    
    const owners = extractOwnersFromResult(result);
    console.log('🔍 TABLE OWNERS DEBUG: extractOwnersFromResult returned:', owners);
    tdOwners.innerHTML = `<div class="table-item-owners">${owners}</div>`;
    
    // Permissions column
    const tdPermissions = document.createElement('td');
    tdPermissions.className = 'table-permissions-cell';
    
    // 🔥 CRITICAL FIX: Filter permissions based on current active filter before displaying in table
    const currentFilter = getCurrentResultsFilter();
    const filteredPermissions = getFilteredPermissions(result.permissions, currentFilter);
    
    console.log(`🔍 TABLE PERMISSION FILTERING: ${result.itemName} - Showing ${filteredPermissions.length} of ${result.permissions.length} permissions (filter: ${currentFilter})`);
    
    let permissionsHtml = '';
    filteredPermissions.forEach((p, index) => {
        const who = configModule.extractUserFromPermission(p, configModule.tenantDomains);
        const roles = (p.roles || []).join(', ') || 'Not specified';
        const exp = configModule.extractExpirationDate(p);
        
        permissionsHtml += `
            <div class="table-permission-item">
                <div class="table-permission-who">${who}</div>
                <div class="table-permission-role">${roles}${exp !== 'No expiration' ? ` • Expires: ${exp}` : ''}</div>
            </div>
        `;
    });
    tdPermissions.innerHTML = permissionsHtml;
    
    // Classifications column - also filter based on active filter
    const tdClassifications = document.createElement('td');
    let classificationsHtml = '';
    filteredPermissions.forEach((p) => {
        const classification = configModule.classifyPermission(p, configModule.tenantDomains);
        const badgeClass = classification === 'external' ? 'external-badge' : 'internal-badge';
        classificationsHtml += `<span class="${badgeClass}">${classification.toUpperCase()}</span><br>`;
    });
    tdClassifications.innerHTML = classificationsHtml;
    
    // Actions column
    const tdActions = document.createElement('td');
    tdActions.className = 'table-actions-cell';
    
    const hasLinks = result.permissions.some(p => p.link);
    
    let actionsHtml = '';
    if (hasLinks) {
        actionsHtml += `
            <button class="table-action-btn purple" onclick="handleTableAction('expiration', ${resultIndex})">
                SET EXP
            </button>
            <button class="table-action-btn orange" onclick="handleTableAction('disableLinks', ${resultIndex})">
                DEL LINKS
            </button>
        `;
    }
    actionsHtml += `
        <button class="table-action-btn red" onclick="handleTableAction('disableAll', ${resultIndex})">
            DEL ALL
        </button>
    `;
    tdActions.innerHTML = actionsHtml;
    
    tr.appendChild(tdItem);
    tr.appendChild(tdItemId);
    tr.appendChild(tdOwners);
    tr.appendChild(tdPermissions);
    tr.appendChild(tdClassifications);
    tr.appendChild(tdActions);
    
    return tr;
}

// Table action handler (global function for onclick handlers)
window.handleTableAction = function(actionType, resultIndex) {
    const configModule = window.configModule;
    const permissionsModule = window.permissionsModule;
    
    if (!configModule || !permissionsModule || resultIndex >= configModule.results.length) {
        console.error('Invalid action request');
        return;
    }
    
    const result = configModule.results[resultIndex];
    const tableRow = document.querySelector(`tr[data-result-index="${resultIndex}"]`);
    
    switch (actionType) {
        case 'expiration':
            if (permissionsModule.showExpirationDialog) {
                permissionsModule.showExpirationDialog(result, tableRow);
            }
            break;
        case 'disableLinks':
            if (confirm('Remove all sharing links for this item?')) {
                if (permissionsModule.disableLinks) {
                    permissionsModule.disableLinks(result, tableRow);
                }
            }
            break;
        case 'disableAll':
            if (confirm('Remove all sharing permissions for this item?')) {
                if (permissionsModule.disableAllSharing) {
                    permissionsModule.disableAllSharing(result, tableRow);
                }
            }
            break;
    }
};

function addResultToTableView(result) {
    const tableBody = document.querySelector('#results-table tbody');
    const configModule = window.configModule;
    
    if (!tableBody || !configModule) return;
    
    const currentFilter = getCurrentResultsFilter();
    const resultIndex = configModule.results.length - 1;
    const shouldShow = shouldShowResultBasedOnFilter(result, currentFilter);
    
    if (shouldShow) {
        const row = createTableRow(result, resultIndex);
        tableBody.appendChild(row);
        
        // Scroll to show new row
        const tableContainer = document.getElementById('results-table-container');
        if (tableContainer) {
            tableContainer.scrollTop = tableContainer.scrollHeight;
        }
        
        console.log(`✅ Added result to table: ${result.itemName}`);
    } else {
        console.log(`🔍 Result hidden from table due to filter: ${result.itemName}`);
    }
}

// Enhanced ensureResultsUIInitialized to support both views
function ensureResultsUIInitializedEnhanced(resultsContainer) {
    // Show view controls when results start appearing
    const viewControlsContainer = document.getElementById('view-controls-container');
    if (viewControlsContainer) {
        viewControlsContainer.style.display = 'flex';
    }
    
    // Initialize based on current view mode
    if (currentView === 'table') {
        let tableContainer = document.getElementById('results-table-container');
        if (!tableContainer) {
            console.log('🔧 Initializing table container for real-time display');
            
            const tableContainerDiv = document.createElement('div');
            tableContainerDiv.className = 'table-container';
            tableContainerDiv.id = 'results-table-container';
            
            const table = document.createElement('table');
            table.className = 'results-table-view';
            table.id = 'results-table';
            
            const thead = document.createElement('thead');
            thead.innerHTML = `
                <tr>
                    <th style="width: 20%;">Item</th>
                    <th style="width: 12%;">ItemID</th>
                    <th style="width: 12%;">Owners</th>
                    <th style="width: 26%;">Permissions</th>
                    <th style="width: 12%;">Classifications</th>
                    <th style="width: 18%;">Actions</th>
                </tr>
            `;
            table.appendChild(thead);
            
            const tbody = document.createElement('tbody');
            table.appendChild(tbody);
            
            tableContainerDiv.appendChild(table);
            resultsContainer.innerHTML = '';
            resultsContainer.appendChild(tableContainerDiv);
        }
    } else {
        // Use existing card view initialization
        ensureResultsUIInitialized(resultsContainer);
    }
    
    // Show UI controls
    const actionsEl = document.getElementById('results-actions');
    const filtersEl = document.getElementById('sharing-filters');
    const bulkEl = document.getElementById('bulk-controls');
    
    if (actionsEl) actionsEl.style.display = 'block';
    if (filtersEl && filtersEl.parentElement) filtersEl.parentElement.style.display = 'flex';
    if (bulkEl) bulkEl.style.display = 'flex';
}

// Enhanced addResultToDisplay to support both views
function addResultToDisplayEnhanced(result) {
    try {
        const resultsContainer = document.getElementById('results-container');
        if (!resultsContainer) {
            console.error('❌ Results container not found');
            return;
        }
        
        // Ensure results UI components are visible and initialized
        ensureResultsUIInitializedEnhanced(resultsContainer);
        
        const currentFilter = getCurrentResultsFilter();
        const configModule = window.configModule;
        const shouldShow = shouldShowResultBasedOnFilter(result, currentFilter);
        
        console.log(`🔍 REAL-TIME: Adding result to ${currentView} view: ${result.itemName} (${result.itemType}) - Filter: ${currentFilter}, Show: ${shouldShow}`);
        
        if (currentView === 'table') {
            addResultToTableView(result);
        } else {
            // Use existing card view logic
            let resultsList = document.getElementById('results-list');
            if (resultsList) {
                const resultDiv = createAndInsertResultElement(result, resultsList, shouldShow);
                resultsList.offsetHeight; // Force reflow
            }
        }
        
        // Update result count display with filter consideration
        updateResultsDisplayWithFilter(currentFilter);
        
        console.log(`✅ Result ${shouldShow ? 'displayed' : 'hidden'} in ${currentView} view: ${result.itemName}`);
        
    } catch (error) {
        console.error('❌ Error in addResultToDisplayEnhanced:', error);
    }
}

// SHAREPOINT GROUPS TOGGLE FUNCTIONALITY
function initializeSharePointGroupsToggle() {
    const groupsCheckbox = document.getElementById('show-sharepoint-groups');
    
    if (groupsCheckbox) {
        groupsCheckbox.addEventListener('change', () => {
            const configModule = window.configModule;
            const isChecked = groupsCheckbox.checked;
            
            console.log(`👥 SharePoint Groups display toggled: ${isChecked ? 'SHOW' : 'HIDE'} default groups`);
            
            // Re-apply current filter to update display
            const currentFilter = getCurrentResultsFilter();
            applyResultsFilter(currentFilter);
            
            if (configModule && configModule.showToast) {
                configModule.showToast(`${isChecked ? 'Showing' : 'Hiding'} SharePoint default groups`);
            }
        });
        
        console.log('✅ SharePoint groups toggle initialized');
    } else {
        console.log('⚠️ SharePoint groups checkbox not found');
    }
}

// DIRECT GRANTS TOGGLE FUNCTIONALITY
function initializeDirectGrantsToggle() {
    const directGrantsCheckbox = document.getElementById('show-direct-grants');
    
    if (directGrantsCheckbox) {
        directGrantsCheckbox.addEventListener('change', () => {
            const configModule = window.configModule;
            const isChecked = directGrantsCheckbox.checked;
            
            console.log(`👤 Direct Grants display toggled: ${isChecked ? 'SHOW' : 'HIDE'} direct grants`);
            
            // Re-apply current filter to update display
            const currentFilter = getCurrentResultsFilter();
            applyResultsFilter(currentFilter);
            
            if (configModule && configModule.showToast) {
                configModule.showToast(`${isChecked ? 'Showing' : 'Hiding'} direct user grants`);
            }
        });
        
        console.log('✅ Direct grants toggle initialized');
    } else {
        console.log('⚠️ Direct grants checkbox not found');
    }
}

// HIERARCHICAL TREE STRUCTURE FUNCTIONALITY

// Tree node structure for hierarchical display
class TreeNode {
    constructor(name, path, type, result = null) {
        this.name = name;
        this.path = path;
        this.type = type; // 'site', 'folder', 'file'
        this.result = result; // Original result object for files and folders with permissions
        this.children = new Map(); // Map of child nodes
        this.parent = null;
        this.isExpanded = false;
        this.level = 0;
    }
    
    addChild(childNode) {
        childNode.parent = this;
        childNode.level = this.level + 1;
        this.children.set(childNode.name, childNode);
        return childNode;
    }
    
    getChild(name) {
        return this.children.get(name);
    }
    
    hasChildren() {
        return this.children.size > 0;
    }
    
    getAllDescendants() {
        const descendants = [];
        for (const child of this.children.values()) {
            descendants.push(child);
            descendants.push(...child.getAllDescendants());
        }
        return descendants;
    }
}

// Build hierarchical tree structure from flat results array
function buildResultsTree(results) {
    const configModule = window.configModule;
    if (!configModule || !results || results.length === 0) {
        return new Map();
    }
    
    console.log('🌲 Building hierarchical tree from', results.length, 'results');
    
    const siteRoots = new Map(); // Map of site root nodes
    
    // Process each result and build tree
    results.forEach((result, index) => {
        console.log(`🔍 Processing result ${index + 1}:`, result.itemName, 'at path:', result.itemPath);
        
        // Create or get site root node
        const siteKey = `${result.scanType}_${result.siteName}`;
        let siteNode = siteRoots.get(siteKey);
        
        if (!siteNode) {
            const siteDisplayName = result.scanType === 'onedrive' ? 
                `☁️ ${result.siteName} (OneDrive)` : 
                `📁 ${result.siteName}`;
            
            siteNode = new TreeNode(siteDisplayName, '', 'site', null);
            siteNode.siteInfo = {
                siteName: result.siteName,
                siteUrl: result.siteUrl,
                scanType: result.scanType
            };
            siteRoots.set(siteKey, siteNode);
            console.log(`🏗️ Created site root:`, siteDisplayName);
        }
        
        // Parse the item path to create folder hierarchy
        let itemPath = result.itemPath || `/${result.itemName}`;
        if (!itemPath.startsWith('/')) itemPath = '/' + itemPath;
        
        // Split path into components
        const pathParts = itemPath.split('/').filter(part => part.length > 0);
        console.log(`📂 Path parts for ${result.itemName}:`, pathParts);
        
        let currentNode = siteNode;
        let currentPath = '';
        
        // Navigate/create folder hierarchy (all parts except the last)
        for (let i = 0; i < pathParts.length - 1; i++) {
            const folderName = pathParts[i];
            currentPath += '/' + folderName;
            
            let folderNode = currentNode.getChild(folderName);
            if (!folderNode) {
                folderNode = new TreeNode(folderName, currentPath, 'folder', null);
                currentNode.addChild(folderNode);
                console.log(`📁 Created folder node: ${folderName} at level ${folderNode.level}`);
            }
            currentNode = folderNode;
        }
        
        // Add the final item (file or folder with permissions)
        const itemName = pathParts[pathParts.length - 1] || result.itemName;
        const finalPath = currentPath + '/' + itemName;
        
        let itemNode = currentNode.getChild(itemName);
        if (!itemNode) {
            const nodeType = result.itemType === 'file' ? 'file' : 'folder';
            itemNode = new TreeNode(itemName, finalPath, nodeType, result);
            currentNode.addChild(itemNode);
            console.log(`📄 Added ${nodeType}: ${itemName} at level ${itemNode.level} with ${result.permissions.length} permissions`);
        } else {
            // Merge permissions if item already exists (shouldn't normally happen)
            if (itemNode.result && result.permissions) {
                itemNode.result.permissions = [...(itemNode.result.permissions || []), ...result.permissions];
                console.log(`🔗 Merged permissions for existing item:`, itemName);
            }
        }
    });
    
    console.log(`✅ Tree built with ${siteRoots.size} site roots:`, Array.from(siteRoots.keys()));
    return siteRoots;
}

// Create hierarchical display as cards with expand/collapse functionality
function displayResultsAsHierarchy() {
    const configModule = window.configModule;
    const resultsContainer = document.getElementById('results-container');
    
    if (!configModule || !resultsContainer || !configModule.results.length) {
        console.log('No results to display in hierarchy');
        return;
    }
    
    console.log('🌲 Converting to hierarchical view...');
    
    // Build tree structure
    const siteRoots = buildResultsTree(configModule.results);
    
    // Create hierarchy container
    const hierarchyContainer = document.createElement('div');
    hierarchyContainer.className = 'hierarchy-container';
    hierarchyContainer.id = 'results-hierarchy-container';
    hierarchyContainer.style.cssText = `
        max-height: 600px; 
        overflow-y: auto; 
        border: 1px solid var(--border); 
        border-radius: 8px; 
        background: white;
        padding: 16px;
    `;
    
    // Render each site root and its children
    const currentFilter = getCurrentResultsFilter();
    siteRoots.forEach((siteNode, siteKey) => {
        const siteElement = createHierarchicalNode(siteNode, currentFilter);
        if (siteElement) {
            hierarchyContainer.appendChild(siteElement);
        }
    });
    
    resultsContainer.innerHTML = '';
    resultsContainer.appendChild(hierarchyContainer);
    
    console.log(`✅ Hierarchical view created with ${siteRoots.size} sites`);
}

// Create hierarchical node element with expand/collapse functionality
function createHierarchicalNode(node, currentFilter) {
    const configModule = window.configModule;
    if (!configModule) return null;
    
    // Check if this node or any of its descendants should be shown based on filter
    const shouldShowNode = shouldShowHierarchicalNode(node, currentFilter);
    if (!shouldShowNode) {
        return null;
    }
    
    const nodeDiv = document.createElement('div');
    nodeDiv.className = 'tree-node';
    nodeDiv.style.cssText = `margin-left: ${node.level * 20}px; margin-bottom: 4px;`;
    nodeDiv.dataset.nodeType = node.type;
    nodeDiv.dataset.nodePath = node.path;
    
    // Create node header with expand/collapse control
    const headerDiv = document.createElement('div');
    headerDiv.className = 'tree-node-header';
    headerDiv.style.cssText = `
        display: flex; 
        align-items: center; 
        padding: 8px 12px; 
        border: 1px solid var(--border); 
        border-radius: 6px; 
        background: white; 
        cursor: pointer;
        margin-bottom: 4px;
        transition: background-color 0.2s;
    `;
    
    // Expand/collapse icon
    const expandIcon = document.createElement('span');
    expandIcon.className = 'expand-icon';
    expandIcon.style.cssText = `
        margin-right: 8px; 
        font-size: 12px; 
        transition: transform 0.2s;
        width: 16px;
        display: inline-block;
        text-align: center;
    `;
    
    if (node.hasChildren()) {
        expandIcon.textContent = node.isExpanded ? '▼' : '▶';
        expandIcon.style.cursor = 'pointer';
    } else {
        expandIcon.textContent = '•';
        expandIcon.style.cursor = 'default';
        expandIcon.style.opacity = '0.3';
    }
    
    // Node icon and name
    const nameSpan = document.createElement('span');
    nameSpan.className = 'tree-node-name';
    nameSpan.style.cssText = 'flex: 1; font-weight: 500;';
    
    let nodeIcon = '';
    if (node.type === 'site') {
        nodeIcon = node.siteInfo?.scanType === 'onedrive' ? '☁️' : '🏢';
        nameSpan.innerHTML = `${nodeIcon} ${node.name}`;
        nameSpan.style.fontWeight = '600';
        nameSpan.style.color = 'var(--primary)';
    } else if (node.type === 'folder') {
        nodeIcon = '📁';
        nameSpan.innerHTML = `${nodeIcon} ${node.name}`;
    } else if (node.type === 'file') {
        nodeIcon = '📄';
        nameSpan.innerHTML = `${nodeIcon} ${node.name}`;
    }
    
    // Permission count and actions for items with permissions
    const actionsDiv = document.createElement('div');
    actionsDiv.className = 'tree-node-actions';
    actionsDiv.style.cssText = 'display: flex; align-items: center; gap: 8px; margin-left: 8px;';
    
    if (node.result && node.result.permissions) {
        const filteredPermissions = getFilteredPermissions(node.result.permissions, currentFilter);
        
        // Permission count badge
        const countBadge = document.createElement('span');
        countBadge.className = 'permission-count-badge';
        countBadge.style.cssText = `
            background: var(--primary-light); 
            color: var(--primary); 
            padding: 2px 6px; 
            border-radius: 4px; 
            font-size: 11px; 
            font-weight: 600;
        `;
        countBadge.textContent = `${filteredPermissions.length} permission${filteredPermissions.length !== 1 ? 's' : ''}`;
        
        // Action buttons
        const hasLinks = node.result.permissions.some(p => p.link);
        
        if (hasLinks) {
            const expBtn = document.createElement('button');
            expBtn.className = 'tree-action-btn';
            expBtn.style.cssText = `
                background: var(--purple); 
                color: white; 
                border: none; 
                padding: 4px 8px; 
                border-radius: 4px; 
                font-size: 10px; 
                cursor: pointer;
                margin-right: 4px;
            `;
            expBtn.textContent = 'EXP';
            expBtn.onclick = (e) => {
                e.stopPropagation();
                if (window.permissionsModule?.showExpirationDialog) {
                    window.permissionsModule.showExpirationDialog(node.result, nodeDiv);
                }
            };
            
            const disableLinksBtn = document.createElement('button');
            disableLinksBtn.className = 'tree-action-btn';
            disableLinksBtn.style.cssText = `
                background: var(--orange); 
                color: white; 
                border: none; 
                padding: 4px 8px; 
                border-radius: 4px; 
                font-size: 10px; 
                cursor: pointer;
                margin-right: 4px;
            `;
            disableLinksBtn.textContent = 'DEL LINKS';
            disableLinksBtn.onclick = (e) => {
                e.stopPropagation();
                if (confirm('Remove all sharing links for this item?')) {
                    if (window.permissionsModule?.disableLinks) {
                        window.permissionsModule.disableLinks(node.result, nodeDiv);
                    }
                }
            };
            
            actionsDiv.appendChild(expBtn);
            actionsDiv.appendChild(disableLinksBtn);
        }
        
        const disableAllBtn = document.createElement('button');
        disableAllBtn.className = 'tree-action-btn';
        disableAllBtn.style.cssText = `
            background: var(--danger); 
            color: white; 
            border: none; 
            padding: 4px 8px; 
            border-radius: 4px; 
            font-size: 10px; 
            cursor: pointer;
        `;
        disableAllBtn.textContent = 'DEL ALL';
        disableAllBtn.onclick = (e) => {
            e.stopPropagation();
            if (confirm('Remove all sharing permissions for this item?')) {
                if (window.permissionsModule?.disableAllSharing) {
                    window.permissionsModule.disableAllSharing(node.result, nodeDiv);
                }
            }
        };
        
        actionsDiv.appendChild(countBadge);
        actionsDiv.appendChild(disableAllBtn);
    }
    
    // Add click handler for expand/collapse
    headerDiv.onclick = (e) => {
        e.stopPropagation();
        if (node.hasChildren()) {
            toggleNodeExpansion(node, nodeDiv);
        }
    };
    
    // Hover effects
    headerDiv.onmouseenter = () => {
        headerDiv.style.backgroundColor = 'var(--bg)';
    };
    headerDiv.onmouseleave = () => {
        headerDiv.style.backgroundColor = 'white';
    };
    
    headerDiv.appendChild(expandIcon);
    headerDiv.appendChild(nameSpan);
    headerDiv.appendChild(actionsDiv);
    nodeDiv.appendChild(headerDiv);
    
    // Create children container
    const childrenContainer = document.createElement('div');
    childrenContainer.className = 'tree-node-children';
    childrenContainer.style.cssText = 'margin-left: 0px;';
    childrenContainer.style.display = node.isExpanded ? 'block' : 'none';
    
    // Add children if expanded
    if (node.hasChildren()) {
        node.children.forEach((childNode) => {
            const childElement = createHierarchicalNode(childNode, currentFilter);
            if (childElement) {
                childrenContainer.appendChild(childElement);
            }
        });
    }
    
    nodeDiv.appendChild(childrenContainer);
    
    // Add permission details if this node has permissions
    if (node.result && node.result.permissions && node.result.permissions.length > 0) {
        const permissionsDetails = createPermissionsDetailsForNode(node.result, currentFilter);
        if (permissionsDetails) {
            nodeDiv.appendChild(permissionsDetails);
        }
    }
    
    return nodeDiv;
}

// Toggle expansion state of a tree node
function toggleNodeExpansion(node, nodeElement) {
    node.isExpanded = !node.isExpanded;
    
    const expandIcon = nodeElement.querySelector('.expand-icon');
    const childrenContainer = nodeElement.querySelector('.tree-node-children');
    
    if (expandIcon) {
        expandIcon.textContent = node.isExpanded ? '▼' : '▶';
        expandIcon.style.transform = node.isExpanded ? 'rotate(0deg)' : 'rotate(-90deg)';
    }
    
    if (childrenContainer) {
        childrenContainer.style.display = node.isExpanded ? 'block' : 'none';
    }
    
    console.log(`🌲 ${node.isExpanded ? 'Expanded' : 'Collapsed'} node: ${node.name}`);
}

// Check if hierarchical node should be shown based on filter
function shouldShowHierarchicalNode(node, currentFilter) {
    const configModule = window.configModule;
    if (!configModule) return false;
    
    // Site nodes are always shown
    if (node.type === 'site') return true;
    
    // For files and folders with permissions, check if they match the filter
    if (node.result && node.result.permissions) {
        const filteredPermissions = getFilteredPermissions(node.result.permissions, currentFilter);
        if (filteredPermissions.length > 0) {
            return true;
        }
    }
    
    // For folder nodes without permissions, check if any descendants match the filter
    if (node.type === 'folder' && node.hasChildren()) {
        for (const child of node.children.values()) {
            if (shouldShowHierarchicalNode(child, currentFilter)) {
                return true;
            }
        }
    }
    
    return false;
}

// Create permission details section for a node
function createPermissionsDetailsForNode(result, currentFilter) {
    const configModule = window.configModule;
    if (!configModule) return null;
    
    const filteredPermissions = getFilteredPermissions(result.permissions, currentFilter);
    if (filteredPermissions.length === 0) return null;
    
    const detailsDiv = document.createElement('div');
    detailsDiv.className = 'tree-node-permissions';
    detailsDiv.style.cssText = `
        margin: 4px 0 8px 36px; 
        padding: 12px; 
        background: var(--bg); 
        border-radius: 6px; 
        border: 1px solid var(--border);
    `;
    
    const table = document.createElement('table');
    table.style.cssText = 'width: 100%; font-size: 12px; border-collapse: collapse;';
    
    const thead = document.createElement('thead');
    thead.innerHTML = `
        <tr>
            <th style="text-align: left; padding: 4px 8px; border-bottom: 1px solid var(--border);">Who</th>
            <th style="text-align: left; padding: 4px 8px; border-bottom: 1px solid var(--border);">Role</th>
            <th style="text-align: left; padding: 4px 8px; border-bottom: 1px solid var(--border);">Type</th>
            <th style="text-align: left; padding: 4px 8px; border-bottom: 1px solid var(--border);">Expires</th>
        </tr>
    `;
    table.appendChild(thead);
    
    const tbody = document.createElement('tbody');
    filteredPermissions.forEach((permission) => {
        const tr = document.createElement('tr');
        
        const who = configModule.extractUserFromPermission(permission, configModule.tenantDomains);
        const roles = (permission.roles || []).join(', ') || 'Not specified';
        const classification = configModule.classifyPermission(permission, configModule.tenantDomains);
        const expiration = configModule.extractExpirationDate(permission);
        
        tr.innerHTML = `
            <td style="padding: 4px 8px;">${who}</td>
            <td style="padding: 4px 8px;">${roles}</td>
            <td style="padding: 4px 8px;">
                <span class="${classification === 'external' ? 'external-badge' : 'internal-badge'}">${classification.toUpperCase()}</span>
            </td>
            <td style="padding: 4px 8px;">${expiration}</td>
        `;
        
        tbody.appendChild(tr);
    });
    
    table.appendChild(tbody);
    detailsDiv.appendChild(table);
    
    return detailsDiv;
}

// Add result to hierarchical view during real-time scanning - ENHANCED VERSION WITH FULL FOLDER HIERARCHY
function addResultToHierarchicalView(result) {
    const hierarchyContainer = document.getElementById('results-hierarchy-container');
    if (!hierarchyContainer) {
        console.log('🌲 Hierarchy container not found, initializing...');
        // If hierarchy container doesn't exist, initialize it
        const configModule = window.configModule;
        if (configModule && configModule.results.length > 0) {
            displayResultsAsHierarchy();
        }
        return;
    }
    
    console.log(`🌲 Adding result to full hierarchy incrementally: ${result.itemName}`);
    
    // ENHANCED: Build full folder hierarchy with expand/collapse at every level
    try {
        const configModule = window.configModule;
        if (!configModule) return;
        
        const currentFilter = getCurrentResultsFilter();
        const shouldShow = shouldShowResultBasedOnFilter(result, currentFilter);
        
        if (!shouldShow) {
            console.log(`🔍 HIERARCHY: Result filtered out: ${result.itemName}`);
            return;
        }
        
        // Build full hierarchical path
        const siteKey = `${result.scanType}_${result.siteName}`;
        const siteDisplayName = result.scanType === 'onedrive' ? 
            `☁️ ${result.siteName} (OneDrive)` : 
            `📁 ${result.siteName}`;
        
        // Find or create site container
        let siteContainer = hierarchyContainer.querySelector(`[data-site-key="${siteKey}"]`);
        if (!siteContainer) {
            siteContainer = createHierarchyNode(siteDisplayName, siteKey, 'site', null, 0);
            siteContainer.setAttribute('data-site-key', siteKey);
            hierarchyContainer.appendChild(siteContainer);
        }
        
        // Parse the item path to build folder hierarchy
        let itemPath = result.itemPath || `/${result.itemName}`;
        if (!itemPath.startsWith('/')) itemPath = '/' + itemPath;
        
        // Split path into components
        const pathParts = itemPath.split('/').filter(part => part.length > 0);
        console.log(`📂 Building hierarchy for: ${pathParts.join(' → ')}`);
        
        let currentContainer = siteContainer.querySelector('.tree-node-children');
        let currentPath = '';
        
        // Build folder hierarchy (all parts except the last one)
        for (let i = 0; i < pathParts.length - 1; i++) {
            const folderName = pathParts[i];
            currentPath += '/' + folderName;
            const folderKey = `${siteKey}${currentPath}`;
            
            // Look for existing folder node
            let folderNode = currentContainer.querySelector(`[data-node-key="${folderKey}"]`);
            if (!folderNode) {
                console.log(`📁 Creating folder node: ${folderName} at level ${i + 1}`);
                folderNode = createHierarchyNode(folderName, folderKey, 'folder', null, i + 1);
                folderNode.setAttribute('data-node-key', folderKey);
                currentContainer.appendChild(folderNode);
            }
            
            // Move to the next level
            currentContainer = folderNode.querySelector('.tree-node-children');
        }
        
        // Add the final item (file or folder with permissions)
        const finalItemName = pathParts[pathParts.length - 1] || result.itemName;
        const finalPath = currentPath + '/' + finalItemName;
        const finalKey = `${siteKey}${finalPath}`;
        
        let finalItemNode = currentContainer.querySelector(`[data-node-key="${finalKey}"]`);
        if (!finalItemNode) {
            console.log(`📄 Adding final item: ${finalItemName} with ${result.permissions.length} permissions`);
            finalItemNode = createHierarchyNode(finalItemName, finalKey, result.itemType, result, pathParts.length);
            finalItemNode.setAttribute('data-node-key', finalKey);
            currentContainer.appendChild(finalItemNode);
        }
        
        // Scroll to show the new item
        hierarchyContainer.scrollTop = hierarchyContainer.scrollHeight;
        
        console.log(`✅ HIERARCHY: Added ${result.itemName} to full folder hierarchy`);
        
    } catch (error) {
        console.error('❌ Error in hierarchical update:', error);
        // Fall back to full rebuild only on error
        const configModule = window.configModule;
        if (configModule && configModule.results.length > 0) {
            displayResultsAsHierarchy();
        }
    }
}

// Create a hierarchical tree node with full expand/collapse functionality
function createHierarchyNode(name, nodeKey, nodeType, result = null, level = 0) {
    const nodeDiv = document.createElement('div');
    nodeDiv.className = `tree-node tree-node-${nodeType}`;
    nodeDiv.style.cssText = `margin-left: ${level * 20}px; margin-bottom: 2px;`;
    
    // Create node header with expand/collapse control
    const headerDiv = document.createElement('div');
    headerDiv.className = 'tree-node-header';
    headerDiv.style.cssText = `
        display: flex; 
        align-items: center; 
        padding: 6px 12px; 
        border: 1px solid var(--border); 
        border-radius: 4px; 
        background: white; 
        cursor: pointer;
        margin-bottom: 2px;
        transition: background-color 0.2s;
        ${nodeType === 'site' ? 'font-weight: 600; color: var(--primary);' : ''}
        ${nodeType === 'folder' ? 'font-weight: 500;' : ''}
    `;
    
    // Expand/collapse icon (only for containers)
    const expandIcon = document.createElement('span');
    expandIcon.className = 'expand-icon';
    expandIcon.style.cssText = `
        margin-right: 8px; 
        font-size: 12px; 
        transition: transform 0.2s;
        width: 16px;
        display: inline-block;
        text-align: center;
    `;
    
    const isContainer = nodeType === 'site' || nodeType === 'folder';
    if (isContainer) {
        expandIcon.textContent = '▼'; // Start expanded
        expandIcon.style.cursor = 'pointer';
    } else {
        expandIcon.textContent = '•';
        expandIcon.style.cursor = 'default';
        expandIcon.style.opacity = '0.3';
    }
    
    // Node icon and name
    const nameSpan = document.createElement('span');
    nameSpan.className = 'tree-node-name';
    nameSpan.style.cssText = 'flex: 1; font-size: 14px;';
    
    let nodeIcon = '';
    if (nodeType === 'site') {
        nodeIcon = name.includes('OneDrive') ? '☁️' : '🏢';
    } else if (nodeType === 'folder') {
        nodeIcon = '📁';
    } else if (nodeType === 'file') {
        nodeIcon = '📄';
    }
    nameSpan.textContent = `${nodeIcon} ${name}`;
    
    // Add content for items with permissions
    const actionsDiv = document.createElement('div');
    actionsDiv.className = 'tree-node-actions';
    actionsDiv.style.cssText = 'display: flex; align-items: center; gap: 6px; margin-left: 8px;';
    
    if (result && result.permissions) {
        const currentFilter = getCurrentResultsFilter();
        const filteredPermissions = getFilteredPermissions(result.permissions, currentFilter);
        
        // Permission count badge
        const countBadge = document.createElement('span');
        countBadge.style.cssText = `
            background: var(--primary-light); 
            color: var(--primary); 
            padding: 2px 6px; 
            border-radius: 4px; 
            font-size: 10px; 
            font-weight: 600;
        `;
        countBadge.textContent = `${filteredPermissions.length}`;
        
        // Action buttons
        const hasLinks = result.permissions.some(p => p.link);
        
        if (hasLinks) {
            const expBtn = document.createElement('button');
            expBtn.style.cssText = `
                background: var(--purple); 
                color: white; 
                border: none; 
                padding: 2px 6px; 
                border-radius: 3px; 
                font-size: 9px; 
                cursor: pointer;
            `;
            expBtn.textContent = 'EXP';
            expBtn.onclick = (e) => {
                e.stopPropagation();
                if (window.permissionsModule?.showExpirationDialog) {
                    window.permissionsModule.showExpirationDialog(result, nodeDiv);
                }
            };
            
            const linksBtn = document.createElement('button');
            linksBtn.style.cssText = `
                background: var(--orange); 
                color: white; 
                border: none; 
                padding: 2px 6px; 
                border-radius: 3px; 
                font-size: 9px; 
                cursor: pointer;
            `;
            linksBtn.textContent = 'LINKS';
            linksBtn.onclick = (e) => {
                e.stopPropagation();
                if (confirm('Remove all sharing links?')) {
                    if (window.permissionsModule?.disableLinks) {
                        window.permissionsModule.disableLinks(result, nodeDiv);
                    }
                }
            };
            
            actionsDiv.appendChild(expBtn);
            actionsDiv.appendChild(linksBtn);
        }
        
        const allBtn = document.createElement('button');
        allBtn.style.cssText = `
            background: var(--danger); 
            color: white; 
            border: none; 
            padding: 2px 6px; 
            border-radius: 3px; 
            font-size: 9px; 
            cursor: pointer;
        `;
        allBtn.textContent = 'ALL';
        allBtn.onclick = (e) => {
            e.stopPropagation();
            if (confirm('Remove all permissions?')) {
                if (window.permissionsModule?.disableAllSharing) {
                    window.permissionsModule.disableAllSharing(result, nodeDiv);
                }
            }
        };
        
        actionsDiv.appendChild(countBadge);
        actionsDiv.appendChild(allBtn);
    }
    
    // Create children container for expandable nodes
    const childrenContainer = document.createElement('div');
    childrenContainer.className = 'tree-node-children';
    childrenContainer.style.cssText = 'margin-left: 0px; display: block;';
    
    // Add click handler for expand/collapse (only for containers)
    if (isContainer) {
        headerDiv.onclick = (e) => {
            e.stopPropagation();
            const isExpanded = childrenContainer.style.display === 'block';
            childrenContainer.style.display = isExpanded ? 'none' : 'block';
            expandIcon.textContent = isExpanded ? '▶' : '▼';
            expandIcon.style.transform = isExpanded ? 'rotate(-90deg)' : 'rotate(0deg)';
            console.log(`🌲 ${isExpanded ? 'Collapsed' : 'Expanded'} ${nodeType}: ${name}`);
        };
    }
    
    // Hover effects
    headerDiv.onmouseenter = () => {
        headerDiv.style.backgroundColor = 'var(--bg)';
    };
    headerDiv.onmouseleave = () => {
        headerDiv.style.backgroundColor = 'white';
    };
    
    // Assemble the node
    headerDiv.appendChild(expandIcon);
    headerDiv.appendChild(nameSpan);
    headerDiv.appendChild(actionsDiv);
    nodeDiv.appendChild(headerDiv);
    nodeDiv.appendChild(childrenContainer);
    
    return nodeDiv;
}

// Enhanced view switching to include hierarchy
function switchViewEnhanced(viewType) {
    currentView = viewType;
    const configModule = window.configModule;
    
    if (!configModule || !configModule.results.length) {
        console.log('No results to switch view for');
        return;
    }
    
    console.log(`🔄 Switching to ${viewType} view with ${configModule.results.length} results`);
    
    if (viewType === 'table') {
        displayResultsAsTable();
    } else if (viewType === 'hierarchy') {
        displayResultsAsHierarchy();
    } else {
        displayResultsAsCards();
    }
}

// Enhanced addResultToDisplayEnhanced to support hierarchy view
function addResultToDisplayEnhancedWithHierarchy(result) {
    try {
        const resultsContainer = document.getElementById('results-container');
        if (!resultsContainer) {
            console.error('❌ Results container not found');
            return;
        }
        
        // Ensure results UI components are visible and initialized
        ensureResultsUIInitializedEnhanced(resultsContainer);
        
        const currentFilter = getCurrentResultsFilter();
        const configModule = window.configModule;
        const shouldShow = shouldShowResultBasedOnFilter(result, currentFilter);
        
        console.log(`🔍 REAL-TIME: Adding result to ${currentView} view: ${result.itemName} (${result.itemType}) - Filter: ${currentFilter}, Show: ${shouldShow}`);
        
        if (currentView === 'table') {
            addResultToTableView(result);
        } else if (currentView === 'hierarchy') {
            addResultToHierarchicalView(result);
        } else {
            // Use existing card view logic
            let resultsList = document.getElementById('results-list');
            if (resultsList) {
                const resultDiv = createAndInsertResultElement(result, resultsList, shouldShow);
                resultsList.offsetHeight; // Force reflow
            }
        }
        
        // Update result count display with filter consideration
        updateResultsDisplayWithFilter(currentFilter);
        
        console.log(`✅ Result ${shouldShow ? 'displayed' : 'hidden'} in ${currentView} view: ${result.itemName}`);
        
    } catch (error) {
        console.error('❌ Error in addResultToDisplayEnhancedWithHierarchy:', error);
    }
}

// Export functions for use in other modules
window.resultsModule = {
    // Filtering functions
    initializeResultsFiltering,
    applyResultsFilter,
    shouldShowResult,
    
    // Real-time display functions
    addResultToDisplay: addResultToDisplayEnhancedWithHierarchy,
    getCurrentResultsFilter,
    shouldShowResultBasedOnFilter,
    updateResultsDisplayWithFilter,
    ensureResultsUIInitialized,
    createAndInsertResultElement,
    updateResultsDisplay,
    
    // Table view functions
    initializeViewToggle,
    switchView: switchViewEnhanced,
    displayResultsAsTable,
    displayResultsAsCards,
    createTableRow,
    addResultToTableView,
    
    // Hierarchical view functions
    buildResultsTree,
    displayResultsAsHierarchy,
    createHierarchicalNode,
    toggleNodeExpansion,
    shouldShowHierarchicalNode,
    addResultToHierarchicalView,
    
    // Permission refresh
    refreshItemPermissions,
    
    // SharePoint groups toggle
    initializeSharePointGroupsToggle,
    
    // Direct grants toggle
    initializeDirectGrantsToggle
};
