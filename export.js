// export.js - Export Module for SharePoint & OneDrive Scanner v3.0
// Handles Excel export functionality and bulk operations via CSV import/export

// EXTRACT OWNERS FROM RESULT FOR EXPORT - ENHANCED TO MATCH RESULTS.JS LOGIC
function extractOwnersForExport(result, configModule) {
    console.log('🚨 EXPORT OWNERS: FUNCTION CALLED!', result ? result.itemName : 'NO RESULT');
    
    if (!result || !result.permissions || !configModule) {
        console.log('🚨 EXPORT OWNERS: Missing dependencies - result:', !!result, 'permissions:', !!(result && result.permissions), 'configModule:', !!configModule);
        return 'N/A';
    }
    
    const owners = [];
    // Expanded role keywords to catch all variations
    const ownerRoleKeywords = [
        'owner', 'owners', 'full control', 'fullcontrol', 'edit', 'write', 'manage', 'control'
    ];
    
    console.log(`🚨 EXPORT OWNERS: Processing ${result.permissions.length} permissions for ${result.itemName}`);
    
    result.permissions.forEach((permission, index) => {
        console.log(`🚨 EXPORT OWNERS: --- Permission ${index + 1} ---`);
        console.log(`🚨 EXPORT OWNERS: Permission roles:`, permission.roles);
        
        const roles = permission.roles || [];
        
        // Check each role
        let hasOwnerRole = false;
        for (let i = 0; i < roles.length; i++) {
            const role = roles[i];
            const roleLower = role.toLowerCase();
            console.log(`🚨 EXPORT OWNERS: Checking role #${i + 1}: "${role}" -> "${roleLower}"`);
            
            // Check against each keyword
            for (let j = 0; j < ownerRoleKeywords.length; j++) {
                const keyword = ownerRoleKeywords[j];
                if (roleLower.includes(keyword)) {
                    console.log(`🚨 EXPORT OWNERS: MATCH! Role "${role}" contains keyword "${keyword}"`);
                    hasOwnerRole = true;
                    break;
                }
            }
            if (hasOwnerRole) break;
        }
        
        console.log(`🚨 EXPORT OWNERS: Permission ${index + 1} hasOwnerRole: ${hasOwnerRole}`);
        
        if (hasOwnerRole) {
            const who = configModule.extractUserFromPermission(permission, configModule.tenantDomains);
            console.log(`🚨 EXPORT OWNERS: Extracted user: "${who}"`);
            
            if (who && who !== '(direct grant)' && who !== 'Anyone (Anonymous Link)' && !owners.includes(who)) {
                owners.push(who);
                console.log(`🚨 EXPORT OWNERS: ✅ ADDED OWNER: "${who}"`);
            } else {
                console.log(`🚨 EXPORT OWNERS: ❌ SKIPPED: "${who}" (conditions not met)`);
            }
        }
    });
    
    const result_text = owners.length > 0 ? owners.join(', ') : 'No owners found';
    console.log(`🚨 EXPORT OWNERS: FINAL RESULT: "${result_text}"`);
    return result_text;
}

// EXCEL EXPORT FUNCTIONALITY
function exportResults() {
    const configModule = window.configModule;
    const resultsModule = window.resultsModule;
    
    if (!configModule) {
        console.error('Config module not available for export');
        return;
    }
    
    if (!configModule.results || configModule.results.length === 0) {
        alert('No results to export');
        return;
    }
    
    console.log('🔍 EXPORT: Starting export with SharePoint groups filter');
    const showSharePointGroups = configModule.shouldShowSharePointGroups();
    console.log('🔍 EXPORT: Show SharePoint groups:', showSharePointGroups);
    
    const exportData = [];
    let totalPermissions = 0;
    let filteredPermissions = 0;
    
    configModule.results.forEach(result => {
        // Extract owners for this result
        const owners = extractOwnersForExport(result, configModule);
        
        // Get current results filter to honor it in export
        const currentFilter = resultsModule ? resultsModule.getCurrentResultsFilter() : 'all';
        console.log('🔍 EXPORT: Using results filter:', currentFilter);
        
        // Apply the same filtering logic used in the display
        let permissionsToExport = result.permissions;
        if (resultsModule && resultsModule.getFilteredPermissions) {
            permissionsToExport = resultsModule.getFilteredPermissions(result.permissions, currentFilter);
            console.log(`🔍 EXPORT: Filtered ${result.permissions.length} -> ${permissionsToExport.length} permissions for ${result.itemName}`);
        }
        
        permissionsToExport.forEach(permission => {
            totalPermissions++;
            
            // Check if this is a default SharePoint group permission that should be filtered
            let shouldInclude = true;
            
            if (!showSharePointGroups) {
                // Check for regular groups
                if (permission.grantedToV2 && permission.grantedToV2.group) {
                    const groupName = permission.grantedToV2.group.displayName || permission.grantedToV2.group.email || '';
                    if (configModule.isDefaultSharePointGroup(groupName)) {
                        console.log(`🚫 EXPORT: Filtering out default SharePoint group: "${groupName}"`);
                        shouldInclude = false;
                    }
                }
                
                // Check for site groups
                if (permission.grantedToV2 && permission.grantedToV2.siteGroup) {
                    const siteGroupName = permission.grantedToV2.siteGroup.displayName || permission.grantedToV2.siteGroup.loginName || '';
                    if (configModule.isDefaultSharePointGroup(siteGroupName)) {
                        console.log(`🚫 EXPORT: Filtering out default SharePoint site group: "${siteGroupName}"`);
                        shouldInclude = false;
                    }
                }
                
                // Check in grantedToIdentitiesV2 for groups (older API format)
                if (Array.isArray(permission.grantedToIdentitiesV2)) {
                    const hasDefaultGroup = permission.grantedToIdentitiesV2.some(g => {
                        if (g.group) {
                            const groupName = g.group.displayName || g.group.email || '';
                            const isDefault = configModule.isDefaultSharePointGroup(groupName);
                            if (isDefault) {
                                console.log(`🚫 EXPORT: Filtering out default SharePoint group (grantedToIdentitiesV2): "${groupName}"`);
                            }
                            return isDefault;
                        }
                        return false;
                    });
                    
                    if (hasDefaultGroup) {
                        shouldInclude = false;
                    }
                }
            }
            
            if (shouldInclude) {
                filteredPermissions++;
                const who = configModule.extractUserFromPermission(permission, configModule.tenantDomains);
                const roles = (permission.roles || []).join(', ') || 'Not specified';
                const expiration = configModule.extractExpirationDate(permission);
                const classification = configModule.classifyPermission(permission, configModule.tenantDomains);
                
                // Enhanced direct grants analysis for export using improved config functions
                const isDirectGrantPermission = configModule.isDirectGrant(permission);
                const directGrantDetails = isDirectGrantPermission ? configModule.extractDirectGrantDetails(permission, configModule.tenantDomains) : null;
                const directGrantDisplay = directGrantDetails ? configModule.formatDirectGrantDisplay(directGrantDetails) : null;
                
                exportData.push({
                    'Source': result.scanType === 'onedrive' ? 'OneDrive' : 'SharePoint',
                    'Site Name': result.siteName || 'OneDrive',
                    'Site URL': result.siteUrl || 'Personal OneDrive',
                    'Item ID': result.itemId,
                    'Item Name': result.itemName,
                    'Item Path': result.itemPath,
                    'Item Type': result.itemType || 'folder',
                    'Owners': owners,
                    'Who Has Access': who,
                    'Permission Level': roles,
                    'Sharing Type': classification.toUpperCase(),
                    'Link Expiration': expiration,
                    
                    // Enhanced Direct Grants Information - Using improved config module functions
                    'Is Direct Grant': isDirectGrantPermission ? 'YES' : 'NO',
                    'Risk Level': directGrantDetails ? directGrantDetails.riskLevel : 'N/A',
                    'Risk Factors': directGrantDetails && directGrantDetails.riskFactors.length > 0 ? 
                        directGrantDetails.riskFactors.join('; ') : 'None',
                    'User Display Name': directGrantDetails ? directGrantDetails.userDisplayName : 
                        (permission.grantedTo && permission.grantedTo.user ? permission.grantedTo.user.displayName : 'N/A'),
                    'User Email': directGrantDetails ? directGrantDetails.userEmail : 
                        (permission.grantedTo && permission.grantedTo.user ? permission.grantedTo.user.email : 'N/A'),
                    'User ID': directGrantDetails ? directGrantDetails.userId : 'N/A',
                    'Permission ID': directGrantDetails ? directGrantDetails.permissionId : (permission.id || 'N/A'),
                    'Is External User': directGrantDetails ? (directGrantDetails.isExternal ? 'YES' : 'NO') : 'N/A',
                    'Is Internal User': directGrantDetails ? (directGrantDetails.isInternal ? 'YES' : 'NO') : 'N/A',
                    'Permission Type': directGrantDetails ? directGrantDetails.permissionType : 
                        (permission.link ? 'Link-based' : 'Group/Other'),
                    'Permission Scope': directGrantDetails ? directGrantDetails.scope : 'N/A',
                    'Granted DateTime': directGrantDetails ? directGrantDetails.grantedDateTime : 'N/A',
                    'Expiration DateTime': directGrantDetails ? directGrantDetails.expirationDateTime : 'N/A',
                    'Inherited From': directGrantDetails && directGrantDetails.inheritedFrom ? 
                        directGrantDetails.inheritedFrom.name : 'No',
                    'Inherited From ID': directGrantDetails && directGrantDetails.inheritedFrom ? 
                        directGrantDetails.inheritedFrom.id : 'N/A',
                    'Inherited From URL': directGrantDetails && directGrantDetails.inheritedFrom ? 
                        directGrantDetails.inheritedFrom.webUrl : 'N/A',
                    'Has Application': directGrantDetails ? (directGrantDetails.hasApplication ? 'YES' : 'NO') : 'N/A',
                    'Application Name': directGrantDetails && directGrantDetails.hasApplication ? 
                        directGrantDetails.applicationDisplayName || 'Unknown Application' : 'N/A',
                    'Has Link': directGrantDetails ? (directGrantDetails.hasLink ? 'YES' : 'NO') : 'N/A',
                    'Grant Roles': directGrantDetails && directGrantDetails.roles ? 
                        directGrantDetails.roles.join('; ') : (roles || 'Not specified')
                });
            }
        });
    });
    
    console.log(`🔍 EXPORT: Filtered ${totalPermissions} -> ${filteredPermissions} permissions for export`);
    
    try {
        const ws = XLSX.utils.json_to_sheet(exportData);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, 'Enhanced Sharing Report');
        
        const filename = `sharepoint_onedrive_enhanced_sharing_${new Date().toISOString().slice(0,10)}.xlsx`;
        XLSX.writeFile(wb, filename);
        
        if (configModule.showToast) {
            configModule.showToast(`Exported ${exportData.length} sharing records to ${filename}`);
        }
    } catch (error) {
        console.error('Export error:', error);
        alert('Export failed: ' + error.message);
    }
}

// BULK OPERATIONS CSV TEMPLATE DOWNLOAD
function downloadCSVTemplate() {
    const template = [
        ['ItemID', 'Action', 'UserEmail', 'Role', 'LinkScope', 'ExpirationDate'],
        ['example-item-id-1', 'add', 'external@example.com', 'read', '', '2024-12-31'],
        ['example-item-id-2', 'remove', 'user@external.com', '', '', ''],
        ['example-item-id-3', 'modify', '', '', 'users', '2024-06-30'],
        ['', 'Actions: add, remove, modify', '', 'Roles: read, write, owner', 'LinkScope: anonymous, users, organization', 'Format: YYYY-MM-DD or blank']
    ];
    
    const ws = XLSX.utils.aoa_to_sheet(template);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Bulk Operations Template');
    
    const filename = `sharepoint_bulk_operations_template_${new Date().toISOString().slice(0,10)}.csv`;
    XLSX.writeFile(wb, filename);
    
    const configModule = window.configModule;
    if (configModule && configModule.showToast) {
        configModule.showToast('CSV template downloaded');
    }
}

// CSV UPLOAD AND PROCESSING
function handleCSVUpload(event) {
    const configModule = window.configModule;
    const uiModule = window.uiModule;
    
    if (!configModule) {
        console.error('Config module not available for CSV upload');
        return;
    }
    
    const file = event.target.files[0];
    if (!file) return;
    
    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const csv = e.target.result;
            const workbook = XLSX.read(csv, { type: 'string' });
            const worksheet = workbook.Sheets[workbook.SheetNames[0]];
            const data = XLSX.utils.sheet_to_json(worksheet);
            
            const validData = data.filter(row => row.ItemID && row.Action);
            
            if (validData.length === 0) {
                configModule.showToast('No valid data found in CSV');
                return;
            }
            
            // Validate CSV data
            const validation = configModule.validateCSVData(validData);
            if (!validation.isValid) {
                alert('CSV Validation Error: ' + validation.message);
                return;
            }
            
            configModule.bulkCsvData = validData;
            
            if (uiModule && uiModule.displayCSVPreview) {
                uiModule.displayCSVPreview(validData);
            }
            
            const processBulkBtn = document.getElementById('process-bulk-btn');
            if (processBulkBtn) {
                processBulkBtn.disabled = false;
            }
            
            configModule.showToast(`Loaded ${validData.length} bulk operations from CSV`);
            
        } catch (error) {
            console.error('Error parsing CSV:', error);
            if (configModule.showToast) {
                configModule.showToast('Error parsing CSV file');
            }
        }
    };
    reader.readAsText(file);
}

// BULK OPERATIONS PROCESSING
async function processBulkOperations() {
    const configModule = window.configModule;
    const permissionsModule = window.permissionsModule;
    
    if (!configModule || !permissionsModule) {
        console.error('Required modules not available for bulk operations');
        return;
    }
    
    if (configModule.bulkCsvData.length === 0) {
        configModule.showToast('No bulk operations to process');
        return;
    }
    
    const processBtn = document.getElementById('process-bulk-btn');
    const bulkStatus = document.getElementById('bulk-status');
    
    if (processBtn) {
        processBtn.disabled = true;
        processBtn.innerText = 'Processing...';
    }
    
    let processed = 0;
    let successful = 0;
    let failed = 0;
    
    for (const operation of configModule.bulkCsvData) {
        if (configModule.controller.stop) break;
        
        try {
            processed++;
            if (bulkStatus) {
                bulkStatus.innerText = `Processing ${processed}/${configModule.bulkCsvData.length}...`;
            }
            
            const success = await permissionsModule.processSingleBulkOperation(operation);
            if (success) {
                successful++;
            } else {
                failed++;
            }
            
        } catch (error) {
            console.error(`Bulk operation failed for ${operation.ItemID}:`, error);
            failed++;
        }
        
        // Small delay between operations
        await new Promise(resolve => setTimeout(resolve, 200));
    }
    
    if (processBtn) {
        processBtn.disabled = false;
        processBtn.innerText = '⚡ Process Bulk Changes';
    }
    
    if (bulkStatus) {
        bulkStatus.innerText = `Complete: ${successful} successful, ${failed} failed`;
    }
    
    configModule.showToast(`Bulk operations complete: ${successful} successful, ${failed} failed`);
}

// BULK OPERATIONS UI INITIALIZATION
function initializeBulkOperationsHandlers() {
    // CSV template download button
    const bulkTemplateBtn = document.querySelector('.bulk-controls .btn');
    if (bulkTemplateBtn) {
        bulkTemplateBtn.addEventListener('click', downloadCSVTemplate);
    }

    // CSV upload handler
    const csvUpload = document.getElementById('csv-upload');
    if (csvUpload) {
        csvUpload.addEventListener('change', handleCSVUpload);
    }

    // Process bulk operations button
    const processBulkBtn = document.getElementById('process-bulk-btn');
    if (processBulkBtn) {
        processBulkBtn.addEventListener('click', processBulkOperations);
    }
}

// EXPORT BUTTON INITIALIZATION
function initializeExportHandlers() {
    const exportBtn = document.getElementById('export-btn');
    if (exportBtn) {
        exportBtn.addEventListener('click', exportResults);
    }
}

// CSV VALIDATION UTILITIES
function validateBulkCSVFormat(data) {
    if (!Array.isArray(data) || data.length === 0) {
        return { isValid: false, message: 'CSV data is empty or invalid format' };
    }
    
    const requiredColumns = ['ItemID', 'Action'];
    const optionalColumns = ['UserEmail', 'Role', 'LinkScope', 'ExpirationDate'];
    const validActions = ['add', 'remove', 'modify'];
    const validRoles = ['read', 'write', 'owner'];
    const validLinkScopes = ['anonymous', 'users', 'organization'];
    
    // Check required columns
    const firstRow = data[0];
    for (const column of requiredColumns) {
        if (!(column in firstRow)) {
            return { isValid: false, message: `Missing required column: ${column}` };
        }
    }
    
    // Validate each row
    for (let i = 0; i < data.length; i++) {
        const row = data[i];
        
        // Validate ItemID
        if (!row.ItemID || row.ItemID.trim() === '') {
            return { isValid: false, message: `Row ${i + 1}: ItemID is required` };
        }
        
        // Validate Action
        if (!row.Action || !validActions.includes(row.Action.toLowerCase())) {
            return { isValid: false, message: `Row ${i + 1}: Invalid action. Must be: ${validActions.join(', ')}` };
        }
        
        // Validate Role if provided
        if (row.Role && row.Role.trim() !== '' && !validRoles.includes(row.Role.toLowerCase())) {
            return { isValid: false, message: `Row ${i + 1}: Invalid role. Must be: ${validRoles.join(', ')}` };
        }
        
        // Validate LinkScope if provided
        if (row.LinkScope && row.LinkScope.trim() !== '' && !validLinkScopes.includes(row.LinkScope.toLowerCase())) {
            return { isValid: false, message: `Row ${i + 1}: Invalid link scope. Must be: ${validLinkScopes.join(', ')}` };
        }
        
        // Validate ExpirationDate if provided
        if (row.ExpirationDate && row.ExpirationDate.trim() !== '') {
            const date = new Date(row.ExpirationDate);
            if (isNaN(date.getTime())) {
                return { isValid: false, message: `Row ${i + 1}: Invalid expiration date format. Use YYYY-MM-DD` };
            }
            if (date < new Date()) {
                return { isValid: false, message: `Row ${i + 1}: Expiration date cannot be in the past` };
            }
        }
        
        // Action-specific validations
        switch (row.Action.toLowerCase()) {
            case 'add':
                if (!row.UserEmail && !row.LinkScope) {
                    return { isValid: false, message: `Row ${i + 1}: Add action requires either UserEmail or LinkScope` };
                }
                if (row.UserEmail && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(row.UserEmail)) {
                    return { isValid: false, message: `Row ${i + 1}: Invalid email format` };
                }
                break;
            case 'remove':
                if (!row.UserEmail) {
                    return { isValid: false, message: `Row ${i + 1}: Remove action requires UserEmail` };
                }
                break;
            case 'modify':
                if (!row.LinkScope && !row.ExpirationDate) {
                    return { isValid: false, message: `Row ${i + 1}: Modify action requires either LinkScope or ExpirationDate` };
                }
                break;
        }
    }
    
    return { isValid: true, message: 'CSV format is valid' };
}

// EXPORT STATISTICS AND REPORTING
function generateExportStatistics(results) {
    const configModule = window.configModule;
    if (!configModule || !results || results.length === 0) {
        return null;
    }
    
    const stats = {
        totalItems: results.length,
        sharePointItems: results.filter(r => r.scanType === 'sharepoint').length,
        oneDriveItems: results.filter(r => r.scanType === 'onedrive').length,
        foldersCount: results.filter(r => r.itemType === 'folder').length,
        filesCount: results.filter(r => r.itemType === 'file').length,
        externalSharingCount: 0,
        internalSharingCount: 0,
        mixedSharingCount: 0,
        totalPermissions: 0,
        linkPermissions: 0,
        directPermissions: 0,
        expiredPermissions: 0
    };
    
    results.forEach(result => {
        let hasExternal = false;
        let hasInternal = false;
        
        result.permissions.forEach(permission => {
            stats.totalPermissions++;
            
            if (permission.link) {
                stats.linkPermissions++;
            } else {
                stats.directPermissions++;
            }
            
            // Check expiration
            const expiration = configModule.extractExpirationDate(permission);
            if (expiration !== 'No expiration') {
                try {
                    const expDate = new Date(expiration);
                    if (expDate < new Date()) {
                        stats.expiredPermissions++;
                    }
                } catch (e) {
                    // Ignore parsing errors
                }
            }
            
            // Classify permission
            const classification = configModule.classifyPermission(permission, configModule.tenantDomains);
            if (classification === 'external') hasExternal = true;
            if (classification === 'internal') hasInternal = true;
        });
        
        // Classify result overall
        if (hasExternal && hasInternal) {
            stats.mixedSharingCount++;
        } else if (hasExternal) {
            stats.externalSharingCount++;
        } else if (hasInternal) {
            stats.internalSharingCount++;
        }
    });
    
    return stats;
}

// ENHANCED EXPORT WITH STATISTICS SHEET
function exportResultsWithStatistics() {
    const configModule = window.configModule;
    
    if (!configModule) {
        console.error('Config module not available for export');
        return;
    }
    
    if (!configModule.results || configModule.results.length === 0) {
        alert('No results to export');
        return;
    }
    
    try {
        const exportData = [];
        
        // Generate main export data with filtering
        const resultsModule = window.resultsModule;
        const showSharePointGroups = configModule.shouldShowSharePointGroups();
        
        configModule.results.forEach(result => {
            // Extract owners for this result
            const owners = extractOwnersForExport(result, configModule);
            
            // Get current results filter to honor it in export
            const currentFilter = resultsModule ? resultsModule.getCurrentResultsFilter() : 'all';
            
            // Apply the same filtering logic used in the display
            let permissionsToExport = result.permissions;
            if (resultsModule && resultsModule.getFilteredPermissions) {
                permissionsToExport = resultsModule.getFilteredPermissions(result.permissions, currentFilter);
            }
            
            permissionsToExport.forEach(permission => {
                // Check if this is a default SharePoint group permission that should be filtered
                let shouldInclude = true;
                
                if (!showSharePointGroups) {
                    // Check for regular groups
                    if (permission.grantedToV2 && permission.grantedToV2.group) {
                        const groupName = permission.grantedToV2.group.displayName || permission.grantedToV2.group.email || '';
                        if (configModule.isDefaultSharePointGroup(groupName)) {
                            shouldInclude = false;
                        }
                    }
                    
                    // Check for site groups
                    if (permission.grantedToV2 && permission.grantedToV2.siteGroup) {
                        const siteGroupName = permission.grantedToV2.siteGroup.displayName || permission.grantedToV2.siteGroup.loginName || '';
                        if (configModule.isDefaultSharePointGroup(siteGroupName)) {
                            shouldInclude = false;
                        }
                    }
                    
                    // Check in grantedToIdentitiesV2 for groups (older API format)
                    if (Array.isArray(permission.grantedToIdentitiesV2)) {
                        const hasDefaultGroup = permission.grantedToIdentitiesV2.some(g => {
                            if (g.group) {
                                const groupName = g.group.displayName || g.group.email || '';
                                return configModule.isDefaultSharePointGroup(groupName);
                            }
                            return false;
                        });
                        
                        if (hasDefaultGroup) {
                            shouldInclude = false;
                        }
                    }
                }
                
                if (shouldInclude) {
                    const who = configModule.extractUserFromPermission(permission, configModule.tenantDomains);
                    const roles = (permission.roles || []).join(', ') || 'Not specified';
                    const expiration = configModule.extractExpirationDate(permission);
                    const classification = configModule.classifyPermission(permission, configModule.tenantDomains);
                    
                    // Enhanced direct grants analysis for export using improved config functions
                    const isDirectGrantPermission = configModule.isDirectGrant(permission);
                    const directGrantDetails = isDirectGrantPermission ? configModule.extractDirectGrantDetails(permission, configModule.tenantDomains) : null;
                    const directGrantDisplay = directGrantDetails ? configModule.formatDirectGrantDisplay(directGrantDetails) : null;
                    
                    exportData.push({
                        'Source': result.scanType === 'onedrive' ? 'OneDrive' : 'SharePoint',
                        'Site Name': result.siteName || 'OneDrive',
                        'Site URL': result.siteUrl || 'Personal OneDrive',
                        'Item ID': result.itemId,
                        'Item Name': result.itemName,
                        'Item Path': result.itemPath,
                        'Item Type': result.itemType || 'folder',
                        'Owners': owners,
                        'Who Has Access': who,
                        'Permission Level': roles,
                        'Sharing Type': classification.toUpperCase(),
                        'Link Expiration': expiration,
                        
                        // Enhanced Direct Grants Information - Using improved config module functions
                        'Is Direct Grant': isDirectGrantPermission ? 'YES' : 'NO',
                        'Risk Level': directGrantDetails ? directGrantDetails.riskLevel : 'N/A',
                        'Risk Factors': directGrantDetails && directGrantDetails.riskFactors.length > 0 ? 
                            directGrantDetails.riskFactors.join('; ') : 'None',
                        'User Display Name': directGrantDetails ? directGrantDetails.userDisplayName : 
                            (permission.grantedTo && permission.grantedTo.user ? permission.grantedTo.user.displayName : 'N/A'),
                        'User Email': directGrantDetails ? directGrantDetails.userEmail : 
                            (permission.grantedTo && permission.grantedTo.user ? permission.grantedTo.user.email : 'N/A'),
                        'User ID': directGrantDetails ? directGrantDetails.userId : 'N/A',
                        'Permission ID': directGrantDetails ? directGrantDetails.permissionId : (permission.id || 'N/A'),
                        'Is External User': directGrantDetails ? (directGrantDetails.isExternal ? 'YES' : 'NO') : 'N/A',
                        'Is Internal User': directGrantDetails ? (directGrantDetails.isInternal ? 'YES' : 'NO') : 'N/A',
                        'Permission Type': directGrantDetails ? directGrantDetails.permissionType : 
                            (permission.link ? 'Link-based' : 'Group/Other'),
                        'Permission Scope': directGrantDetails ? directGrantDetails.scope : 'N/A',
                        'Granted DateTime': directGrantDetails ? directGrantDetails.grantedDateTime : 'N/A',
                        'Expiration DateTime': directGrantDetails ? directGrantDetails.expirationDateTime : 'N/A',
                        'Inherited From': directGrantDetails && directGrantDetails.inheritedFrom ? 
                            directGrantDetails.inheritedFrom.name : 'No',
                        'Inherited From ID': directGrantDetails && directGrantDetails.inheritedFrom ? 
                            directGrantDetails.inheritedFrom.id : 'N/A',
                        'Inherited From URL': directGrantDetails && directGrantDetails.inheritedFrom ? 
                            directGrantDetails.inheritedFrom.webUrl : 'N/A',
                        'Has Application': directGrantDetails ? (directGrantDetails.hasApplication ? 'YES' : 'NO') : 'N/A',
                        'Application Name': directGrantDetails && directGrantDetails.hasApplication ? 
                            directGrantDetails.applicationDisplayName || 'Unknown Application' : 'N/A',
                        'Has Link': directGrantDetails ? (directGrantDetails.hasLink ? 'YES' : 'NO') : 'N/A',
                        'Grant Roles': directGrantDetails && directGrantDetails.roles ? 
                            directGrantDetails.roles.join('; ') : (roles || 'Not specified')
                    });
                }
            });
        });
        
        // Generate statistics
        const stats = generateExportStatistics(configModule.results);
        const statisticsData = [];
        
        if (stats) {
            statisticsData.push(['Metric', 'Value']);
            statisticsData.push(['Total Items with Sharing', stats.totalItems]);
            statisticsData.push(['SharePoint Items', stats.sharePointItems]);
            statisticsData.push(['OneDrive Items', stats.oneDriveItems]);
            statisticsData.push(['Folders', stats.foldersCount]);
            statisticsData.push(['Files', stats.filesCount]);
            statisticsData.push(['']);
            statisticsData.push(['External Sharing Items', stats.externalSharingCount]);
            statisticsData.push(['Internal Sharing Items', stats.internalSharingCount]);
            statisticsData.push(['Mixed Sharing Items', stats.mixedSharingCount]);
            statisticsData.push(['']);
            statisticsData.push(['Total Permissions', stats.totalPermissions]);
            statisticsData.push(['Link-based Permissions', stats.linkPermissions]);
            statisticsData.push(['Direct Permissions', stats.directPermissions]);
            statisticsData.push(['Expired Permissions', stats.expiredPermissions]);
            statisticsData.push(['']);
            statisticsData.push(['Scan Date', new Date().toLocaleDateString()]);
            statisticsData.push(['Scan Time', new Date().toLocaleTimeString()]);
            statisticsData.push(['Scanner Version', configModule.APP_CONFIG ? configModule.APP_CONFIG.version : '3.0.0']);
        }
        
        // Create workbook with multiple sheets
        const wb = XLSX.utils.book_new();
        
        // Main data sheet
        const ws1 = XLSX.utils.json_to_sheet(exportData);
        XLSX.utils.book_append_sheet(wb, ws1, 'Sharing Report');
        
        // Statistics sheet
        if (statisticsData.length > 0) {
            const ws2 = XLSX.utils.aoa_to_sheet(statisticsData);
            XLSX.utils.book_append_sheet(wb, ws2, 'Statistics');
        }
        
        const filename = `sharepoint_onedrive_enhanced_sharing_${new Date().toISOString().slice(0,10)}.xlsx`;
        XLSX.writeFile(wb, filename);
        
        if (configModule.showToast) {
            configModule.showToast(`Exported ${exportData.length} sharing records with statistics to ${filename}`);
        }
        
    } catch (error) {
        console.error('Enhanced export error:', error);
        alert('Export failed: ' + error.message);
    }
}

// Initialize all export functionality
function initializeExportModule() {
    console.log('📊 Initializing export functionality...');
    
    try {
        initializeExportHandlers();
        initializeBulkOperationsHandlers();
        
        console.log('✅ Export module initialized successfully');
    } catch (error) {
        console.error('❌ Error initializing export module:', error);
    }
}

// REAL-TIME CSV EXPORT FUNCTIONALITY
let csvFileHandle = null;
let csvWriterStream = null;
let csvRowCount = 0;
let realtimeCsvEnabled = false;

// Initialize CSV file for real-time export
async function initializeCsvExport() {
    const configModule = window.configModule;
    
    if (!configModule) {
        console.error('Config module not available for CSV export');
        return false;
    }
    
    try {
        // Check if File System Access API is supported
        if (!('showSaveFilePicker' in window)) {
            alert('Real-time CSV export requires a modern browser that supports the File System Access API (Chrome, Edge, etc.)');
            return false;
        }
        
        // Show file picker for user to choose location
        const suggestedName = `sharepoint_realtime_export_${new Date().toISOString().slice(0,10)}.csv`;
        
        csvFileHandle = await window.showSaveFilePicker({
            suggestedName: suggestedName,
            types: [{
                description: 'CSV files',
                accept: { 'text/csv': ['.csv'] }
            }]
        });
        
        // Create writable stream
        csvWriterStream = await csvFileHandle.createWritable();
        
        // Write CSV header
        const headers = [
            'Timestamp',
            'Source',
            'Site Name', 
            'Site URL',
            'Item ID',
            'Item Name',
            'Item Path',
            'Item Type',
            'Owners',
            'Who Has Access',
            'Permission Level',
            'Sharing Type',
            'Link Expiration'
        ].join(',') + '\n';
        
        await csvWriterStream.write(headers);
        csvRowCount = 0;
        
        console.log('📝 Real-time CSV export initialized:', csvFileHandle.name);
        configModule.showToast(`CSV export initialized: ${csvFileHandle.name}`);
        
        return true;
        
    } catch (error) {
        console.error('Failed to initialize CSV export:', error);
        if (error.name === 'AbortError') {
            configModule.showToast('CSV export cancelled by user');
        } else {
            configModule.showToast('Failed to setup CSV export: ' + error.message);
        }
        return false;
    }
}

// Write a result to the CSV file in real-time
async function writeResultToCsv(result) {
    const configModule = window.configModule;
    const resultsModule = window.resultsModule;
    
    if (!csvWriterStream || !realtimeCsvEnabled || !configModule) {
        return;
    }
    
    try {
        const timestamp = new Date().toISOString();
        const showSharePointGroups = configModule.shouldShowSharePointGroups();
        
        // Get current results filter to honor it in real-time export
        const currentFilter = resultsModule ? resultsModule.getCurrentResultsFilter() : 'all';
        
        // Apply the same filtering logic used in the display
        let permissionsToExport = result.permissions;
        if (resultsModule && resultsModule.getFilteredPermissions) {
            permissionsToExport = resultsModule.getFilteredPermissions(result.permissions, currentFilter);
        }
        
        // Extract owners for this result (using existing function)
        const owners = extractOwnersForExport(result, configModule);
        
        for (const permission of permissionsToExport) {
            // Apply SharePoint groups filter
            let shouldInclude = true;
            
            if (!showSharePointGroups) {
                // Check for regular groups
                if (permission.grantedToV2 && permission.grantedToV2.group) {
                    const groupName = permission.grantedToV2.group.displayName || permission.grantedToV2.group.email || '';
                    if (configModule.isDefaultSharePointGroup(groupName)) {
                        shouldInclude = false;
                    }
                }
                
                // Check for site groups
                if (permission.grantedToV2 && permission.grantedToV2.siteGroup) {
                    const siteGroupName = permission.grantedToV2.siteGroup.displayName || permission.grantedToV2.siteGroup.loginName || '';
                    if (configModule.isDefaultSharePointGroup(siteGroupName)) {
                        shouldInclude = false;
                    }
                }
                
                // Check in grantedToIdentitiesV2 for groups (older API format)
                if (Array.isArray(permission.grantedToIdentitiesV2)) {
                    const hasDefaultGroup = permission.grantedToIdentitiesV2.some(g => {
                        if (g.group) {
                            const groupName = g.group.displayName || g.group.email || '';
                            return configModule.isDefaultSharePointGroup(groupName);
                        }
                        return false;
                    });
                    
                    if (hasDefaultGroup) {
                        shouldInclude = false;
                    }
                }
            }
            
            if (shouldInclude) {
                const who = configModule.extractUserFromPermission(permission, configModule.tenantDomains);
                const roles = (permission.roles || []).join('; ') || 'Not specified'; // Use semicolon to avoid CSV issues
                const expiration = configModule.extractExpirationDate(permission);
                const classification = configModule.classifyPermission(permission, configModule.tenantDomains);
                
                // Escape CSV values and handle commas/quotes
                const csvRow = [
                    timestamp,
                    result.scanType === 'onedrive' ? 'OneDrive' : 'SharePoint',
                    escapeCsvValue(result.siteName || 'OneDrive'),
                    escapeCsvValue(result.siteUrl || 'Personal OneDrive'),
                    escapeCsvValue(result.itemId),
                    escapeCsvValue(result.itemName),
                    escapeCsvValue(result.itemPath),
                    result.itemType || 'folder',
                    escapeCsvValue(owners),
                    escapeCsvValue(who),
                    escapeCsvValue(roles),
                    classification.toUpperCase(),
                    escapeCsvValue(expiration)
                ].join(',') + '\n';
                
                await csvWriterStream.write(csvRow);
                csvRowCount++;
            }
        }
        
        // Update status periodically
        if (csvRowCount % 10 === 0) {
            updateCsvLocationStatus(`${csvRowCount} rows exported`);
        }
        
    } catch (error) {
        console.error('Error writing to CSV:', error);
        configModule.showToast('Error writing to CSV file: ' + error.message);
        await finalizeCsvExport(); // Close the file on error
    }
}

// Escape CSV values to handle commas, quotes, and newlines
function escapeCsvValue(value) {
    if (value === null || value === undefined) {
        return '';
    }
    
    const stringValue = String(value);
    
    // If the value contains comma, quote, or newline, wrap it in quotes and escape internal quotes
    if (stringValue.includes(',') || stringValue.includes('"') || stringValue.includes('\n')) {
        return '"' + stringValue.replace(/"/g, '""') + '"';
    }
    
    return stringValue;
}

// Finalize and close the CSV file
async function finalizeCsvExport() {
    const configModule = window.configModule;
    
    if (csvWriterStream) {
        try {
            await csvWriterStream.close();
            console.log('📝 Real-time CSV export finalized:', csvRowCount, 'rows written');
            if (configModule && configModule.showToast) {
                configModule.showToast(`CSV export completed: ${csvRowCount} rows written to ${csvFileHandle ? csvFileHandle.name : 'file'}`);
            }
        } catch (error) {
            console.error('Error finalizing CSV export:', error);
        }
    }
    
    // Reset state
    csvFileHandle = null;
    csvWriterStream = null;
    csvRowCount = 0;
    realtimeCsvEnabled = false;
    
    // Update UI
    const checkbox = document.getElementById('enable-realtime-csv');
    if (checkbox) {
        checkbox.checked = false;
    }
    
    const chooseButton = document.getElementById('choose-csv-location');
    if (chooseButton) {
        chooseButton.disabled = true;
    }
    
    updateCsvLocationStatus('Export completed');
    
    // Reset the status after a delay
    setTimeout(() => {
        updateCsvLocationStatus('No location selected');
    }, 5000);
}

// Update CSV location status display
function updateCsvLocationStatus(status) {
    const statusElement = document.getElementById('csv-location-status');
    if (statusElement) {
        statusElement.textContent = status;
    }
}

// Handle CSV location selection
async function handleCsvLocationSelection() {
    const configModule = window.configModule;
    
    try {
        const success = await initializeCsvExport();
        
        if (success) {
            realtimeCsvEnabled = true;
            updateCsvLocationStatus(`Ready to export to: ${csvFileHandle.name}`);
            
            if (configModule && configModule.showToast) {
                configModule.showToast('Real-time CSV export enabled');
            }
        } else {
            // Disable the checkbox if initialization failed
            const checkbox = document.getElementById('enable-realtime-csv');
            if (checkbox) {
                checkbox.checked = false;
            }
        }
    } catch (error) {
        console.error('Error selecting CSV location:', error);
        if (configModule && configModule.showToast) {
            configModule.showToast('Failed to setup CSV export');
        }
        
        // Disable the checkbox on error
        const checkbox = document.getElementById('enable-realtime-csv');
        if (checkbox) {
            checkbox.checked = false;
        }
    }
}

// Initialize real-time CSV export controls
function initializeRealtimeCsvExport() {
    const enableCheckbox = document.getElementById('enable-realtime-csv');
    const chooseButton = document.getElementById('choose-csv-location');
    
    if (enableCheckbox) {
        enableCheckbox.addEventListener('change', function() {
            const isEnabled = this.checked;
            
            if (chooseButton) {
                chooseButton.disabled = !isEnabled;
            }
            
            if (!isEnabled && realtimeCsvEnabled) {
                // User unchecked while export was active - finalize the export
                finalizeCsvExport();
            }
            
            if (isEnabled && !realtimeCsvEnabled) {
                updateCsvLocationStatus('Click "Choose Location" to select file');
            } else if (!isEnabled) {
                updateCsvLocationStatus('No location selected');
            }
        });
    }
    
    if (chooseButton) {
        chooseButton.addEventListener('click', handleCsvLocationSelection);
    }
    
    console.log('✅ Real-time CSV export controls initialized');
}

// Enhanced export module initialization to include real-time CSV
function initializeExportModuleEnhanced() {
    console.log('📊 Initializing enhanced export functionality...');
    
    try {
        initializeExportHandlers();
        initializeBulkOperationsHandlers();
        initializeRealtimeCsvExport();
        
        console.log('✅ Enhanced export module initialized successfully');
    } catch (error) {
        console.error('❌ Error initializing enhanced export module:', error);
    }
}

// Export functions for use in other modules
window.exportModule = {
    // Main export functions
    exportResults,
    exportResultsWithStatistics,
    
    // Bulk operations
    downloadCSVTemplate,
    handleCSVUpload,
    processBulkOperations,
    
    // Validation
    validateBulkCSVFormat,
    
    // Statistics
    generateExportStatistics,
    
    // Real-time CSV export
    initializeCsvExport,
    writeResultToCsv,
    finalizeCsvExport,
    handleCsvLocationSelection,
    initializeRealtimeCsvExport,
    
    // State getters
    get realtimeCsvEnabled() { return realtimeCsvEnabled; },
    get csvRowCount() { return csvRowCount; },
    
    // Initialization
    initializeExportModule: initializeExportModuleEnhanced,
    initializeExportHandlers,
    initializeBulkOperationsHandlers
};
