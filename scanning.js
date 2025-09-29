// scanning.js - Scanning Module for SharePoint & OneDrive Scanner v3.0
// Handles SharePoint and OneDrive scanning logic with enhanced filtering and real-time results

// ENHANCED SHAREPOINT SCANNING WITH NEW FEATURES
async function scanSharePointSites() {
    const configModule = window.configModule;
    const apiModule = window.apiModule;
    const uiModule = window.uiModule;
    const authModule = window.authModule;
    
    if (!configModule || !apiModule) {
        console.error('Required modules not available');
        return;
    }
    
    if (configModule.scanning) return;
    
    configModule.scanning = true;
    configModule.resetScanController();
    
    // Start enhanced token monitoring for scanning operations
    if (authModule && authModule.startScanningTokenMonitoring) {
        authModule.startScanningTokenMonitoring();
    }
    
    const progressSection = document.getElementById('sharepoint-progress-section');
    const progressBar = document.getElementById('sharepoint-progress-bar');
    const progressText = document.getElementById('sharepoint-progress-text');
    const scanBtn = document.getElementById('scan-sharepoint-btn');
    const stopBtn = document.getElementById('stop-sharepoint-btn');
    
    configModule.showProgressSection('sharepoint-progress-section');
    configModule.updateProgressBar('sharepoint-progress-bar', 0);
    configModule.updateProgressText('sharepoint-progress-text', 'Initializing SharePoint scan with enhanced features...');
    
    if (uiModule) {
        uiModule.updateButtonStates(true);
    }
    
    try {
        console.log('ENHANCED SHAREPOINT SCAN STARTING');
        console.log('Settings:', configModule.scanSettings);
        
        configModule.updateProgressText('sharepoint-progress-text', 'Loading tenant domains...');
        configModule.tenantDomains = await apiModule.loadTenantDomains();
        await apiModule.delay(100);
        
        if (configModule.controller.stop) {
            configModule.updateProgressText('sharepoint-progress-text', 'SharePoint scan stopped');
            return;
        }
        
        const selectedSites = configModule.sites.filter(s => configModule.selectedSiteIds.has(s.id));
        const filterText = configModule.scanSettings.sharingFilter === 'external' ? 'external sharing' : 
                          configModule.scanSettings.sharingFilter === 'internal' ? 'internal sharing' : 'all sharing';
        const scopeText = configModule.scanSettings.contentScope === 'folders' ? 'folders' : 'all content';
        
        configModule.showToast(`Starting enhanced SharePoint scan of ${selectedSites.length} sites (${filterText}, ${scopeText})...`);
        
        let currentSiteIndex = 0;
        let totalDrives = 0;
        let currentDrive = 0;
        
        for (const site of selectedSites) {
            if (configModule.controller.stop) break;
            
            try {
                currentSiteIndex++;
                
                const siteProgress = (currentSiteIndex / selectedSites.length) * 20;
                configModule.updateProgressBar('sharepoint-progress-bar', siteProgress);
                configModule.updateProgressText('sharepoint-progress-text', `ANALYZING SITE ${currentSiteIndex}/${selectedSites.length}: ${site.name}...`);
                
                console.log(`PROCESSING SITE ${currentSiteIndex}/${selectedSites.length}: ${site.name}`);
                
                const allDrives = await apiModule.getSiteDrives(site.id);
                // Filter out preservation hold libraries before processing
                const drives = allDrives.filter(drive => {
                    const driveName = drive.name || 'Documents';
                    const isPreservationHold = configModule.shouldSkipPreservationHoldLibrary(driveName);
                    if (isPreservationHold) {
                        console.log(`🚫 SKIPPING PRESERVATION HOLD DRIVE: ${driveName} in site ${site.name}`);
                    }
                    return !isPreservationHold;
                });
                
                totalDrives += drives.length;
                console.log(`Found ${allDrives.length} total drives (${drives.length} after filtering preservation holds) in ${site.name}`);
                
                for (const drive of drives) {
                    if (configModule.controller.stop) break;
                    
                    currentDrive++;
                    
                    const driveProgress = 20 + ((currentDrive / Math.max(totalDrives, 1)) * 80);
                    configModule.updateProgressBar('sharepoint-progress-bar', driveProgress);
                    
                    configModule.updateProgressText('sharepoint-progress-text', `SCANNING DRIVE ${currentDrive}/${totalDrives}: ${site.name}/${drive.name || 'Documents'}...`);
                    
                    await scanDriveWithDelta(site, drive, 'sharepoint-progress-text', 'sharepoint');
                }
                
            } catch (e) {
                console.warn(`Error scanning site ${site.name}:`, e);
            }
        }
        
        if (!configModule.controller.stop) {
            configModule.updateProgressBar('sharepoint-progress-bar', 100);
            const sharePointResults = configModule.results.filter(r => r.scanType === 'sharepoint').length;
            configModule.updateProgressText('sharepoint-progress-text', `SHAREPOINT SCAN COMPLETED • ${sharePointResults} items with ${filterText} found`);
            configModule.showToast(`SharePoint scan completed! Found ${sharePointResults} items with ${filterText}.`);
        } else {
            configModule.updateProgressText('sharepoint-progress-text', 'SharePoint scan stopped by user');
        }
        
        // Finalize CSV export if it was active during scanning
        const exportModule = window.exportModule;
        if (exportModule && exportModule.realtimeCsvEnabled) {
            try {
                await exportModule.finalizeCsvExport();
                console.log('📝 CSV: SharePoint scan CSV export finalized');
            } catch (csvError) {
                console.warn('Failed to finalize CSV export:', csvError);
            }
        }
        
    } catch (error) {
        console.error('SharePoint scan error:', error);
        alert('SharePoint scan error: ' + error.message);
        configModule.updateProgressText('sharepoint-progress-text', 'SharePoint scan failed - check console for details');
        
        // Finalize CSV export on error
        const exportModule = window.exportModule;
        if (exportModule && exportModule.realtimeCsvEnabled) {
            try {
                await exportModule.finalizeCsvExport();
            } catch (csvError) {
                console.warn('Failed to finalize CSV export on error:', csvError);
            }
        }
    } finally {
        configModule.scanning = false;
        
        // Stop enhanced token monitoring for scanning operations
        if (authModule && authModule.stopScanningTokenMonitoring) {
            authModule.stopScanningTokenMonitoring();
        }
        
        if (uiModule) {
            uiModule.updateButtonStates(false);
        }
    }
}

// ENHANCED ONEDRIVE SCANNING WITH NEW FEATURES
async function scanOneDriveUsers() {
    const configModule = window.configModule;
    const apiModule = window.apiModule;
    const uiModule = window.uiModule;
    const authModule = window.authModule;
    
    if (!configModule || !apiModule) {
        console.error('Required modules not available');
        return;
    }
    
    if (configModule.scanning) return;
    
    configModule.scanning = true;
    configModule.resetScanController();
    
    // Start enhanced token monitoring for scanning operations
    if (authModule && authModule.startScanningTokenMonitoring) {
        authModule.startScanningTokenMonitoring();
    }
    
    configModule.showProgressSection('onedrive-progress-section');
    configModule.updateProgressBar('onedrive-progress-bar', 0);
    configModule.updateProgressText('onedrive-progress-text', 'Initializing OneDrive user scan...');
    
    if (uiModule) {
        uiModule.updateButtonStates(true);
    }
    
    try {
        console.log('ENHANCED ONEDRIVE USER SCAN STARTING');
        console.log('Settings:', configModule.scanSettings);
        
        configModule.updateProgressText('onedrive-progress-text', 'Loading tenant domains...');
        configModule.tenantDomains = await apiModule.loadTenantDomains();
        configModule.updateProgressBar('onedrive-progress-bar', 10);
        await apiModule.delay(100);
        
        if (configModule.controller.stop) {
            configModule.updateProgressText('onedrive-progress-text', 'OneDrive scan stopped');
            return;
        }
        
        const selectedUsers = configModule.users.filter(u => configModule.selectedUserIds.has(u.id));
        if (selectedUsers.length === 0) {
            configModule.updateProgressText('onedrive-progress-text', 'No users selected for scanning');
            configModule.showToast('Please select users to scan');
            return;
        }
        
        const filterText = configModule.scanSettings.sharingFilter === 'external' ? 'external sharing' : 
                          configModule.scanSettings.sharingFilter === 'internal' ? 'internal sharing' : 'all sharing';
        const scopeText = configModule.scanSettings.contentScope === 'folders' ? 'folders' : 'all content';
        
        configModule.showToast(`Starting OneDrive scan for ${selectedUsers.length} users (${filterText}, ${scopeText})...`);
        
        let currentUserIndex = 0;
        for (const user of selectedUsers) {
            if (configModule.controller.stop) break;
            
            currentUserIndex++;
            
            // Periodic token validation during long scans
            if (authModule && authModule.ensureValidTokenForScanning && currentUserIndex % 3 === 0) {
                try {
                    await authModule.ensureValidTokenForScanning();
                } catch (tokenError) {
                    console.error('❌ Token validation failed during scan:', tokenError);
                    throw new Error('Authentication token expired during scan - please sign in again');
                }
            }
            
            const userProgress = 10 + ((currentUserIndex / selectedUsers.length) * 90);
            configModule.updateProgressBar('onedrive-progress-bar', userProgress);
            
            configModule.updateProgressText('onedrive-progress-text', `SCANNING USER ${currentUserIndex}/${selectedUsers.length}: ${user.displayName || user.userPrincipalName}...`);
            
            try {
                const drive = await apiModule.getUserOneDrive(user.id);
                
                const oneDriveSite = {
                    name: `${user.displayName || user.userPrincipalName} OneDrive`,
                    id: `onedrive-${user.id}`,
                    webUrl: drive.webUrl || 'https://onedrive.live.com'
                };
                
                await scanDriveWithDelta(oneDriveSite, drive, 'onedrive-progress-text', 'onedrive');
                
            } catch (error) {
                console.warn(`Failed to scan OneDrive for user ${user.displayName || user.userPrincipalName}:`, error);
                if (error.message.includes('404') || error.message.includes('mysite not found')) {
                    console.log(`User ${user.displayName} does not have OneDrive provisioned`);
                }
                // If it's an authentication error, re-throw it to stop the scan
                if (error.message && error.message.includes('Authentication token expired')) {
                    throw error;
                }
            }
        }
        
        if (!configModule.controller.stop) {
            configModule.updateProgressBar('onedrive-progress-bar', 100);
            const oneDriveResults = configModule.results.filter(r => r.scanType === 'onedrive').length;
            configModule.updateProgressText('onedrive-progress-text', `ONEDRIVE SCAN COMPLETED • ${oneDriveResults} items with ${filterText} found`);
            configModule.showToast(`OneDrive scan completed! Found ${oneDriveResults} items with ${filterText}.`);
        } else {
            configModule.updateProgressText('onedrive-progress-text', 'OneDrive scan stopped by user');
        }
        
        // Finalize CSV export if it was active during scanning
        const exportModule = window.exportModule;
        if (exportModule && exportModule.realtimeCsvEnabled) {
            try {
                await exportModule.finalizeCsvExport();
                console.log('📝 CSV: OneDrive scan CSV export finalized');
            } catch (csvError) {
                console.warn('Failed to finalize CSV export:', csvError);
            }
        }
        
    } catch (error) {
        console.error('OneDrive scan error:', error);
        alert('OneDrive scan error: ' + error.message);
        configModule.updateProgressText('onedrive-progress-text', 'OneDrive scan failed - check console for details');
        
        // Finalize CSV export on error
        const exportModule = window.exportModule;
        if (exportModule && exportModule.realtimeCsvEnabled) {
            try {
                await exportModule.finalizeCsvExport();
            } catch (csvError) {
                console.warn('Failed to finalize CSV export on error:', csvError);
            }
        }
    } finally {
        configModule.scanning = false;
        
        // Stop enhanced token monitoring for scanning operations
        if (authModule && authModule.stopScanningTokenMonitoring) {
            authModule.stopScanningTokenMonitoring();
        }
        
        if (uiModule) {
            uiModule.updateButtonStates(false);
        }
    }
}

// ENHANCED DELTA SCANNING WITH NEW FILTERING
async function scanDriveWithDelta(site, drive, progressTextId, scanType = 'sharepoint') {
    const configModule = window.configModule;
    const apiModule = window.apiModule;
    
    if (!configModule || !apiModule) {
        console.error('Required modules not available');
        return;
    }
    
    try {
        const sourceName = scanType === 'onedrive' ? 'OneDrive' : `${site.name}/${drive.name || 'Documents'}`;
        console.log(`STARTING ENHANCED DELTA QUERY for: ${sourceName}`);
        
        if (progressTextId) {
            configModule.updateProgressText(progressTextId, `DELTA SCANNING: ${sourceName} (enhanced filtering)...`);
        }
        
        const allItems = await apiModule.performDeltaQuery(drive.id);
        
        if (progressTextId) {
            configModule.updateProgressText(progressTextId, `DELTA SCANNING ${sourceName}: ${allItems.length} items processed...`);
        }
        
        await processEnhancedDeltaItems(site, drive, allItems, scanType);
        
        if (progressTextId) {
            const currentResults = configModule.results.filter(r => r.scanType === scanType).length;
            configModule.updateProgressText(progressTextId, `DELTA COMPLETED for ${sourceName}: ${allItems.length} items • ${currentResults} shared items found`);
        }
        
        console.log(`ENHANCED DELTA SCAN COMPLETED for ${sourceName}: ${allItems.length} total items processed`);
        
    } catch (error) {
        console.warn(`ENHANCED DELTA SCAN FAILED for ${drive.name || 'OneDrive'}:`, error);
        
        if (progressTextId) {
            configModule.updateProgressText(progressTextId, `Delta failed for ${drive.name || 'OneDrive'}, switching to COMPREHENSIVE MODE...`);
        }
        await apiModule.delay(300); // Reduced delay when switching to comprehensive mode
        
        console.log('FALLING BACK to enhanced comprehensive folder traversal...');
        await scanDriveComprehensive(site, drive, progressTextId, scanType);
    }
}

// ENHANCED DELTA ITEM PROCESSING WITH NEW FILTERING
async function processEnhancedDeltaItems(site, drive, items, scanType) {
    const configModule = window.configModule;
    const apiModule = window.apiModule;
    const resultsModule = window.resultsModule;
    
    if (!configModule || !apiModule) {
        console.error('Required modules not available');
        return;
    }
    
    for (const item of items) {
        if (configModule.controller.stop) return;
        
        // Skip preservation hold libraries in delta scanning too
        if (item.folder && configModule.shouldSkipPreservationHoldLibrary(item.name)) {
            console.log(`🚫 DELTA SKIPPING PRESERVATION HOLD: ${item.name}`);
            continue;
        }
        
        // Skip items that don't match content scope
        if (configModule.scanSettings.contentScope === 'folders' && !item.folder) {
            continue;
        }
        
        if (!item.permissions || item.permissions.length === 0) continue;
        
        // 🔥 CRITICAL FIX: Store BOTH all permissions AND filtered permissions for owner detection
        // Apply enhanced filtering based on scan settings for display
        const interesting = item.permissions.filter(p => 
            configModule.shouldIncludePermission(p, configModule.tenantDomains, configModule.scanSettings.sharingFilter)
        );

        if (interesting.length > 0) {
            let itemPath = '';
            let displayLocation = '';
            
            if (scanType === 'onedrive') {
                if (item.parentReference?.path) {
                    let parentPath = item.parentReference.path;
                    parentPath = parentPath.replace('/drive/root:', '');
                    parentPath = parentPath.replace(/^\/drives\/[^\/]+/, '');
                    itemPath = parentPath ? `${parentPath}/${item.name}` : `/${item.name}`;
                } else {
                    itemPath = `/${item.name}`;
                }
                displayLocation = 'OneDrive';
            } else {
                const driveName = drive.name || 'Documents';
                if (item.parentReference?.path) {
                    let parentPath = item.parentReference.path;
                    parentPath = parentPath.replace('/drive/root:', '');
                    parentPath = parentPath.replace(/^\/drives\/[^\/]+/, '');
                    if (parentPath && parentPath !== '/') {
                        itemPath = `/${driveName}${parentPath}/${item.name}`;
                    } else {
                        itemPath = `/${driveName}/${item.name}`;
                    }
                } else {
                    itemPath = `/${driveName}/${item.name}`;
                }
                displayLocation = site.name;
            }
            
            itemPath = itemPath.replace(/\/+/g, '/');
            if (!itemPath.startsWith('/')) itemPath = '/' + itemPath;

            const scanResult = {
                siteName: displayLocation,
                siteUrl: site.webUrl,
                driveId: drive.id,
                itemId: item.id,
                itemName: item.name,
                itemPath: itemPath,
                itemType: item.folder ? 'folder' : 'file',
                permissions: interesting, // Filtered permissions for display
                allPermissions: item.permissions, // Complete permissions set for owner detection
                scanType: scanType,
                driveName: drive.name || (scanType === 'onedrive' ? 'OneDrive' : 'Documents')
            };
            
            configModule.results.push(scanResult);
            console.log(`REAL-TIME: Adding result #${configModule.results.length} to display: ${scanResult.itemName}`);
            
            // Write to real-time CSV if enabled
            const exportModule = window.exportModule;
            if (exportModule && exportModule.realtimeCsvEnabled) {
                try {
                    await exportModule.writeResultToCsv(scanResult);
                    console.log(`📝 CSV: Exported result to CSV: ${scanResult.itemName}`);
                } catch (csvError) {
                    console.warn('Failed to write result to CSV:', csvError);
                }
            }
            
            // Force immediate DOM update
            if (resultsModule && resultsModule.updateResultsDisplay) {
                resultsModule.updateResultsDisplay();
                resultsModule.addResultToDisplay(scanResult);
            }
            
            // Force browser to render the update
            await new Promise(resolve => setTimeout(resolve, 10));

            console.log(`ENHANCED DELTA FOUND shared ${scanResult.itemType}: ${itemPath} (${interesting.length} permissions, filter: ${configModule.scanSettings.sharingFilter})`);
        }
    }
}

// ENHANCED COMPREHENSIVE SCANNING WITH NEW FILTERING
async function scanDriveComprehensive(site, drive, progressTextId, scanType) {
    const configModule = window.configModule;
    const apiModule = window.apiModule;
    
    if (!configModule || !apiModule) {
        console.error('Required modules not available');
        return;
    }
    
    try {
        const sourceName = scanType === 'onedrive' ? 'OneDrive' : `${site.name}/${drive.name || 'Documents'}`;
        console.log(`ENHANCED COMPREHENSIVE SCAN starting for ${sourceName}`);
        
        if (progressTextId) {
            configModule.updateProgressText(progressTextId, `COMPREHENSIVE MODE: ${sourceName} (enhanced filtering)...`);
        }
        
        const scanState = { scannedFolders: 0, totalBatches: 0, foundItems: 0 };
        const suppressedPaths = new Set();
        
        await traverseFolderEnhanced(site, drive, "root", "", suppressedPaths, scanState, scanType, progressTextId);
        
        if (progressTextId) {
            configModule.updateProgressText(progressTextId, `COMPREHENSIVE COMPLETED for ${sourceName}: ${scanState.scannedFolders} folders, ${scanState.totalBatches} batches, ${scanState.foundItems} found`);
        }
        
        console.log(`ENHANCED COMPREHENSIVE SCAN COMPLETED for ${sourceName}: ${scanState.scannedFolders} folders, ${scanState.totalBatches} batches`);
        
    } catch (error) {
        console.warn(`ENHANCED COMPREHENSIVE SCAN FAILED for ${drive.name || 'OneDrive'}:`, error);
        if (progressTextId) {
            configModule.updateProgressText(progressTextId, `Enhanced comprehensive scan failed for ${drive.name || 'OneDrive'}`);
        }
    }
}

// ENHANCED FOLDER TRAVERSAL WITH NEW FEATURES
async function traverseFolderEnhanced(site, drive, itemId, path, suppressedPaths, scanState, scanType, progressTextId) {
    const configModule = window.configModule;
    const apiModule = window.apiModule;
    const resultsModule = window.resultsModule;
    
    if (!configModule || !apiModule) {
        console.error('Required modules not available');
        return;
    }
    
    if (configModule.controller.stop) return;
    
    const sourceName = scanType === 'onedrive' ? 'OneDrive' : `${site.name}/${drive.name || 'Documents'}`;
    
    // Get children based on content scope
    const includeFiles = configModule.scanSettings.contentScope === 'all';
    const children = await apiModule.getFolderChildren(drive.id, itemId, includeFiles);

        const validItems = children.filter(f => {
            // Skip preservation hold libraries - CRITICAL FIX: Apply at folder level too
            if (f.folder && configModule.shouldSkipPreservationHoldLibrary(f.name)) {
                console.log(`🚫 SKIPPING PRESERVATION HOLD FOLDER: ${f.name} in ${sourceName}`);
                return false;
            }
            
            // Skip system folders
            if (f.folder && configModule.shouldSkipFolder(f.name)) {
                console.log(`🚫 SKIPPING SYSTEM FOLDER: ${f.name} in ${sourceName}`);
                return false;
            }
            
            if (configModule.scanSettings.contentScope === 'folders') {
                return f.folder;
            } else {
                // For all content, include files and folders (already filtered above)
                return f.file || f.folder;
            }
        });

    if (validItems.length === 0) return;

    const itemsToCheck = [];
    
    for (const f of validItems) {
        if (configModule.controller.stop) return;
        
        let itemPath = configModule.formatItemPath(f.parentReference?.path, f.name, drive.name, scanType);

        let suppressed = false;
        for (const sp of suppressedPaths) {
            if (!sp) continue;
            if (itemPath.toLowerCase().startsWith(sp.toLowerCase())) { 
                suppressed = true; 
                break; 
            }
        }

        if (!suppressed) {
            itemsToCheck.push({
                item: f,
                itemPath: itemPath,
                url: `https://graph.microsoft.com/v1.0/drives/${drive.id}/items/${f.id}/permissions`
            });
        }
    }

    if (itemsToCheck.length === 0) return;
    
    scanState.totalBatches++;
    
    if (progressTextId) {
        configModule.updateProgressText(progressTextId, `BATCH ${scanState.totalBatches}: Checking ${itemsToCheck.length} items in ${sourceName}...`);
        await apiModule.delay(20);
    }
    
    const permissionResults = await apiModule.batchGetPermissions(itemsToCheck, configModule.controller);
    const recursionTasks = [];
    
    for (const result of permissionResults) {
        if (configModule.controller.stop) return;
        
        scanState.scannedFolders++;
        
        if (scanState.scannedFolders % 5 === 0 && progressTextId) {
            configModule.updateProgressText(progressTextId, `SCANNING ${sourceName}: ${scanState.scannedFolders} items • ${scanState.totalBatches} batches • ${scanState.foundItems} found`);
            await apiModule.delay(5);
        }

        // 🔥 CRITICAL FIX: Store BOTH all permissions AND filtered permissions for owner detection  
        // Apply enhanced filtering
        const interesting = result.permissions.filter(p => 
            configModule.shouldIncludePermission(p, configModule.tenantDomains, configModule.scanSettings.sharingFilter)
        );

        if (interesting.length > 0) {
            scanState.foundItems++;
            
            let itemPath = configModule.formatItemPath(result.item.parentReference?.path, result.item.name, drive.name, scanType);

            const scanResult = {
                siteName: scanType === 'onedrive' ? 'OneDrive' : site.name,
                siteUrl: site.webUrl,
                driveId: drive.id,
                itemId: result.item.id,
                itemName: result.item.name,
                itemPath: itemPath,
                itemType: result.item.folder ? 'folder' : 'file',
                permissions: interesting, // Filtered permissions for display
                allPermissions: result.permissions, // Complete permissions set for owner detection
                scanType: scanType,
                driveName: drive.name || (scanType === 'onedrive' ? 'OneDrive' : 'Documents')
            };
            
            configModule.results.push(scanResult);
            suppressedPaths.add(result.itemPath);
            
            // Write to real-time CSV if enabled
            const exportModule = window.exportModule;
            if (exportModule && exportModule.realtimeCsvEnabled) {
                try {
                    await exportModule.writeResultToCsv(scanResult);
                    console.log(`📝 CSV: Exported result to CSV: ${scanResult.itemName}`);
                } catch (csvError) {
                    console.warn('Failed to write result to CSV:', csvError);
                }
            }
            
            if (resultsModule) {
                resultsModule.updateResultsDisplay();
                resultsModule.addResultToDisplay(scanResult);
            }
            
            console.log(`ENHANCED COMPREHENSIVE FOUND shared ${scanResult.itemType}: ${itemPath} (${interesting.length} permissions, filter: ${configModule.scanSettings.sharingFilter})`);
            
            if (progressTextId) {
                configModule.updateProgressText(progressTextId, `FOUND shared ${scanResult.itemType} in ${sourceName}: ${itemPath} (${scanState.foundItems} total)`);
                await apiModule.delay(300);
            }
        } else if (result.item.folder) {
            // Only add folders for recursion
            recursionTasks.push(
                traverseFolderEnhanced(site, drive, result.item.id, result.itemPath, suppressedPaths, scanState, scanType, progressTextId)
            );
        }
    }

    // Optimized recursion batching: increased from 1 to 3 for better throughput
    const recursionBatchSize = 3;
    for (let i = 0; i < recursionTasks.length; i += recursionBatchSize) {
        if (configModule.controller.stop) return;
        
        const batch = recursionTasks.slice(i, i + recursionBatchSize);
        
        if (progressTextId && recursionTasks.length > 1) {
            configModule.updateProgressText(progressTextId, `DEEP SCANNING ${sourceName}: ${scanState.scannedFolders} items • ${scanState.foundItems} found`);
        }
        
        await Promise.all(batch);
        await apiModule.delay(300);
    }
}

// Export functions for use in other modules
window.scanningModule = {
    // Main scanning functions
    scanSharePointSites,
    scanOneDriveUsers,
    
    // Delta scanning
    scanDriveWithDelta,
    processEnhancedDeltaItems,
    
    // Comprehensive scanning
    scanDriveComprehensive,
    traverseFolderEnhanced
};
