// permissions.js - Permissions Module for SharePoint & OneDrive Scanner v3.0
// Handles permission management, action buttons, and API operations for permission modifications

// ENHANCED ACTION BUTTON FUNCTIONS WITH REAL API CALLS
async function showExpirationDialog(result, resultDiv) {
    const configModule = window.configModule;
    const apiModule = window.apiModule;
    const resultsModule = window.resultsModule;
    
    if (!configModule || !apiModule) {
        console.error('Required modules not available for expiration dialog');
        return;
    }
    
    const expirationDate = prompt(
        `Set expiration date for: ${result.itemName}\n\nEnter date (YYYY-MM-DD format) or leave blank to remove expiration:`,
        ''
    );
    
    if (expirationDate === null) return; // User cancelled
    
    const btn = resultDiv.querySelector('.action-btn-purple');
    if (!btn) return;
    
    const originalText = btn.innerText;
    btn.disabled = true;
    btn.innerText = 'Setting...';
    
    try {
        const success = await setExpirationForItem(result, expirationDate);
        if (success) {
            configModule.showToast(`✅ Updated expiration for ${result.itemName}`);
            
            // Update the result in memory
            const resultIndex = configModule.results.findIndex(r => r.itemId === result.itemId);
            if (resultIndex !== -1 && resultsModule) {
                // Refresh the permissions for this item
                await resultsModule.refreshItemPermissions(result, resultIndex);
                // Re-render this specific result
                resultDiv.style.animation = 'fadeIn 0.3s ease-in';
            }
        } else {
            configModule.showToast(`❌ Failed to update expiration for ${result.itemName}`, 4000);
        }
    } catch (error) {
        console.error('Expiration update error:', error);
        configModule.showToast(`❌ Error updating expiration: ${error.message}`, 4000);
    } finally {
        btn.disabled = false;
        btn.innerText = originalText;
    }
}

async function disableLinks(result, resultDiv) {
    const configModule = window.configModule;
    const resultsModule = window.resultsModule;
    
    if (!configModule) {
        console.error('Required modules not available for disabling links');
        return;
    }
    
    const btn = resultDiv.querySelector('.action-btn-orange');
    if (!btn) return;
    
    const originalText = btn.innerText;
    btn.disabled = true;
    btn.innerText = 'Disabling...';
    
    try {
        const success = await removeSharingLinksForItem(result);
        if (success) {
            configModule.showToast(`✅ Disabled sharing links for ${result.itemName}`);
            
            // Update the result in memory and UI
            const resultIndex = configModule.results.findIndex(r => r.itemId === result.itemId);
            if (resultIndex !== -1 && resultsModule) {
                await resultsModule.refreshItemPermissions(result, resultIndex);
                // If no permissions remain, remove from results
                if (configModule.results[resultIndex].permissions.length === 0) {
                    configModule.results.splice(resultIndex, 1);
                    resultDiv.style.animation = 'fadeOut 0.3s ease-out';
                    setTimeout(() => resultDiv.remove(), 300);
                    if (resultsModule.updateResultsDisplay) {
                        resultsModule.updateResultsDisplay();
                    }
                } else {
                    resultDiv.style.animation = 'fadeIn 0.3s ease-in';
                }
            }
        } else {
            configModule.showToast(`❌ Failed to disable links for ${result.itemName}`, 4000);
        }
    } catch (error) {
        console.error('Disable links error:', error);
        configModule.showToast(`❌ Error disabling links: ${error.message}`, 4000);
    } finally {
        btn.disabled = false;
        btn.innerText = originalText;
    }
}

async function disableAllSharing(result, resultDiv) {
    const configModule = window.configModule;
    const resultsModule = window.resultsModule;
    
    if (!configModule) {
        console.error('Required modules not available for disabling all sharing');
        return;
    }
    
    const btn = resultDiv.querySelector('.action-btn-red');
    if (!btn) return;
    
    const originalText = btn.innerText;
    btn.disabled = true;
    btn.innerText = 'Removing...';
    
    try {
        const success = await removeAllSharingForItem(result);
        if (success) {
            configModule.showToast(`✅ Removed all sharing for ${result.itemName}`);
            
            // Remove from results since no sharing remains
            const resultIndex = configModule.results.findIndex(r => r.itemId === result.itemId);
            if (resultIndex !== -1) {
                configModule.results.splice(resultIndex, 1);
                if (resultsModule && resultsModule.updateResultsDisplay) {
                    resultsModule.updateResultsDisplay();
                }
            }
            
            // Animate removal
            resultDiv.style.animation = 'fadeOut 0.3s ease-out';
            setTimeout(() => resultDiv.remove(), 300);
        } else {
            configModule.showToast(`❌ Failed to remove sharing for ${result.itemName}`, 4000);
        }
    } catch (error) {
        console.error('Remove all sharing error:', error);
        configModule.showToast(`❌ Error removing sharing: ${error.message}`, 4000);
    } finally {
        btn.disabled = false;
        btn.innerText = originalText;
    }
}

// SUPPORTING FUNCTIONS FOR PERMISSION MANAGEMENT
async function setExpirationForItem(result, expirationDate) {
    const apiModule = window.apiModule;
    
    if (!apiModule) {
        console.error('API module not available for setting expiration');
        return false;
    }
    
    try {
        const permissionsUrl = `https://graph.microsoft.com/v1.0/drives/${result.driveId}/items/${result.itemId}/permissions`;
        
        // Get current permissions
        const currentPermissions = await apiModule.requestQueue.add(async () => {
            return await apiModule.graphGetAll(permissionsUrl);
        });
        
        let success = true;
        
        // Update each link permission with expiration
        for (const permission of currentPermissions) {
            if (permission.link) {
                try {
                    const updateBody = {};
                    
                    if (expirationDate && expirationDate.trim()) {
                        // Set expiration
                        updateBody.expirationDateTime = new Date(expirationDate + 'T23:59:59.000Z').toISOString();
                    } else {
                        // Remove expiration by setting it to null
                        updateBody.expirationDateTime = null;
                    }
                    
                    await apiModule.requestQueue.add(async () => {
                        return await apiModule.graphRequestWithRetry(`${permissionsUrl}/${permission.id}`, {
                            method: 'PATCH',
                            headers: { 'Content-Type': 'application/json' },
                            body: JSON.stringify(updateBody)
                        });
                    });
                    
                    console.log(`Updated expiration for permission ${permission.id}`);
                } catch (error) {
                    console.warn(`Failed to update permission ${permission.id}:`, error);
                    success = false;
                }
            }
        }
        
        return success;
    } catch (error) {
        console.error('Error setting expiration:', error);
        return false;
    }
}

async function removeSharingLinksForItem(result) {
    const apiModule = window.apiModule;
    
    if (!apiModule) {
        console.error('API module not available for removing sharing links');
        return false;
    }
    
    try {
        const permissionsUrl = `https://graph.microsoft.com/v1.0/drives/${result.driveId}/items/${result.itemId}/permissions`;
        
        // Get current permissions
        const currentPermissions = await apiModule.requestQueue.add(async () => {
            return await apiModule.graphGetAll(permissionsUrl);
        });
        
        let success = true;
        
        // Remove only link permissions, keep direct user permissions
        for (const permission of currentPermissions) {
            if (permission.link) {
                try {
                    await apiModule.requestQueue.add(async () => {
                        return await apiModule.graphRequestWithRetry(`${permissionsUrl}/${permission.id}`, {
                            method: 'DELETE'
                        });
                    });
                    
                    console.log(`Removed link permission ${permission.id}`);
                } catch (error) {
                    console.warn(`Failed to remove permission ${permission.id}:`, error);
                    success = false;
                }
            }
        }
        
        return success;
    } catch (error) {
        console.error('Error removing sharing links:', error);
        return false;
    }
}

async function removeAllSharingForItem(result) {
    const apiModule = window.apiModule;
    
    if (!apiModule) {
        console.error('API module not available for removing all sharing');
        return false;
    }
    
    try {
        const permissionsUrl = `https://graph.microsoft.com/v1.0/drives/${result.driveId}/items/${result.itemId}/permissions`;
        
        // Get current permissions
        const currentPermissions = await apiModule.requestQueue.add(async () => {
            return await apiModule.graphGetAll(permissionsUrl);
        });
        
        let success = true;
        
        // Remove all permissions except owner permissions
        for (const permission of currentPermissions) {
            // Skip owner permissions and system permissions
            if (permission.roles && permission.roles.includes('owner')) {
                console.log(`Skipping owner permission ${permission.id}`);
                continue;
            }
            
            if (permission.id === 'root') {
                console.log('Skipping root permission');
                continue;
            }
            
            try {
                await apiModule.requestQueue.add(async () => {
                    return await apiModule.graphRequestWithRetry(`${permissionsUrl}/${permission.id}`, {
                        method: 'DELETE'
                    });
                });
                
                console.log(`Removed permission ${permission.id}`);
            } catch (error) {
                console.warn(`Failed to remove permission ${permission.id}:`, error);
                success = false;
            }
        }
        
        return success;
    } catch (error) {
        console.error('Error removing all sharing:', error);
        return false;
    }
}

// BULK PERMISSION OPERATIONS
async function addBulkPermission(permissionsUrl, userEmail, role, linkScope, expirationDate) {
    const apiModule = window.apiModule;
    
    if (!apiModule) {
        console.error('API module not available for adding bulk permission');
        return false;
    }
    
    try {
        const body = {
            requireSignIn: true,
            sendInvitation: false,
            roles: [role || 'read']
        };
        
        if (linkScope) {
            body.link = { scope: linkScope };
        }
        
        if (userEmail) {
            body.recipients = [{ email: userEmail }];
        }
        
        if (expirationDate) {
            body.expirationDateTime = new Date(expirationDate).toISOString();
        }
        
        await apiModule.requestQueue.add(async () => {
            return await apiModule.graphRequestWithRetry(permissionsUrl, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(body)
            });
        });
        
        return true;
    } catch (error) {
        console.error('Failed to add permission:', error);
        return false;
    }
}

async function removeBulkPermission(permissionsUrl, userEmail) {
    const apiModule = window.apiModule;
    
    if (!apiModule) {
        console.error('API module not available for removing bulk permission');
        return false;
    }
    
    try {
        // Get current permissions
        const currentPermissions = await apiModule.requestQueue.add(async () => {
            return await apiModule.graphGetAll(permissionsUrl);
        });
        
        // Find permission to remove
        const permissionToRemove = currentPermissions.find(p => 
            (p.grantedTo && p.grantedTo.user && p.grantedTo.user.email === userEmail) ||
            (p.grantedToIdentitiesV2 && p.grantedToIdentitiesV2.some(g => g.user && g.user.email === userEmail))
        );
        
        if (permissionToRemove) {
            await apiModule.requestQueue.add(async () => {
                return await apiModule.graphRequestWithRetry(`${permissionsUrl}/${permissionToRemove.id}`, {
                    method: 'DELETE'
                });
            });
            return true;
        }
        
        return false;
    } catch (error) {
        console.error('Failed to remove permission:', error);
        return false;
    }
}

async function modifyBulkPermission(permissionsUrl, linkScope, expirationDate) {
    const apiModule = window.apiModule;
    
    if (!apiModule) {
        console.error('API module not available for modifying bulk permission');
        return false;
    }
    
    try {
        // Get current permissions
        const currentPermissions = await apiModule.requestQueue.add(async () => {
            return await apiModule.graphGetAll(permissionsUrl);
        });
        
        // Find link permissions to modify
        const linkPermissions = currentPermissions.filter(p => p.link);
        
        for (const permission of linkPermissions) {
            const updateBody = {};
            
            if (expirationDate) {
                updateBody.expirationDateTime = new Date(expirationDate).toISOString();
            }
            
            if (Object.keys(updateBody).length > 0) {
                await apiModule.requestQueue.add(async () => {
                    return await apiModule.graphRequestWithRetry(`${permissionsUrl}/${permission.id}`, {
                        method: 'PATCH',
                        headers: { 'Content-Type': 'application/json' },
                        body: JSON.stringify(updateBody)
                    });
                });
            }
        }
        
        return true;
    } catch (error) {
        console.error('Failed to modify permission:', error);
        return false;
    }
}

async function processSingleBulkOperation(operation) {
    const configModule = window.configModule;
    
    if (!configModule) {
        console.error('Config module not available for bulk operations');
        return false;
    }
    
    const { ItemID, Action, UserEmail, Role, LinkScope, ExpirationDate } = operation;
    
    try {
        // Find the item by ItemID across all results
        const targetResult = configModule.results.find(r => r.itemId === ItemID);
        if (!targetResult) {
            console.warn(`Item not found in results: ${ItemID}`);
            return false;
        }
        
        const permissionsUrl = `https://graph.microsoft.com/v1.0/drives/${targetResult.driveId}/items/${ItemID}/permissions`;
        
        switch (Action.toLowerCase()) {
            case 'add':
                return await addBulkPermission(permissionsUrl, UserEmail, Role, LinkScope, ExpirationDate);
            case 'remove':
                return await removeBulkPermission(permissionsUrl, UserEmail);
            case 'modify':
                return await modifyBulkPermission(permissionsUrl, LinkScope, ExpirationDate);
            default:
                console.warn(`Unknown action: ${Action}`);
                return false;
        }
    } catch (error) {
        console.error(`Error processing bulk operation for ${ItemID}:`, error);
        return false;
    }
}

// PERMISSION VALIDATION UTILITIES
function validatePermissionChange(result, action, parameters = {}) {
    if (!result || !result.itemId) {
        return { isValid: false, message: 'Invalid result object' };
    }
    
    switch (action) {
        case 'setExpiration':
            if (parameters.expirationDate) {
                const date = new Date(parameters.expirationDate);
                if (isNaN(date.getTime())) {
                    return { isValid: false, message: 'Invalid expiration date format' };
                }
                if (date < new Date()) {
                    return { isValid: false, message: 'Expiration date cannot be in the past' };
                }
            }
            break;
            
        case 'addPermission':
            if (!parameters.userEmail && !parameters.linkScope) {
                return { isValid: false, message: 'Either user email or link scope is required' };
            }
            if (parameters.userEmail && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(parameters.userEmail)) {
                return { isValid: false, message: 'Invalid email format' };
            }
            break;
            
        case 'removePermission':
            if (!parameters.userEmail && !parameters.permissionId) {
                return { isValid: false, message: 'Either user email or permission ID is required' };
            }
            break;
    }
    
    return { isValid: true, message: 'Valid permission change' };
}

// PERMISSION CHANGE LOGGING
function logPermissionChange(result, action, parameters, success, error = null) {
    const logEntry = {
        timestamp: new Date().toISOString(),
        itemName: result.itemName,
        itemId: result.itemId,
        siteName: result.siteName,
        action: action,
        parameters: parameters,
        success: success,
        error: error ? error.message : null
    };
    
    console.log('Permission Change Log:', logEntry);
    
    // Store in localStorage for audit trail (optional)
    try {
        const existingLogs = JSON.parse(localStorage.getItem('sp_scanner_permission_logs') || '[]');
        existingLogs.push(logEntry);
        
        // Keep only the last 100 log entries
        if (existingLogs.length > 100) {
            existingLogs.splice(0, existingLogs.length - 100);
        }
        
        localStorage.setItem('sp_scanner_permission_logs', JSON.stringify(existingLogs));
    } catch (storageError) {
        console.warn('Could not save permission log to localStorage:', storageError);
    }
}

// PERMISSION SUMMARY UTILITIES
function getPermissionSummary(result) {
    if (!result || !result.permissions) {
        return {
            totalPermissions: 0,
            externalCount: 0,
            internalCount: 0,
            linkCount: 0,
            directGrantCount: 0
        };
    }
    
    const configModule = window.configModule;
    if (!configModule) {
        return { totalPermissions: result.permissions.length };
    }
    
    let externalCount = 0;
    let internalCount = 0;
    let linkCount = 0;
    let directGrantCount = 0;
    
    result.permissions.forEach(permission => {
        const classification = configModule.classifyPermission(permission, configModule.tenantDomains);
        
        if (classification === 'external') externalCount++;
        else if (classification === 'internal') internalCount++;
        
        if (permission.link) {
            linkCount++;
        } else {
            directGrantCount++;
        }
    });
    
    return {
        totalPermissions: result.permissions.length,
        externalCount,
        internalCount,
        linkCount,
        directGrantCount
    };
}

// Export functions for use in other modules
window.permissionsModule = {
    // Main action button functions
    showExpirationDialog,
    disableLinks,
    disableAllSharing,
    
    // Supporting permission management functions
    setExpirationForItem,
    removeSharingLinksForItem,
    removeAllSharingForItem,
    
    // Bulk operation functions
    addBulkPermission,
    removeBulkPermission,
    modifyBulkPermission,
    processSingleBulkOperation,
    
    // Validation and logging
    validatePermissionChange,
    logPermissionChange,
    
    // Utility functions
    getPermissionSummary
};
