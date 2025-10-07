// main.js - Main Application Module for SharePoint & OneDrive Scanner v3.0
// Coordinates all modules and initializes the application

console.log('*** SHAREPOINT ONEDRIVE SCANNER V3.0 WITH MODULAR ARCHITECTURE ***');

// CLEAR RESULTS BUTTON INITIALIZATION
function initializeClearResultsButton() {
    const clearResultsBtn = document.getElementById('clear-results-btn');
    if (clearResultsBtn) {
        clearResultsBtn.addEventListener('click', function() {
            if (window.configModule) {
                // Clear results using the existing clearResults function
                window.configModule.clearResults();
                
                // Also hide the view controls and results actions
                const viewControlsContainer = document.getElementById('view-controls-container');
                if (viewControlsContainer) {
                    viewControlsContainer.style.display = 'none';
                }
                
                const bulkControls = document.getElementById('bulk-controls');
                if (bulkControls) {
                    bulkControls.style.display = 'none';
                }
                
                // Show success message
                if (window.configModule.showToast) {
                    window.configModule.showToast('Results cleared successfully', 2000);
                }
                
                console.log('🗑️ Results cleared by user');
            }
        });
        console.log('✅ Clear results button initialized');
    } else {
        console.warn('⚠️ Clear results button not found in DOM');
    }
}

// DEBUG CONSOLE INITIALIZATION
function initializeDebugConsole() {
    const checkbox = document.getElementById('enable-debug-console');
    if (checkbox) {
        checkbox.addEventListener('change', function() {
            const enabled = this.checked;
            if (window.configModule) {
                window.configModule.setDebugEnabled(enabled);
                console.log(`🐛 Debug console output ${enabled ? 'ENABLED' : 'DISABLED'}`);
                if (window.configModule.showToast) {
                    window.configModule.showToast(`Debug console ${enabled ? 'enabled' : 'disabled'}`, 2000);
                }
            }
        });
        console.log('✅ Debug console checkbox initialized');
    } else {
        console.warn('⚠️ Debug console checkbox not found in DOM');
    }
}

// MAIN APPLICATION INITIALIZATION
document.addEventListener('DOMContentLoaded', function() {
    console.log('🚀 Starting SharePoint & OneDrive Scanner v3.0 initialization...');
    
    try {
        // Check that all required modules are loaded
        const requiredModules = [
            'configModule',
            'authModule', 
            'apiModule',
            'uiModule',
            'scanningModule',
            'resultsModule',
            'permissionsModule',
            'exportModule'
        ];
        
        console.log('📋 Checking module dependencies...');
        const missingModules = [];
        
        for (const moduleName of requiredModules) {
            if (!window[moduleName]) {
                missingModules.push(moduleName);
            }
        }
        
        if (missingModules.length > 0) {
            console.error('❌ Missing required modules:', missingModules);
            alert(`Missing required modules: ${missingModules.join(', ')}. Please refresh the page.`);
            return;
        }
        
        console.log('✅ All required modules loaded successfully');
        
        // Initialize modules in the correct order
        console.log('🔧 Initializing modules...');
        
        // 1. Initialize authentication handlers
        if (window.authModule.initializeAuthenticationHandlers) {
            window.authModule.initializeAuthenticationHandlers();
            console.log('✅ Authentication module initialized');
        }
        
        // 2. Initialize UI components
        if (window.uiModule.initializeUI) {
            window.uiModule.initializeUI();
            console.log('✅ UI module initialized');
        }
        
        // 3. Initialize results filtering
        if (window.resultsModule.initializeResultsFiltering) {
            window.resultsModule.initializeResultsFiltering();
            console.log('✅ Results filtering initialized');
        }
        
        // 4. Initialize view toggle functionality
        if (window.resultsModule.initializeViewToggle) {
            window.resultsModule.initializeViewToggle();
            console.log('✅ View toggle initialized');
        }
        
        // 5. Initialize SharePoint groups toggle
        if (window.resultsModule.initializeSharePointGroupsToggle) {
            window.resultsModule.initializeSharePointGroupsToggle();
            console.log('✅ SharePoint groups toggle initialized');
        }
        
        // 6. Initialize direct grants toggle
        if (window.resultsModule.initializeDirectGrantsToggle) {
            window.resultsModule.initializeDirectGrantsToggle();
            console.log('✅ Direct grants toggle initialized');
        }
        
        // 7. Initialize export functionality
        if (window.exportModule.initializeExportModule) {
            window.exportModule.initializeExportModule();
            console.log('✅ Export module initialized');
        }
        
        // 8. Initialize debug console toggle
        initializeDebugConsole();
        console.log('✅ Debug console initialized');
        
        // 9. Initialize clear results button
        initializeClearResultsButton();
        console.log('✅ Clear results button initialized');
        
        // 10. Check for existing authentication on page load
        if (window.authModule.checkExistingAuthentication) {
            window.authModule.checkExistingAuthentication();
            console.log('✅ Authentication check initiated');
        }
        
        console.log('🎉 SharePoint & OneDrive Scanner v3.0 initialization completed successfully!');
        console.log('📱 Application ready for use');
        
        // Add global error handler
        window.addEventListener('error', function(e) {
            console.error('💥 Global application error:', e.error);
            if (window.configModule && window.configModule.showToast) {
                window.configModule.showToast('Application error occurred - check console for details', 5000);
            }
        });
        
        // Add unhandled promise rejection handler
        window.addEventListener('unhandledrejection', function(e) {
            console.error('💥 Unhandled promise rejection:', e.reason);
            if (window.configModule && window.configModule.showToast) {
                window.configModule.showToast('Async operation failed - check console for details', 5000);
            }
        });
        
        console.log('✅ Global error handlers registered');
        
    } catch (error) {
        console.error('❌ Critical error during application initialization:', error);
        alert('Critical application error during initialization. Please refresh the page and try again.');
    }
});

// GLOBAL UTILITY FUNCTIONS AVAILABLE TO ALL MODULES
window.appUtils = {
    // Version information
    version: '3.0.0',
    name: 'SharePoint & OneDrive Scanner Enhanced',
    
    // Module status checker
    checkModuleStatus: function() {
        const modules = [
            'configModule',
            'authModule', 
            'apiModule',
            'uiModule',
            'scanningModule',
            'resultsModule',
            'permissionsModule',
            'exportModule'
        ];
        
        const status = {};
        modules.forEach(moduleName => {
            status[moduleName] = {
                loaded: !!window[moduleName],
                functions: window[moduleName] ? Object.keys(window[moduleName]).length : 0
            };
        });
        
        return status;
    },
    
    // Debug information
    getDebugInfo: function() {
        const configModule = window.configModule;
        
        return {
            version: this.version,
            timestamp: new Date().toISOString(),
            authenticated: !!(window.authModule && window.authModule.account),
            resultsCount: configModule ? configModule.results.length : 0,
            sitesCount: configModule ? configModule.sites.length : 0,
            usersCount: configModule ? configModule.users.length : 0,
            scanning: configModule ? configModule.scanning : false,
            moduleStatus: this.checkModuleStatus(),
            userAgent: navigator.userAgent,
            location: window.location.href
        };
    },
    
    // Performance monitoring
    performanceMonitor: {
        startTime: performance.now(),
        markers: {},
        
        mark: function(name) {
            this.markers[name] = performance.now();
            console.log(`⏱️ Performance marker: ${name} at ${(this.markers[name] - this.startTime).toFixed(2)}ms`);
        },
        
        measure: function(startMark, endMark) {
            if (this.markers[startMark] && this.markers[endMark]) {
                const duration = this.markers[endMark] - this.markers[startMark];
                console.log(`📏 Performance measure: ${startMark} to ${endMark} = ${duration.toFixed(2)}ms`);
                return duration;
            }
            return null;
        },
        
        getReport: function() {
            const currentTime = performance.now();
            const totalTime = currentTime - this.startTime;
            
            return {
                totalTime: totalTime,
                markers: { ...this.markers },
                memory: performance.memory ? {
                    usedJSMemory: performance.memory.usedJSMemory,
                    totalJSMemory: performance.memory.totalJSMemory
                } : null
            };
        }
    }
};

// Mark application start
window.appUtils.performanceMonitor.mark('app_start');

// Export for debugging
window.debugScanner = function() {
    console.log('🔍 SharePoint & OneDrive Scanner Debug Information:');
    console.log(window.appUtils.getDebugInfo());
    console.log('⏱️ Performance Report:');
    console.log(window.appUtils.performanceMonitor.getReport());
};

// Health check function
window.healthCheck = function() {
    console.log('🏥 Running health check...');
    
    const issues = [];
    
    // Check modules
    const moduleStatus = window.appUtils.checkModuleStatus();
    Object.entries(moduleStatus).forEach(([name, status]) => {
        if (!status.loaded) {
            issues.push(`Module not loaded: ${name}`);
        }
    });
    
    // Check authentication
    if (window.authModule && !window.authModule.account) {
        issues.push('Not authenticated - sign in required');
    }
    
    // Check for errors in console
    const hasErrors = console.error.toString().includes('called');
    if (hasErrors) {
        issues.push('Console errors detected');
    }
    
    if (issues.length === 0) {
        console.log('✅ Health check passed - application is healthy');
        if (window.configModule && window.configModule.showToast) {
            window.configModule.showToast('Health check passed ✅', 3000);
        }
    } else {
        console.warn('⚠️ Health check found issues:', issues);
        if (window.configModule && window.configModule.showToast) {
            window.configModule.showToast(`Health check found ${issues.length} issues - check console`, 5000);
        }
    }
    
    return { healthy: issues.length === 0, issues };
};

console.log('📚 Main application module loaded. Waiting for DOM ready...');
