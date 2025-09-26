// auth.js - Authentication Module for SharePoint & OneDrive Scanner v3.0
// Handles MSAL authentication, login/logout, token management, and persistent sessions

let msalInstance = null;
let account = null;
let accessToken = '';

const requiredScopes = ["User.Read", "Sites.Read.All", "Files.Read.All", "Directory.Read.All", "Files.ReadWrite.All"];

// ENHANCED AUTHENTICATION STATE MANAGEMENT FOR PAGE REFRESH BEHAVIOR
async function checkExistingAuthentication() {
    console.log('🔍 ENHANCED AUTHENTICATION CHECK ON PAGE REFRESH...');
    
    try {
        // Check if we have stored tenant/client configuration
        const storedTenantId = localStorage.getItem('sp_scanner_tenant_id');
        const storedClientId = localStorage.getItem('sp_scanner_client_id');
        
        if (storedTenantId && storedClientId) {
            document.getElementById('tenant-id').value = storedTenantId;
            document.getElementById('client-id').value = storedClientId;
            console.log('✅ Restored saved tenant/client configuration');
        }
        
        // Only proceed if we have both IDs
        if (!storedTenantId || !storedClientId) {
            console.log('❌ No stored authentication configuration found - requiring fresh login');
            updateAuthenticationUI(false);
            showToast('Please enter your tenant and client credentials to sign in', 3000);
            return false;
        }
        
        // Load MSAL and initialize
        await loadMSAL();
        
        msalInstance = new msal.PublicClientApplication({
            auth: {
                clientId: storedClientId,
                authority: `https://login.microsoftonline.com/${storedTenantId}`,
                redirectUri: window.location.origin
            }
        });

        if (msalInstance.initialize) {
            await msalInstance.initialize();
        }
        
        // Check for existing accounts - critical for page refresh persistence
        const accounts = msalInstance.getAllAccounts();
        console.log(`🔍 Found ${accounts?.length || 0} cached accounts after page refresh`);
        
        if (accounts && accounts.length > 0) {
            account = accounts[0];
            console.log(`🔍 Attempting to restore session for: ${account.username}`);
            
            // Try to acquire token silently - this is the key for page refresh persistence
            try {
                console.log('🔄 Attempting silent token acquisition for existing session...');
                const tokenResult = await msalInstance.acquireTokenSilent({
                    account: account,
                    scopes: requiredScopes,
                    forceRefresh: false // Allow cached tokens for better UX
                });
                
                accessToken = tokenResult.accessToken;
                console.log('✅ Successfully restored authentication session with valid token');
                
                // Validate the token with a lightweight API call
                try {
                    const testResponse = await graphRequestWithRetry('https://graph.microsoft.com/v1.0/me', {}, 1); // Single retry
                    const userData = await testResponse.json();
                    
                    console.log('✅ Token validation successful - user authenticated:', userData.userPrincipalName);
                    updateAuthenticationUI(true, userData.userPrincipalName || account.username);
                    
                    // Start token refresh monitoring for long-running operations
                    startTokenRefreshMonitoring();
                    
                    // Save authentication success for future page loads
                    localStorage.setItem('sp_scanner_last_auth_success', Date.now().toString());
                    
                    return true;
                    
                } catch (apiError) {
                    console.warn('❌ Token validation failed - API test unsuccessful:', apiError.message);
                    throw apiError; // Let the outer catch handle this as token failure
                }
                
            } catch (tokenError) {
                console.warn('❌ Silent token acquisition failed:', tokenError.message);
                
                // Clear invalid session state and require fresh login
                await clearAuthenticationState();
                updateAuthenticationUI(false);
                
                if (tokenError.message.includes('interaction_required') || tokenError.message.includes('login_required')) {
                    showToast('Session expired - please sign in again to continue', 4000);
                } else {
                    showToast('Authentication issue detected - please sign in again', 4000);
                }
                
                return false;
            }
        } else {
            console.log('❌ No cached accounts found - requiring fresh login');
            updateAuthenticationUI(false);
            showToast('No existing session found - please sign in', 3000);
            return false;
        }
        
    } catch (error) {
        console.error('❌ Critical error during authentication check:', error);
        await clearAuthenticationState();
        updateAuthenticationUI(false);
        showToast('Authentication system error - please try signing in again', 4000);
        return false;
    }
}

// NEW: Clear all authentication state function
async function clearAuthenticationState() {
    console.log('🧹 CLEARING ALL AUTHENTICATION STATE...');
    
    try {
        // Stop token refresh monitoring
        stopTokenRefreshMonitoring();
        
        // Clear application variables
        account = null;
        accessToken = '';
        tokenExpirationTime = 0;
        
        // Clear MSAL cache if available
        if (msalInstance) {
            try {
                const accounts = msalInstance.getAllAccounts();
                if (accounts && accounts.length > 0) {
                    for (const acc of accounts) {
                        await msalInstance.removeAccount(acc);
                    }
                    console.log('✅ Cleared MSAL account cache');
                }
            } catch (msalError) {
                console.warn('⚠️ Error clearing MSAL cache:', msalError);
            }
        }
        
        // Clear application data - these will be imported from other modules
        if (typeof clearResults === 'function') clearResults();
        if (typeof clearSitesAndUsers === 'function') clearSitesAndUsers();
        
        // Remove authentication success marker
        localStorage.removeItem('sp_scanner_last_auth_success');
        
        console.log('✅ Authentication state cleared successfully');
        
    } catch (error) {
        console.error('❌ Error clearing authentication state:', error);
    }
}

function updateAuthenticationUI(isAuthenticated, username = '') {
    const userDisplay = document.getElementById('user-display');
    const authStatus = document.getElementById('auth-status');
    const logoutBtn = document.getElementById('logout-btn');
    const tabsContainer = document.getElementById('tabs-container');
    const loginBtn = document.getElementById('login-btn');
    
    if (isAuthenticated) {
        userDisplay.innerText = username || 'Authenticated User';
        authStatus.innerText = 'Connected';
        authStatus.className = 'status-badge status-approved';
        logoutBtn.style.display = 'inline-flex';
        tabsContainer.style.display = 'block';
        loginBtn.innerText = 'Already Signed In';
        loginBtn.disabled = true;
        
        console.log('✅ UI updated for authenticated state');
        showToast(`Welcome back, ${username}!`, 3000);
    } else {
        userDisplay.innerText = 'Not signed in';
        authStatus.innerText = 'Not Connected';
        authStatus.className = 'status-badge status-info';
        logoutBtn.style.display = 'none';
        tabsContainer.style.display = 'none';
        loginBtn.innerText = 'Sign In';
        loginBtn.disabled = false;
        
        // Clear any existing data - these will be imported from other modules
        if (typeof clearResults === 'function') clearResults();
        if (typeof clearSitesAndUsers === 'function') clearSitesAndUsers();
        
        console.log('✅ UI updated for unauthenticated state');
    }
}

async function performLogout() {
    console.log('� PERFORMING LOGOUT...');
    
    try {
        // Clear application state
        account = null;
        accessToken = '';
        
        // Clear MSAL cache if instance exists
        if (msalInstance) {
            try {
                const accounts = msalInstance.getAllAccounts();
                if (accounts && accounts.length > 0) {
                    // Remove all accounts from cache
                    for (const acc of accounts) {
                        await msalInstance.removeAccount(acc);
                    }
                }
                console.log('✅ Cleared MSAL account cache');
            } catch (msalError) {
                console.warn('⚠️ Error clearing MSAL cache:', msalError);
            }
        }
        
        // Clear application data - these will be imported from other modules
        if (typeof clearResults === 'function') clearResults();
        if (typeof clearSitesAndUsers === 'function') clearSitesAndUsers();
        
        // Update UI
        updateAuthenticationUI(false);
        
        showToast('Successfully signed out', 3000);
        console.log('✅ Logout completed successfully');
        
    } catch (error) {
        console.error('❌ Error during logout:', error);
        showToast('Error during logout: ' + error.message, 4000);
        
        // Force UI update even if logout had errors
        updateAuthenticationUI(false);
    }
}

function saveAuthenticationConfig(tenantId, clientId) {
    try {
        localStorage.setItem('sp_scanner_tenant_id', tenantId);
        localStorage.setItem('sp_scanner_client_id', clientId);
        console.log('✅ Saved authentication configuration to localStorage');
    } catch (error) {
        console.warn('⚠️ Could not save authentication configuration:', error);
    }
}

// Load MSAL library
function loadMSAL() {
    return new Promise((resolve, reject) => {
        if (typeof msal !== 'undefined') {
            resolve();
            return;
        }
        
        const script = document.createElement('script');
        script.src = 'https://alcdn.msauth.net/browser/2.37.0/js/msal-browser.min.js';
        script.onload = resolve;
        script.onerror = () => reject(new Error('Failed to load MSAL'));
        document.head.appendChild(script);
    });
}

// ENHANCED TOKEN ACQUISITION WITH EXPIRATION TRACKING
let tokenExpirationTime = 0;
let tokenRefreshInProgress = false;

async function acquireToken(forceRefresh = false) {
    try {
        const result = await msalInstance.acquireTokenSilent({
            account: account,
            scopes: requiredScopes,
            forceRefresh: forceRefresh
        });
        
        // Track token expiration (MSAL tokens typically expire in 1 hour)
        tokenExpirationTime = Date.now() + (55 * 60 * 1000); // Refresh 5 minutes before expiration
        console.log(`✅ Token acquired, expires at: ${new Date(tokenExpirationTime).toLocaleTimeString()}`);
        
        return result.accessToken;
    } catch (e) {
        console.log('🔄 Silent token acquisition failed, using popup...');
        const result = await msalInstance.acquireTokenPopup({
            account: account,
            scopes: requiredScopes
        });
        
        tokenExpirationTime = Date.now() + (55 * 60 * 1000);
        console.log(`✅ Token acquired via popup, expires at: ${new Date(tokenExpirationTime).toLocaleTimeString()}`);
        
        return result.accessToken;
    }
}

// AUTOMATIC TOKEN REFRESH MECHANISM
async function refreshTokenIfNeeded(forceRefresh = false) {
    if (!msalInstance || !account) {
        throw new Error('Authentication not initialized');
    }
    
    const now = Date.now();
    const timeUntilExpiration = tokenExpirationTime - now;
    
    // Check if token is expiring soon (within 5 minutes) or force refresh requested
    if (forceRefresh || timeUntilExpiration <= 0 || !accessToken) {
        if (tokenRefreshInProgress) {
            // Wait for ongoing refresh to complete
            console.log('🔄 Token refresh already in progress, waiting...');
            while (tokenRefreshInProgress) {
                await new Promise(resolve => setTimeout(resolve, 100));
            }
            return accessToken;
        }
        
        console.log(`🔄 Refreshing token (${forceRefresh ? 'forced' : 'expiring in ' + Math.round(timeUntilExpiration / 1000) + 's'})...`);
        tokenRefreshInProgress = true;
        
        try {
            const newToken = await acquireToken(true);
            accessToken = newToken;
            
            // Show user-friendly notification for proactive refresh
            if (!forceRefresh && timeUntilExpiration > 0) {
                showToast('🔄 Authentication token refreshed to continue scanning', 3000);
            }
            
            console.log('✅ Token refreshed successfully');
            return newToken;
        } catch (error) {
            console.error('❌ Token refresh failed:', error);
            
            // Clear invalid session and require re-authentication
            await clearAuthenticationState();
            updateAuthenticationUI(false);
            showToast('⚠️ Session expired - please sign in again to continue', 5000);
            
            throw new Error('Token refresh failed - authentication required');
        } finally {
            tokenRefreshInProgress = false;
        }
    }
    
    return accessToken;
}

// PROACTIVE TOKEN REFRESH MONITORING WITH ENHANCED SCANNING SUPPORT
let tokenRefreshInterval = null;
let scanningTokenRefreshInterval = null;
let isScanning = false;

function startTokenRefreshMonitoring(enableScanningMode = false) {
    // Stop any existing monitoring
    if (tokenRefreshInterval) {
        clearInterval(tokenRefreshInterval);
    }
    if (scanningTokenRefreshInterval) {
        clearInterval(scanningTokenRefreshInterval);
    }
    
    // Standard monitoring - check token expiration every 2 minutes during active usage
    tokenRefreshInterval = setInterval(async () => {
        try {
            if (account && accessToken) {
                const timeUntilExpiration = tokenExpirationTime - Date.now();
                
                // Proactively refresh if token expires within 10 minutes
                if (timeUntilExpiration <= 10 * 60 * 1000 && timeUntilExpiration > 0) {
                    console.log(`⏰ Standard token refresh triggered (expires in ${Math.round(timeUntilExpiration / 1000 / 60)} minutes)`);
                    await refreshTokenIfNeeded();
                }
            }
        } catch (error) {
            console.warn('⚠️ Standard token refresh failed:', error);
        }
    }, 2 * 60 * 1000); // Check every 2 minutes
    
    // Enhanced monitoring for scanning operations - more frequent checks
    if (enableScanningMode) {
        scanningTokenRefreshInterval = setInterval(async () => {
            try {
                if (account && accessToken && isScanning) {
                    const timeUntilExpiration = tokenExpirationTime - Date.now();
                    
                    // More aggressive refresh during scanning - refresh if expires within 15 minutes
                    if (timeUntilExpiration <= 15 * 60 * 1000 && timeUntilExpiration > 0) {
                        console.log(`🔄 Scanning-mode token refresh triggered (expires in ${Math.round(timeUntilExpiration / 1000 / 60)} minutes)`);
                        await refreshTokenIfNeeded();
                        showToast(`🔄 Token refreshed during scan to maintain connection`, 3000);
                    }
                    
                    // Warning if token will expire soon
                    if (timeUntilExpiration <= 5 * 60 * 1000 && timeUntilExpiration > 0) {
                        console.warn(`⚠️ Token expiring very soon during scan: ${Math.round(timeUntilExpiration / 1000)} seconds`);
                    }
                }
            } catch (error) {
                console.warn('⚠️ Scanning token refresh failed:', error);
                // Show user notification for scan-critical token refresh failures
                if (isScanning) {
                    showToast(`⚠️ Token refresh issue during scan - authentication may be required`, 5000);
                }
            }
        }, 60 * 1000); // Check every 1 minute during scanning
        
        console.log('✅ Enhanced scanning token refresh monitoring started');
    }
    
    console.log('✅ Token refresh monitoring started (scanning mode:', enableScanningMode, ')');
}

function stopTokenRefreshMonitoring() {
    if (tokenRefreshInterval) {
        clearInterval(tokenRefreshInterval);
        tokenRefreshInterval = null;
    }
    if (scanningTokenRefreshInterval) {
        clearInterval(scanningTokenRefreshInterval);
        scanningTokenRefreshInterval = null;
    }
    isScanning = false;
    console.log('🛑 Token refresh monitoring stopped');
}

// NEW: Scanning-specific token management functions
function startScanningTokenMonitoring() {
    isScanning = true;
    startTokenRefreshMonitoring(true);
    console.log('🔄 Enhanced token monitoring activated for scanning operation');
}

function stopScanningTokenMonitoring() {
    isScanning = false;
    if (scanningTokenRefreshInterval) {
        clearInterval(scanningTokenRefreshInterval);
        scanningTokenRefreshInterval = null;
    }
    // Keep standard monitoring active, just stop scanning-specific monitoring
    console.log('🛑 Scanning token monitoring stopped, reverting to standard monitoring');
}

// NEW: Force token refresh during scanning with better error handling
async function ensureValidTokenForScanning() {
    if (!account || !accessToken) {
        throw new Error('Not authenticated - cannot perform scanning operations');
    }
    
    const timeUntilExpiration = tokenExpirationTime - Date.now();
    
    // If token expires within 20 minutes during scanning, refresh it proactively
    if (timeUntilExpiration <= 20 * 60 * 1000) {
        console.log(`🔄 Proactively refreshing token for scanning (expires in ${Math.round(timeUntilExpiration / 1000 / 60)} minutes)`);
        try {
            const refreshedToken = await refreshTokenIfNeeded(true);
            showToast('🔄 Token refreshed to ensure continuous scanning', 2000);
            return refreshedToken;
        } catch (error) {
            console.error('❌ Failed to refresh token for scanning:', error);
            showToast('⚠️ Token refresh failed - scan may be interrupted', 5000);
            throw error;
        }
    }
    
    return accessToken;
}

// NEW: Get token expiration info for monitoring
function getTokenExpirationInfo() {
    if (!tokenExpirationTime) {
        return { 
            isValid: false, 
            timeUntilExpiration: 0, 
            minutesUntilExpiration: 0,
            expirationTime: null
        };
    }
    
    const timeUntilExpiration = tokenExpirationTime - Date.now();
    const minutesUntilExpiration = Math.round(timeUntilExpiration / 1000 / 60);
    
    return {
        isValid: timeUntilExpiration > 0,
        timeUntilExpiration,
        minutesUntilExpiration,
        expirationTime: new Date(tokenExpirationTime),
        isExpiring: timeUntilExpiration <= 15 * 60 * 1000 // Expiring within 15 minutes
    };
}

// CHECK IF TOKEN IS EXPIRING SOON
function isTokenExpiring(withinMinutes = 5) {
    if (!tokenExpirationTime) return true; // Assume expiring if no expiration tracked
    
    const timeUntilExpiration = tokenExpirationTime - Date.now();
    return timeUntilExpiration <= (withinMinutes * 60 * 1000);
}

async function performLogin() {
    const tenantId = document.getElementById('tenant-id').value.trim();
    const clientId = document.getElementById('client-id').value.trim();
    
    if (!tenantId || !clientId) {
        alert('Please enter both Tenant ID and Client ID');
        return;
    }

    const loginBtn = document.getElementById('login-btn');
    loginBtn.disabled = true;
    loginBtn.innerText = 'Signing in...';

    try {
        // Save configuration for future page loads
        saveAuthenticationConfig(tenantId, clientId);
        
        await loadMSAL();
        
        msalInstance = new msal.PublicClientApplication({
            auth: {
                clientId: clientId,
                authority: `https://login.microsoftonline.com/${tenantId}`,
                redirectUri: window.location.origin
            }
        });

        if (msalInstance.initialize) {
            await msalInstance.initialize();
        }

        const loginResult = await msalInstance.loginPopup({ scopes: requiredScopes });
        
        account = loginResult.account;
        accessToken = await acquireToken();

        // Update UI for successful authentication
        updateAuthenticationUI(true, account.username);
        
        // Start token refresh monitoring for long-running operations
        startTokenRefreshMonitoring();
        
        showToast(`Successfully signed in as ${account.username}`, 4000);
        console.log('✅ Fresh login successful and saved for future sessions');

    } catch (error) {
        console.error('Login failed:', error);
        showToast('Login failed - check console for details', 4000);
        alert('Login failed: ' + error.message);
        updateAuthenticationUI(false);
    } finally {
        loginBtn.disabled = false;
        loginBtn.innerText = 'Sign In';
    }
}

// Initialize authentication event handlers
function initializeAuthenticationHandlers() {
    // Login button handler
    const loginBtn = document.getElementById('login-btn');
    if (loginBtn) {
        loginBtn.addEventListener('click', performLogin);
    }

    // Logout button handler
    const logoutBtn = document.getElementById('logout-btn');
    if (logoutBtn) {
        logoutBtn.addEventListener('click', async function() {
            console.log('🚪 LOGOUT BUTTON CLICKED - Starting logout process...');
            
            this.disabled = true;
            this.innerText = 'Signing out...';
            
            try {
                await performLogout();
                console.log('✅ Logout completed successfully');
            } catch (error) {
                console.error('❌ Logout error:', error);
                showToast('Error during logout - clearing session anyway', 3000);
                
                // Force clear even if logout had errors
                await clearAuthenticationState();
                updateAuthenticationUI(false);
            } finally {
                this.disabled = false;
                this.innerText = 'Sign Out';
            }
        });
    }
}

// Export functions and variables for use in other modules
window.authModule = {
    // Variables
    get msalInstance() { return msalInstance; },
    get account() { return account; },
    get accessToken() { return accessToken; },
    set accessToken(value) { accessToken = value; },
    get tokenExpirationTime() { return tokenExpirationTime; },
    
    // Functions
    checkExistingAuthentication,
    clearAuthenticationState,
    updateAuthenticationUI,
    performLogout,
    saveAuthenticationConfig,
    loadMSAL,
    acquireToken,
    performLogin,
    initializeAuthenticationHandlers,
    
    // Token refresh functions
    refreshTokenIfNeeded,
    startTokenRefreshMonitoring,
    stopTokenRefreshMonitoring,
    isTokenExpiring,
    
    // NEW: Scanning-specific token management functions
    startScanningTokenMonitoring,
    stopScanningTokenMonitoring,
    ensureValidTokenForScanning,
    getTokenExpirationInfo
};
