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
        // Clear application variables
        account = null;
        accessToken = '';
        
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
    console.log('🚪 PERFORMING LOGOUT...');
    
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

async function acquireToken() {
    try {
        const result = await msalInstance.acquireTokenSilent({
            account: account,
            scopes: requiredScopes
        });
        return result.accessToken;
    } catch (e) {
        const result = await msalInstance.acquireTokenPopup({
            account: account,
            scopes: requiredScopes
        });
        return result.accessToken;
    }
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
    
    // Functions
    checkExistingAuthentication,
    clearAuthenticationState,
    updateAuthenticationUI,
    performLogout,
    saveAuthenticationConfig,
    loadMSAL,
    acquireToken,
    performLogin,
    initializeAuthenticationHandlers
};
