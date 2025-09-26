// api.js - API Module for SharePoint & OneDrive Scanner v3.0
// Handles Microsoft Graph API requests, throttling, rate limiting, and request queue management

let globalThrottleState = { isThrottled: false, resumeTime: 0 };

const APP_INFO = {
    name: "SharePoint & OneDrive Scanner Enhanced",
    version: "3.0.0",
    userAgent: "NONISV|YourCompany|SharePointOneDriveScanner/3.0.0"
};

// REQUEST QUEUE CLASS FOR THROTTLING AND RATE LIMITING
class RequestQueue {
    constructor(maxConcurrent = 2, delayBetweenRequests = 500) {
        this.maxConcurrent = maxConcurrent;
        this.delayBetweenRequests = delayBetweenRequests;
        this.queue = [];
        this.running = 0;
    }

    async add(requestFn) {
        return new Promise((resolve, reject) => {
            this.queue.push({ requestFn, resolve, reject });
            this.process();
        });
    }

    async process() {
        if (this.running >= this.maxConcurrent || this.queue.length === 0) {
            return;
        }

        if (globalThrottleState.isThrottled && Date.now() < globalThrottleState.resumeTime) {
            setTimeout(() => this.process(), Math.max(100, globalThrottleState.resumeTime - Date.now()));
            return;
        }

        this.running++;
        const { requestFn, resolve, reject } = this.queue.shift();

        try {
            const result = await requestFn();
            resolve(result);
        } catch (error) {
            reject(error);
        } finally {
            this.running--;
            setTimeout(() => this.process(), this.delayBetweenRequests);
        }
    }
}

const requestQueue = new RequestQueue(6, 200);

// Utility function for delays
function delay(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
}

// ENHANCED GRAPH API REQUEST WITH RETRY, THROTTLING, AND AUTOMATIC TOKEN REFRESH
async function graphRequestWithRetry(url, options = {}, maxRetries = 3) {
    const authModule = window.authModule;
    if (!authModule) {
        throw new Error('Authentication module not available');
    }
    
    // Check if token needs refresh before making request
    let accessToken = authModule.accessToken;
    if (!accessToken || authModule.isTokenExpiring(5)) {
        console.log('🔄 Token expiring soon, refreshing before API request...');
        try {
            accessToken = await authModule.refreshTokenIfNeeded();
        } catch (refreshError) {
            console.error('❌ Token refresh failed before API request:', refreshError);
            throw new Error('Authentication token expired and refresh failed. Please sign in again.');
        }
    }

    for (let attempt = 0; attempt <= maxRetries; attempt++) {
        try {
            const response = await fetch(url, {
                ...options,
                headers: { 
                    Authorization: `Bearer ${accessToken}`,
                    'User-Agent': APP_INFO.userAgent,
                    ...options.headers 
                }
            });

            const rateLimitRemaining = response.headers.get('RateLimit-Remaining');
            const rateLimitReset = response.headers.get('RateLimit-Reset');
            
            if (rateLimitRemaining !== null) {
                const remaining = parseInt(rateLimitRemaining);
                const reset = parseInt(rateLimitReset);
                
                console.log(`API Rate limit: ${remaining} remaining, resets in ${reset}s`);
                
                if (remaining < 100) {
                    console.warn(`APPROACHING RATE LIMIT (${remaining} remaining), slowing down requests`);
                    await delay(1000);
                }
            }

            if (response.status === 429 || response.status === 503) {
                const retryAfter = response.headers.get('Retry-After');
                const waitTime = retryAfter ? 
                    parseInt(retryAfter) * 1000 : 
                    Math.pow(2, attempt) * 1000 + Math.random() * 1000;
                
                console.warn(`THROTTLED (${response.status}), pausing ALL requests for ${waitTime}ms`);
                
                globalThrottleState.isThrottled = true;
                globalThrottleState.resumeTime = Date.now() + waitTime;
                
                if (attempt < maxRetries) {
                    await delay(waitTime);
                    globalThrottleState.isThrottled = false;
                    continue;
                }
            }

            if (response.status === 501) {
                const errorText = await response.text();
                if (errorText.includes('notSupported') || errorText.includes('Permission is not supported')) {
                    console.log(`Item doesn't support permissions (HTTP 501) - skipping retries: ${url.split('/').pop()}`);
                    const error = new Error(`HTTP ${response.status}: ${errorText}`);
                    error.isNonRetryable = true;
                    throw error;
                }
            }

            // Handle 401 Unauthorized - token likely expired
            if (response.status === 401) {
                console.warn('🔄 Received 401 Unauthorized - attempting token refresh...');
                try {
                    // Force token refresh and update our access token
                    accessToken = await authModule.refreshTokenIfNeeded(true);
                    
                    if (attempt < maxRetries) {
                        console.log('✅ Token refreshed, retrying request...');
                        continue; // Retry with new token
                    }
                } catch (tokenRefreshError) {
                    console.error('❌ Token refresh failed on 401 response:', tokenRefreshError);
                    const error = new Error('Authentication token expired and refresh failed. Please sign in again.');
                    error.isAuthenticationError = true;
                    throw error;
                }
            }

            if (response.status === 404) {
                const errorText = await response.text();
                if (errorText.includes('mysite not found') || errorText.includes('ResourceNotFound')) {
                    console.log(`OneDrive not provisioned (HTTP 404) - skipping retries`);
                    const error = new Error(`HTTP ${response.status}: ${errorText}`);
                    error.isNonRetryable = true;
                    throw error;
                }
            }

            if (!response.ok) {
                const text = await response.text();
                throw new Error(`HTTP ${response.status}: ${text}`);
            }

            return response;
        } catch (error) {
            // Don't retry authentication errors or non-retryable errors
            if (error.isNonRetryable || error.isAuthenticationError || attempt === maxRetries) {
                throw error;
            }
            
            const waitTime = Math.pow(2, attempt) * 1000 + Math.random() * 1000;
            console.warn(`Request failed, retrying in ${waitTime}ms:`, error.message);
            await delay(waitTime);
        }
    }
}

// GET ALL ITEMS FROM PAGINATED API RESPONSE
async function graphGetAll(url) {
    let items = [];
    let next = url;
    while (next) {
        const resp = await graphRequestWithRetry(next);
        const j = await resp.json();
        items = items.concat(j.value || []);
        next = j['@odata.nextLink'] || null;
        
        // Dynamic throttling - only delay if rate limit headers suggest we should
        const rateLimitRemaining = resp.headers.get('RateLimit-Remaining');
        if (rateLimitRemaining !== null && parseInt(rateLimitRemaining) < 50) {
            await delay(200); // Slow down when approaching limits
        } else if (next) {
            await delay(25); // Minimal delay for pagination, only if there are more pages
        }
    }
    return items;
}

// BATCH PERMISSIONS FUNCTION FOR EFFICIENT PERMISSION CHECKING
async function batchGetPermissions(requests, controller = { stop: false }) {
    const batchSize = 15; // Optimized: increased from 5 to 15 for better throughput
    const results = [];
    
    for (let i = 0; i < requests.length; i += batchSize) {
        if (controller.stop) break;
        
        const batch = requests.slice(i, i + batchSize);
        
        try {
            const batchRequests = batch.map((req, index) => ({
                id: (i + index).toString(),
                method: "GET",
                url: req.url.replace('https://graph.microsoft.com/v1.0', '')
            }));

            const response = await requestQueue.add(async () => {
                return await graphRequestWithRetry('https://graph.microsoft.com/v1.0/$batch', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ requests: batchRequests })
                });
            });

            const batchResult = await response.json();

            for (const response of batchResult.responses) {
                const itemIndex = parseInt(response.id) - i;
                if (response.status === 200 && response.body && response.body.value) {
                    results.push({ 
                        item: batch[itemIndex].item, 
                        permissions: response.body.value 
                    });
                } else {
                    results.push({ 
                        item: batch[itemIndex].item, 
                        permissions: [] 
                    });
                }
            }

        } catch (error) {
            console.warn(`Batch failed, using individual requests:`, error);
            
            for (const req of batch) {
                if (controller.stop) break;
                try {
                    const permissions = await requestQueue.add(async () => {
                        return await graphGetAll(req.url);
                    });
                    results.push({ item: req.item, permissions });
                } catch (e) {
                    console.warn(`Failed individual request for ${req.item.name}:`, e);
                    results.push({ item: req.item, permissions: [] });
                }
            }
        }

        await delay(500);
    }
    
    return results;
}

// TENANT DOMAIN LOADING FOR INTERNAL/EXTERNAL CLASSIFICATION
async function loadTenantDomains() {
    try {
        console.log('🔍 LOADING TENANT DOMAINS - Including .onmicrosoft.com domains per custom instructions...');
        
        const domains = await requestQueue.add(async () => {
            return await graphGetAll('https://graph.microsoft.com/v1.0/domains');
        });
        
        const tenantDomains = new Set();
        
        // STRICT VALIDATION: Only add domains that are verified AND owned by the tenant
        for (const domain of domains) {
            if (domain.isVerified === true && domain.isDefault !== false) {
                const domainName = domain.id.toLowerCase().trim();
                if (domainName && domainName.includes('.')) {
                    tenantDomains.add(domainName);
                    console.log(`✅ VERIFIED TENANT DOMAIN: ${domainName} (isVerified: ${domain.isVerified}, isDefault: ${domain.isDefault})`);
                }
            } else {
                console.log(`❌ REJECTED DOMAIN: ${domain.id} (isVerified: ${domain.isVerified}, isDefault: ${domain.isDefault})`);
            }
        }
        
        // CUSTOM INSTRUCTION: Add .onmicrosoft.com domains as internal
        // Find the tenant's default .onmicrosoft.com domain
        const onMicrosoftDomains = domains.filter(d => 
            d.id.toLowerCase().endsWith('.onmicrosoft.com') && d.isVerified === true
        );
        
        for (const onMsDomain of onMicrosoftDomains) {
            const domainName = onMsDomain.id.toLowerCase().trim();
            if (!tenantDomains.has(domainName)) {
                tenantDomains.add(domainName);
                console.log(`✅ ADDED TENANT .ONMICROSOFT.COM DOMAIN: ${domainName} (per custom instructions)`);
            }
        }
        
        // Only add current user's domain if it's not already included and passes validation
        const account = window.authModule ? window.authModule.account : null;
        if (account && account.username) {
            const primaryDomain = account.username.split('@')[1];
            if (primaryDomain && !tenantDomains.has(primaryDomain.toLowerCase())) {
                // Double-check this domain is actually verified in the tenant OR is .onmicrosoft.com
                const matchingDomain = domains.find(d => 
                    d.id.toLowerCase() === primaryDomain.toLowerCase() && 
                    (d.isVerified === true || d.id.toLowerCase().endsWith('.onmicrosoft.com'))
                );
                if (matchingDomain) {
                    tenantDomains.add(primaryDomain.toLowerCase());
                    console.log(`✅ ADDED CURRENT USER DOMAIN: ${primaryDomain} (verified in tenant)`);
                } else {
                    console.warn(`⚠️ CURRENT USER DOMAIN NOT VERIFIED IN TENANT: ${primaryDomain}`);
                }
            }
        }
        
        console.log(`🎯 FINAL TENANT DOMAINS (${tenantDomains.size}):`, Array.from(tenantDomains));
        
        if (tenantDomains.size === 0) {
            throw new Error('No verified tenant domains found - this will affect internal/external classification');
        }
        
        return tenantDomains;
        
    } catch (e) {
        console.error('❌ CRITICAL: Failed to load tenant domains:', e);
        const tenantDomains = new Set();
        
        // Fallback: Only use current user's domain if available, but log this as a critical issue
        const account = window.authModule ? window.authModule.account : null;
        if (account && account.username) {
            const primaryDomain = account.username.split('@')[1];
            if (primaryDomain) {
                tenantDomains.add(primaryDomain.toLowerCase());
                console.warn(`⚠️ FALLBACK: Using only current user domain: ${primaryDomain}`);
            }
        }
        
        if (tenantDomains.size === 0) {
            console.error('🚨 CRITICAL: No tenant domains available for internal/external classification');
        }
        
        return tenantDomains;
    }
}

// DISCOVER SHAREPOINT SITES
async function discoverSharePointSites() {
    try {
        const response = await graphRequestWithRetry('https://graph.microsoft.com/v1.0/sites?search=*');
        const data = await response.json();
        return data.value || [];
    } catch (error) {
        console.error('Error discovering SharePoint sites:', error);
        throw error;
    }
}

// DISCOVER ONEDRIVE USERS
async function discoverOneDriveUsers() {
    try {
        const response = await graphRequestWithRetry('https://graph.microsoft.com/v1.0/users?$select=id,displayName,userPrincipalName,mail&$top=50');
        const data = await response.json();
        return data.value || [];
    } catch (error) {
        console.error('Error discovering OneDrive users:', error);
        throw error;
    }
}

// GET DRIVES FOR A SITE
async function getSiteDrives(siteId) {
    try {
        return await requestQueue.add(async () => {
            return await graphGetAll(`https://graph.microsoft.com/v1.0/sites/${siteId}/drives`);
        });
    } catch (error) {
        console.error(`Error getting drives for site ${siteId}:`, error);
        throw error;
    }
}

// GET ONEDRIVE FOR A USER
async function getUserOneDrive(userId) {
    try {
        const response = await requestQueue.add(async () => {
            return await graphRequestWithRetry(`https://graph.microsoft.com/v1.0/users/${userId}/drive`);
        });
        return await response.json();
    } catch (error) {
        console.error(`Error getting OneDrive for user ${userId}:`, error);
        throw error;
    }
}

// DELTA QUERY FOR EFFICIENT SCANNING
async function performDeltaQuery(driveId) {
    try {
        const deltaUrl = `https://graph.microsoft.com/v1.0/drives/${driveId}/root/delta?$expand=permissions&$select=id,name,folder,file,parentReference,permissions,createdBy,lastModifiedBy`;
        
        let allItems = [];
        let nextUrl = deltaUrl;
        let pageCount = 0;
        
        while (nextUrl) {
            const response = await requestQueue.add(async () => {
                return await graphRequestWithRetry(nextUrl);
            });
            
            const data = await response.json();
            const newItems = data.value || [];
            allItems = allItems.concat(newItems);
            nextUrl = data['@odata.nextLink'];
            pageCount++;
            
            console.log(`Delta page ${pageCount}: ${newItems.length} items, total: ${allItems.length}`);
            
            await delay(200);
        }
        
        return allItems;
    } catch (error) {
        console.error(`Delta query failed for drive ${driveId}:`, error);
        throw error;
    }
}

// GET FOLDER CHILDREN WITH FILTERING
async function getFolderChildren(driveId, itemId, includeFiles = false) {
    try {
        let url;
        if (includeFiles) {
            url = itemId === "root"
                ? `https://graph.microsoft.com/v1.0/drives/${driveId}/root/children?$select=id,name,folder,file,parentReference`
                : `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${itemId}/children?$select=id,name,folder,file,parentReference`;
        } else {
            url = itemId === "root"
                ? `https://graph.microsoft.com/v1.0/drives/${driveId}/root/children?$select=id,name,folder,parentReference&$filter=folder ne null`
                : `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${itemId}/children?$select=id,name,folder,parentReference&$filter=folder ne null`;
        }
        
        return await requestQueue.add(async () => {
            return await graphGetAll(url);
        });
    } catch (error) {
        console.error(`Error getting children for ${itemId} in drive ${driveId}:`, error);
        return [];
    }
}

// Export functions and variables for use in other modules
window.apiModule = {
    // Classes
    RequestQueue,
    
    // Variables
    requestQueue,
    globalThrottleState,
    APP_INFO,
    
    // Functions
    delay,
    graphRequestWithRetry,
    graphGetAll,
    batchGetPermissions,
    loadTenantDomains,
    discoverSharePointSites,
    discoverOneDriveUsers,
    getSiteDrives,
    getUserOneDrive,
    performDeltaQuery,
    getFolderChildren
};
