// virtual-dom.js - Virtual DOM Module for SharePoint & OneDrive Scanner v3.0
// Provides efficient DOM updates during CSV export operations

// VIRTUAL DOM STATE MANAGEMENT
let virtualDom = {
    pendingUpdates: [],
    isRendering: false,
    batchSize: 5, // Process 5 updates per frame
    rafId: null,
    enabled: false
};

// VIRTUAL DOM NODE STRUCTURE
class VirtualDOMNode {
    constructor(type, props = {}, children = []) {
        this.type = type;
        this.props = props;
        this.children = Array.isArray(children) ? children : [children].filter(Boolean);
        this.element = null;
        this.isVirtual = true;
    }
    
    // Create actual DOM element from virtual node
    render() {
        if (typeof this.type === 'string') {
            this.element = document.createElement(this.type);
            
            // Apply props
            Object.keys(this.props).forEach(prop => {
                if (prop === 'style' && typeof this.props[prop] === 'object') {
                    Object.assign(this.element.style, this.props[prop]);
                } else if (prop === 'className') {
                    this.element.className = this.props[prop];
                } else if (prop === 'onclick') {
                    this.element.onclick = this.props[prop];
                } else if (prop.startsWith('data-')) {
                    this.element.setAttribute(prop, this.props[prop]);
                } else if (prop !== 'children') {
                    this.element[prop] = this.props[prop];
                }
            });
            
            // Render children
            this.children.forEach(child => {
                if (typeof child === 'string') {
                    this.element.appendChild(document.createTextNode(child));
                } else if (child instanceof VirtualDOMNode) {
                    const childElement = child.render();
                    if (childElement) {
                        this.element.appendChild(childElement);
                    }
                }
            });
        }
        
        return this.element;
    }
}

// VIRTUAL DOM RESULT QUEUE
class ResultUpdateQueue {
    constructor() {
        this.queue = [];
        this.processing = false;
    }
    
    add(update) {
        this.queue.push(update);
        if (!this.processing) {
            this.process();
        }
    }
    
    process() {
        if (this.processing || this.queue.length === 0) return;
        
        this.processing = true;
        
        const processChunk = () => {
            const startTime = performance.now();
            let processed = 0;
            
            // Process items for up to 5ms or batch size limit
            while (this.queue.length > 0 && processed < virtualDom.batchSize && (performance.now() - startTime) < 5) {
                const update = this.queue.shift();
                try {
                    update.execute();
                    processed++;
                } catch (error) {
                    console.error('❌ Virtual DOM update error:', error);
                }
            }
            
            // Schedule next chunk if more items remain
            if (this.queue.length > 0) {
                virtualDom.rafId = requestAnimationFrame(processChunk);
            } else {
                this.processing = false;
                virtualDom.isRendering = false;
            }
        };
        
        virtualDom.rafId = requestAnimationFrame(processChunk);
    }
    
    clear() {
        if (virtualDom.rafId) {
            cancelAnimationFrame(virtualDom.rafId);
            virtualDom.rafId = null;
        }
        this.queue = [];
        this.processing = false;
    }
}

// GLOBAL UPDATE QUEUE INSTANCE
const updateQueue = new ResultUpdateQueue();

// VIRTUAL DOM UPDATE FUNCTIONS
class VirtualDOMUpdate {
    constructor(type, data, callback) {
        this.type = type;
        this.data = data;
        this.callback = callback;
        this.timestamp = Date.now();
    }
    
    execute() {
        switch (this.type) {
            case 'ADD_RESULT':
                this.addResultToDOM();
                break;
            case 'UPDATE_COUNT':
                this.updateResultCount();
                break;
            case 'SHOW_TOAST':
                this.showToastMessage();
                break;
            case 'UPDATE_PROGRESS':
                this.updateProgressDisplay();
                break;
            default:
                console.warn('Unknown virtual DOM update type:', this.type);
        }
        
        if (this.callback) {
            this.callback();
        }
    }
    
    addResultToDOM() {
        try {
            const { result, shouldShow } = this.data;
            const resultsContainer = document.getElementById('results-container');
            
            if (!resultsContainer) {
                console.error('❌ Virtual DOM: Results container not found');
                return;
            }
            
            // Ensure UI is initialized
            this.ensureResultsUIInitialized(resultsContainer);
            
            // Create result element based on current view
            if (window.resultsModule && window.resultsModule.currentView) {
                const currentView = window.resultsModule.currentView || 'card';
                
                if (currentView === 'table') {
                    this.addResultToTableVirtual(result, shouldShow);
                } else if (currentView === 'hierarchy') {
                    this.addResultToHierarchyVirtual(result, shouldShow);
                } else {
                    this.addResultToCardsVirtual(result, shouldShow);
                }
            } else {
                this.addResultToCardsVirtual(result, shouldShow);
            }
            
            console.log(`✅ Virtual DOM: Added ${result.itemName} (${shouldShow ? 'visible' : 'hidden'})`);
            
        } catch (error) {
            console.error('❌ Virtual DOM add result error:', error);
        }
    }
    
    addResultToCardsVirtual(result, shouldShow) {
        const resultsList = document.getElementById('results-list');
        if (!resultsList) return;
        
        // Create virtual DOM structure for result card
        const virtualCard = this.createVirtualResultCard(result, shouldShow);
        const cardElement = virtualCard.render();
        
        if (cardElement) {
            resultsList.appendChild(cardElement);
            
            // Smooth scroll to new item
            if (shouldShow) {
                requestAnimationFrame(() => {
                    resultsList.scrollTop = resultsList.scrollHeight;
                });
            }
        }
    }
    
    addResultToTableVirtual(result, shouldShow) {
        const tableBody = document.querySelector('#results-table tbody');
        if (!tableBody) return;
        
        if (shouldShow) {
            const configModule = window.configModule;
            if (configModule) {
                const resultIndex = configModule.results.length - 1;
                
                // Create virtual table row
                const virtualRow = this.createVirtualTableRow(result, resultIndex);
                const rowElement = virtualRow.render();
                
                if (rowElement) {
                    tableBody.appendChild(rowElement);
                    
                    // Smooth scroll to new row
                    requestAnimationFrame(() => {
                        const tableContainer = document.getElementById('results-table-container');
                        if (tableContainer) {
                            tableContainer.scrollTop = tableContainer.scrollHeight;
                        }
                    });
                }
            }
        }
    }
    
    addResultToHierarchyVirtual(result, shouldShow) {
        if (!shouldShow) return;
        
        const hierarchyContainer = document.getElementById('results-hierarchy-container');
        if (!hierarchyContainer) return;
        
        // Use existing hierarchy logic but with virtual DOM batching
        try {
            const configModule = window.configModule;
            if (configModule) {
                const currentFilter = window.resultsModule ? window.resultsModule.getCurrentResultsFilter() : 'all';
                
                // Build hierarchy incrementally with virtual DOM
                this.addToHierarchyVirtual(result, hierarchyContainer, currentFilter);
            }
        } catch (error) {
            console.error('❌ Virtual DOM hierarchy error:', error);
        }
    }
    
    createVirtualResultCard(result, shouldShow) {
        const configModule = window.configModule;
        if (!configModule) return new VirtualDOMNode('div');
        
        // Create virtual card structure
        const cardStyle = {
            border: '1px solid var(--border)',
            borderRadius: '8px',
            padding: '16px',
            marginBottom: '12px',
            background: 'white',
            animation: 'fadeIn 0.3s ease-in',
            display: shouldShow ? 'block' : 'none'
        };
        
        // Header section
        const headerStyle = {
            display: 'grid',
            gridTemplateColumns: '1fr auto auto',
            gap: '16px',
            alignItems: 'flex-start',
            marginBottom: '16px'
        };
        
        // Title and path
        let displayPath = result.itemPath;
        if (!displayPath || displayPath === 'undefined') {
            displayPath = `/${result.itemName}`;
        }
        displayPath = displayPath.replace(/\/+/g, '/');
        if (!displayPath.startsWith('/')) displayPath = '/' + displayPath;
        
        const displayIcon = result.itemType === 'file' 
            ? (result.scanType === 'onedrive' ? '☁️📄' : '📄')
            : (result.scanType === 'onedrive' ? '☁️📁' : '📁');
        
        const titleContent = `${displayIcon} ${result.siteName}${displayPath}`;
        
        // Virtual DOM structure
        return new VirtualDOMNode('div', { style: cardStyle }, [
            new VirtualDOMNode('div', { style: headerStyle }, [
                this.createVirtualResultInfo(result, titleContent),
                this.createVirtualItemId(result),
                this.createVirtualOwners(result)
            ]),
            this.createVirtualPermissionsTable(result)
        ]);
    }
    
    createVirtualResultInfo(result, titleContent) {
        const infoStyle = { flex: '1' };
        const titleStyle = {
            margin: '0 0 4px 0',
            fontSize: '16px',
            color: 'var(--text)',
            fontWeight: '600'
        };
        const urlStyle = {
            margin: '0',
            fontSize: '12px',
            color: 'var(--text-muted)'
        };
        
        return new VirtualDOMNode('div', { style: infoStyle }, [
            new VirtualDOMNode('h3', { style: titleStyle }, [titleContent]),
            new VirtualDOMNode('p', { style: urlStyle }, [
                result.siteUrl || (result.scanType === 'onedrive' ? 'Personal OneDrive' : 'SharePoint Site')
            ])
        ]);
    }
    
    createVirtualItemId(result) {
        const idStyle = {
            display: 'flex',
            flexDirection: 'column',
            alignItems: 'flex-start',
            paddingRight: '32px',
            marginRight: '16px'
        };
        
        const labelStyle = {
            fontSize: '10px',
            color: 'var(--text-muted)',
            textTransform: 'uppercase',
            fontWeight: '600',
            marginBottom: '2px'
        };
        
        const valueStyle = {
            fontSize: '11px',
            color: 'var(--text-muted)',
            fontFamily: 'monospace',
            maxWidth: '200px',
            overflow: 'hidden',
            textOverflow: 'ellipsis',
            whiteSpace: 'nowrap'
        };
        
        return new VirtualDOMNode('div', { style: idStyle }, [
            new VirtualDOMNode('span', { style: labelStyle }, ['ItemID']),
            new VirtualDOMNode('span', { 
                style: valueStyle, 
                title: result.itemId 
            }, [result.itemId])
        ]);
    }
    
    createVirtualOwners(result) {
        const configModule = window.configModule;
        const resultsModule = window.resultsModule;
        
        let owners = 'n/a';
        if (resultsModule && resultsModule.extractOwnersFromResult) {
            owners = resultsModule.extractOwnersFromResult(result);
        } else if (configModule && configModule.extractOwnersFromResult) {
            owners = configModule.extractOwnersFromResult(result);
        }
        
        const ownersStyle = {
            display: 'flex',
            flexDirection: 'column',
            alignItems: 'flex-start',
            paddingRight: '16px',
            marginRight: '8px'
        };
        
        const labelStyle = {
            fontSize: '10px',
            color: 'var(--text-muted)',
            textTransform: 'uppercase',
            fontWeight: '600',
            marginBottom: '2px'
        };
        
        const valueStyle = {
            fontSize: '11px',
            color: 'var(--text)',
            maxWidth: '150px',
            overflow: 'hidden',
            textOverflow: 'ellipsis',
            whiteSpace: 'nowrap'
        };
        
        return new VirtualDOMNode('div', { style: ownersStyle }, [
            new VirtualDOMNode('span', { style: labelStyle }, ['Owners']),
            new VirtualDOMNode('span', { 
                style: valueStyle, 
                title: owners 
            }, [owners])
        ]);
    }
    
    createVirtualPermissionsTable(result) {
        const configModule = window.configModule;
        const resultsModule = window.resultsModule;
        
        if (!configModule) return new VirtualDOMNode('div');
        
        // Get filtered permissions
        const currentFilter = resultsModule ? resultsModule.getCurrentResultsFilter() : 'all';
        let permissions = result.permissions;
        
        if (resultsModule && resultsModule.getFilteredPermissions) {
            permissions = resultsModule.getFilteredPermissions(result.permissions, currentFilter);
        }
        
        // Create table structure
        const table = new VirtualDOMNode('table', { className: 'results-table' }, [
            new VirtualDOMNode('thead', {}, [
                new VirtualDOMNode('tr', {}, [
                    new VirtualDOMNode('th', {}, ['Who Has Access']),
                    new VirtualDOMNode('th', {}, ['Permission Level']),
                    new VirtualDOMNode('th', {}, ['Sharing Type']),
                    new VirtualDOMNode('th', {}, ['Link Expiration']),
                    new VirtualDOMNode('th', {}, ['Actions'])
                ])
            ]),
            new VirtualDOMNode('tbody', {}, 
                permissions.map(permission => this.createVirtualPermissionRow(permission, result))
            )
        ]);
        
        return table;
    }
    
    createVirtualPermissionRow(permission, result) {
        const configModule = window.configModule;
        if (!configModule) return new VirtualDOMNode('tr');
        
        const who = configModule.extractUserFromPermission(permission, configModule.tenantDomains);
        const roles = (permission.roles || []).join(', ') || 'Not specified';
        const classification = configModule.classifyPermission(permission, configModule.tenantDomains);
        const expiration = configModule.extractExpirationDate(permission);
        
        const badgeClass = classification === 'external' ? 'external-badge' : 'internal-badge';
        
        // Create action buttons
        const hasLink = permission.link;
        const actionButtons = [];
        
        if (hasLink) {
            actionButtons.push(this.createVirtualActionButton('EXP', 'user-action-exp', () => {
                if (window.permissionsModule?.showExpirationDialog) {
                    window.permissionsModule.showExpirationDialog(result, null, permission);
                }
            }));
            
            actionButtons.push(this.createVirtualActionButton('LINK', 'user-action-link', () => {
                if (confirm(`Remove sharing link for ${who}?`)) {
                    if (window.permissionsModule?.disableLinks) {
                        window.permissionsModule.disableLinks(result, null, permission);
                    }
                }
            }));
        }
        
        actionButtons.push(this.createVirtualActionButton('USER', 'user-action-user', () => {
            if (confirm(`Remove all permissions for ${who}?`)) {
                if (window.permissionsModule?.disableAllSharing) {
                    window.permissionsModule.disableAllSharing(result, null, permission);
                }
            }
        }));
        
        return new VirtualDOMNode('tr', {}, [
            new VirtualDOMNode('td', {}, [who]),
            new VirtualDOMNode('td', {}, [roles]),
            new VirtualDOMNode('td', {}, [
                new VirtualDOMNode('span', { className: badgeClass }, [classification.toUpperCase()])
            ]),
            new VirtualDOMNode('td', {}, [expiration]),
            new VirtualDOMNode('td', { style: { textAlign: 'center', padding: '8px' } }, [
                new VirtualDOMNode('div', { 
                    style: { 
                        display: 'flex', 
                        gap: '6px', 
                        justifyContent: 'center', 
                        flexWrap: 'wrap' 
                    } 
                }, actionButtons)
            ])
        ]);
    }
    
    createVirtualActionButton(text, className, onClick) {
        const buttonStyle = {
            background: className.includes('exp') ? 'var(--purple)' :
                       className.includes('link') ? 'var(--orange)' : 'var(--danger)',
            color: 'white',
            border: 'none',
            padding: '6px 10px',
            borderRadius: '4px',
            fontSize: '11px',
            fontWeight: '600',
            cursor: 'pointer',
            transition: 'all 0.2s ease',
            boxShadow: '0 2px 4px rgba(0,0,0,0.1)',
            minWidth: '45px'
        };
        
        return new VirtualDOMNode('button', {
            className: `user-action-btn ${className}`,
            style: buttonStyle,
            onclick: onClick
        }, [text]);
    }
    
    createVirtualTableRow(result, resultIndex) {
        const configModule = window.configModule;
        if (!configModule) return new VirtualDOMNode('tr');
        
        // Extract display information
        let displayPath = result.itemPath;
        if (!displayPath || displayPath === 'undefined') {
            displayPath = `/${result.itemName}`;
        }
        displayPath = displayPath.replace(/\/+/g, '/');
        if (!displayPath.startsWith('/')) displayPath = '/' + displayPath;
        
        const displayIcon = result.itemType === 'file' 
            ? (result.scanType === 'onedrive' ? '☁️📄' : '📄')
            : (result.scanType === 'onedrive' ? '☁️📁' : '📁');
        
        // Get owners
        let owners = 'n/a';
        if (window.resultsModule && window.resultsModule.extractOwnersFromResult) {
            owners = window.resultsModule.extractOwnersFromResult(result);
        }
        
        // Get filtered permissions
        const currentFilter = window.resultsModule ? window.resultsModule.getCurrentResultsFilter() : 'all';
        let permissions = result.permissions;
        if (window.resultsModule && window.resultsModule.getFilteredPermissions) {
            permissions = window.resultsModule.getFilteredPermissions(result.permissions, currentFilter);
        }
        
        return new VirtualDOMNode('tr', { 'data-result-index': resultIndex }, [
            // Item column
            new VirtualDOMNode('td', {}, [
                new VirtualDOMNode('div', { className: 'table-item-name' }, [
                    `${displayIcon} ${result.siteName}${displayPath}`
                ]),
                new VirtualDOMNode('div', { className: 'table-item-path' }, [
                    result.siteUrl || (result.scanType === 'onedrive' ? 'Personal OneDrive' : 'SharePoint Site')
                ])
            ]),
            // ItemID column
            new VirtualDOMNode('td', {}, [
                new VirtualDOMNode('div', { className: 'table-item-id' }, [result.itemId])
            ]),
            // Owners column
            new VirtualDOMNode('td', {}, [
                new VirtualDOMNode('div', { className: 'table-item-owners' }, [owners])
            ]),
            // Permissions column
            new VirtualDOMNode('td', { className: 'table-permissions-cell' }, 
                permissions.map(p => this.createVirtualTablePermissionItem(p))
            ),
            // Classifications column
            new VirtualDOMNode('td', {}, 
                permissions.map(p => {
                    const classification = configModule.classifyPermission(p, configModule.tenantDomains);
                    const badgeClass = classification === 'external' ? 'external-badge' : 'internal-badge';
                    return new VirtualDOMNode('span', { className: badgeClass }, [classification.toUpperCase()]);
                })
            ),
            // Actions column
            new VirtualDOMNode('td', { className: 'table-actions-cell' }, [
                this.createVirtualTableActions(result, resultIndex)
            ])
        ]);
    }
    
    createVirtualTablePermissionItem(permission) {
        const configModule = window.configModule;
        if (!configModule) return new VirtualDOMNode('div');
        
        const who = configModule.extractUserFromPermission(permission, configModule.tenantDomains);
        const roles = (permission.roles || []).join(', ') || 'Not specified';
        const exp = configModule.extractExpirationDate(permission);
        const roleText = exp !== 'No expiration' ? `${roles} • Expires: ${exp}` : roles;
        
        return new VirtualDOMNode('div', { className: 'table-permission-item' }, [
            new VirtualDOMNode('div', { className: 'table-permission-who' }, [who]),
            new VirtualDOMNode('div', { className: 'table-permission-role' }, [roleText])
        ]);
    }
    
    createVirtualTableActions(result, resultIndex) {
        const hasLinks = result.permissions.some(p => p.link);
        const buttons = [];
        
        if (hasLinks) {
            buttons.push(
                new VirtualDOMNode('button', {
                    className: 'table-action-btn purple',
                    onclick: () => window.handleTableAction('expiration', resultIndex)
                }, ['SET EXP']),
                new VirtualDOMNode('button', {
                    className: 'table-action-btn orange',
                    onclick: () => window.handleTableAction('disableLinks', resultIndex)
                }, ['DEL LINKS'])
            );
        }
        
        buttons.push(
            new VirtualDOMNode('button', {
                className: 'table-action-btn red',
                onclick: () => window.handleTableAction('disableAll', resultIndex)
            }, ['DEL ALL'])
        );
        
        return new VirtualDOMNode('div', {}, buttons);
    }
    
    addToHierarchyVirtual(result, hierarchyContainer, currentFilter) {
        // Simplified hierarchy update for virtual DOM
        // This would be expanded with full hierarchy logic similar to results.js
        const siteKey = `${result.scanType}_${result.siteName}`;
        
        // Find or create site container
        let siteContainer = hierarchyContainer.querySelector(`[data-site-key="${siteKey}"]`);
        if (!siteContainer) {
            const siteDisplayName = result.scanType === 'onedrive' ? 
                `☁️ ${result.siteName} (OneDrive)` : 
                `📁 ${result.siteName}`;
            
            const virtualSiteNode = this.createVirtualHierarchyNode(siteDisplayName, siteKey, 'site', null, 0);
            siteContainer = virtualSiteNode.render();
            siteContainer.setAttribute('data-site-key', siteKey);
            hierarchyContainer.appendChild(siteContainer);
        }
        
        // Add item to site (simplified - full hierarchy would parse path)
        const childrenContainer = siteContainer.querySelector('.tree-node-children');
        if (childrenContainer) {
            const virtualItemNode = this.createVirtualHierarchyNode(result.itemName, `${siteKey}_${result.itemId}`, result.itemType, result, 1);
            const itemElement = virtualItemNode.render();
            childrenContainer.appendChild(itemElement);
        }
    }
    
    createVirtualHierarchyNode(name, nodeKey, nodeType, result, level) {
        const nodeStyle = {
            marginLeft: `${level * 20}px`,
            marginBottom: '2px'
        };
        
        const headerStyle = {
            display: 'flex',
            alignItems: 'center',
            padding: '6px 12px',
            border: '1px solid var(--border)',
            borderRadius: '4px',
            background: 'white',
            cursor: 'pointer',
            marginBottom: '2px',
            transition: 'background-color 0.2s'
        };
        
        const isContainer = nodeType === 'site' || nodeType === 'folder';
        const expandIcon = isContainer ? '▼' : '•';
        
        const nodeIcon = nodeType === 'site' ? (name.includes('OneDrive') ? '☁️' : '🏢') :
                         nodeType === 'folder' ? '📁' : '📄';
        
        const children = [
            new VirtualDOMNode('div', { className: 'tree-node-header', style: headerStyle }, [
                new VirtualDOMNode('span', { className: 'expand-icon', style: { marginRight: '8px', fontSize: '12px', width: '16px' } }, [expandIcon]),
                new VirtualDOMNode('span', { className: 'tree-node-name', style: { flex: '1', fontSize: '14px' } }, [`${nodeIcon} ${name}`])
            ]),
            new VirtualDOMNode('div', { className: 'tree-node-children', style: { marginLeft: '0px', display: 'block' } })
        ];
        
        return new VirtualDOMNode('div', {
            className: `tree-node tree-node-${nodeType}`,
            style: nodeStyle,
            'data-node-key': nodeKey
        }, children);
    }
    
    ensureResultsUIInitialized(resultsContainer) {
        // Ensure UI components exist
        let resultsList = document.getElementById('results-list');
        if (!resultsList) {
            resultsContainer.innerHTML = '<div style="max-height: 400px; overflow-y: auto; border: 1px solid var(--border); border-radius: 8px; background: white;" id="results-list"></div>';
            
            // Show UI controls
            const actionsEl = document.getElementById('results-actions');
            const filtersEl = document.getElementById('sharing-filters');
            const bulkEl = document.getElementById('bulk-controls');
            
            if (actionsEl) actionsEl.style.display = 'block';
            if (filtersEl && filtersEl.parentElement) filtersEl.parentElement.style.display = 'flex';
            if (bulkEl) bulkEl.style.display = 'flex';
        }
    }
    
    updateResultCount() {
        const { count, filter } = this.data;
        const resultCount = document.getElementById('result-count');
        
        if (resultCount) {
            if (filter === 'all') {
                resultCount.innerText = `${count} found`;
            } else {
                resultCount.innerText = `${count} results (${filter} filter)`;
            }
            
            if (count > 0) {
                resultCount.className = 'status-badge status-approved';
                const exportBtn = document.getElementById('export-btn');
                if (exportBtn) {
                    exportBtn.disabled = false;
                }
            }
        }
    }
    
    showToastMessage() {
        const { message } = this.data;
        const configModule = window.configModule;
        if (configModule && configModule.showToast) {
            configModule.showToast(message);
        }
    }
    
    updateProgressDisplay() {
        const { progressId, text } = this.data;
        const progressElement = document.getElementById(progressId);
        if (progressElement && text) {
            progressElement.innerText = text;
        }
    }
}

// VIRTUAL DOM API FUNCTIONS
function enableVirtualDOM() {
    virtualDom.enabled = true;
    console.log('✅ Virtual DOM enabled for performance optimization');
}

function disableVirtualDOM() {
    virtualDom.enabled = false;
    updateQueue.clear();
    console.log('🔧 Virtual DOM disabled - switching to direct DOM updates');
}

function isVirtualDOMEnabled() {
    return virtualDom.enabled;
}

// Add result to virtual DOM queue
function addResult(result, shouldShow = true) {
    if (!virtualDom.enabled) {
        // Fall back to direct DOM updates when virtual DOM is disabled
        if (window.resultsModule && window.resultsModule.addResultToDisplay) {
            window.resultsModule.addResultToDisplay(result);
        }
        return;
    }
    
    const update = new VirtualDOMUpdate('ADD_RESULT', { result, shouldShow });
    updateQueue.add(update);
}

// Update result count through virtual DOM
function updateCount(count, filter = 'all') {
    if (!virtualDom.enabled) {
        // Direct update when virtual DOM disabled
        const resultCount = document.getElementById('result-count');
        if (resultCount) {
            if (filter === 'all') {
                resultCount.innerText = `${count} found`;
            } else {
                resultCount.innerText = `${count} results (${filter} filter)`;
            }
        }
        return;
    }
    
    const update = new VirtualDOMUpdate('UPDATE_COUNT', { count, filter });
    updateQueue.add(update);
}

// Show toast message through virtual DOM
function showToast(message) {
    if (!virtualDom.enabled) {
        // Direct update when virtual DOM disabled
        const configModule = window.configModule;
        if (configModule && configModule.showToast) {
            configModule.showToast(message);
        }
        return;
    }
    
    const update = new VirtualDOMUpdate('SHOW_TOAST', { message });
    updateQueue.add(update);
}

// Update progress display through virtual DOM
function updateProgress(progressId, text) {
    if (!virtualDom.enabled) {
        // Direct update when virtual DOM disabled
        const progressElement = document.getElementById(progressId);
        if (progressElement) {
            progressElement.innerText = text;
        }
        return;
    }
    
    const update = new VirtualDOMUpdate('UPDATE_PROGRESS', { progressId, text });
    updateQueue.add(update);
}

// Clear virtual DOM queue
function clearQueue() {
    updateQueue.clear();
    console.log('🔧 Virtual DOM queue cleared');
}

// Get queue statistics
function getQueueStats() {
    return {
        queueLength: updateQueue.queue.length,
        processing: updateQueue.processing,
        enabled: virtualDom.enabled,
        batchSize: virtualDom.batchSize
    };
}

// Set virtual DOM batch size for performance tuning
function setBatchSize(size) {
    virtualDom.batchSize = Math.max(1, Math.min(20, size));
    console.log(`🔧 Virtual DOM batch size set to ${virtualDom.batchSize}`);
}

// SCANNING INTEGRATION FUNCTIONS
function handleScanningWithVirtualDOM(csvExportActive) {
    if (csvExportActive) {
        console.log('🚀 VIRTUAL DOM: Enabling virtual DOM for CSV export performance optimization');
        enableVirtualDOM();
        
        // Show performance mode notification
        showToast('Performance mode active - real-time results will continue during CSV export');
        
        // Adjust batch size for optimal performance during CSV export
        setBatchSize(3);
        
        return true;
    } else {
        console.log('🔧 VIRTUAL DOM: CSV export disabled - switching to direct DOM updates');
        disableVirtualDOM();
        
        // Reset batch size to default
        setBatchSize(5);
        
        return false;
    }
}

// Replace existing result display function with virtual DOM version
function createVirtualResultDisplay(result, shouldShow = true) {
    if (virtualDom.enabled) {
        addResult(result, shouldShow);
    } else {
        // Fall back to existing direct DOM update
        if (window.resultsModule && window.resultsModule.addResultToDisplay) {
            window.resultsModule.addResultToDisplay(result);
        }
    }
}

// INITIALIZATION AND CLEANUP
function initializeVirtualDOM() {
    console.log('🎯 Initializing Virtual DOM module...');
    
    // Set initial state
    virtualDom.enabled = false;
    virtualDom.batchSize = 5;
    virtualDom.isRendering = false;
    
    // Clear any existing queues
    updateQueue.clear();
    
    console.log('✅ Virtual DOM module initialized successfully');
    
    return {
        enableVirtualDOM,
        disableVirtualDOM,
        isVirtualDOMEnabled,
        addResult,
        updateCount,
        showToast,
        updateProgress,
        clearQueue,
        getQueueStats,
        setBatchSize,
        handleScanningWithVirtualDOM,
        createVirtualResultDisplay
    };
}

// Cleanup function for page unload
function cleanupVirtualDOM() {
    if (virtualDom.rafId) {
        cancelAnimationFrame(virtualDom.rafId);
        virtualDom.rafId = null;
    }
    
    updateQueue.clear();
    virtualDom.enabled = false;
    virtualDom.isRendering = false;
    
    console.log('🧹 Virtual DOM cleaned up');
}

// Add cleanup listener
if (typeof window !== 'undefined') {
    window.addEventListener('beforeunload', cleanupVirtualDOM);
}

// Export functions for use in other modules
window.virtualDomModule = {
    // Core functions
    enableVirtualDOM,
    disableVirtualDOM,
    isVirtualDOMEnabled,
    
    // Update functions
    addResult,
    updateCount,
    showToast,
    updateProgress,
    
    // Queue management
    clearQueue,
    getQueueStats,
    setBatchSize,
    
    // Integration functions
    handleScanningWithVirtualDOM,
    createVirtualResultDisplay,
    
    // Lifecycle functions
    initializeVirtualDOM,
    cleanupVirtualDOM,
    
    // Classes for advanced usage
    VirtualDOMNode,
    VirtualDOMUpdate,
    
    // State getters
    get enabled() { return virtualDom.enabled; },
    get queueLength() { return updateQueue.queue.length; },
    get processing() { return updateQueue.processing; }
};

console.log('📦 Virtual DOM module loaded and ready');
