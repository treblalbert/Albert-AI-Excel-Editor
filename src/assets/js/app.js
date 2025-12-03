let apiKey = '';
let apiKeyPath = '';
let workbookData = [];
let headers = [];
let history = [];
let columnWidths = {};
let activeFilters = {};
let cellColors = {};
let originalFileName = '';
let modifiedCells = new Set();
let highlightEnabled = false;
let lastAIChange = null;

// DOM Elements
const apiSetup = document.getElementById('apiSetup');
const mainApp = document.getElementById('mainApp');
const apiKeyFile = document.getElementById('apiKeyFile');
const excelFile = document.getElementById('excelFile');
const aiPrompt = document.getElementById('aiPrompt');
const aiButton = document.getElementById('aiButton');
const undoBtn = document.getElementById('undoBtn');
const addRowBtn = document.getElementById('addRowBtn');
const addColBtn = document.getElementById('addColBtn');
const compareBtn = document.getElementById('compareBtn');
const downloadBtn = document.getElementById('downloadBtn');
const statusDiv = document.getElementById('status');
const tableHead = document.getElementById('tableHead');
const tableBody = document.getElementById('tableBody');
const chatMessages = document.getElementById('chatMessages');
const highlightToggle = document.getElementById('highlightToggle');
const compareModal = document.getElementById('compareModal');
const closeCompareModal = document.getElementById('closeCompareModal');

// Load saved preferences on startup
window.addEventListener('load', () => {
    const savedPath = localStorage.getItem('apiKeyPath');
    const savedHighlight = localStorage.getItem('highlightEnabled');
    
    if (savedHighlight === 'true') {
        highlightToggle.checked = true;
        highlightEnabled = true;
    }
    
    if (savedPath) {
        addChatMessage('system', t('attemptingLoad'));
        addChatMessage('system', t('previouslyUsed', { path: savedPath }));
    }
});

// Highlight toggle handler
highlightToggle.addEventListener('change', (e) => {
    highlightEnabled = e.target.checked;
    localStorage.setItem('highlightEnabled', highlightEnabled);
    renderTable();
});

function showStatus(message, type = 'info') {
    statusDiv.className = `status ${type}`;
    statusDiv.textContent = message;
    statusDiv.style.display = 'block';
    setTimeout(() => statusDiv.style.display = 'none', 5000);
}

function addChatMessage(type, content) {
    const msgDiv = document.createElement('div');
    msgDiv.className = `message ${type}`;
    msgDiv.textContent = content;
    chatMessages.appendChild(msgDiv);
    chatMessages.scrollTop = chatMessages.scrollHeight;
}

apiKeyFile.addEventListener('change', async (e) => {
    const file = e.target.files[0];
    if (!file) return;

    const text = await file.text();
    apiKey = text.trim();
    apiKeyPath = file.name;
    
    localStorage.setItem('apiKeyPath', file.name);
    
    addChatMessage('system', t('apiKeyLoaded'));
    apiSetup.classList.add('hidden');
    mainApp.classList.remove('hidden');
    updateButtonStates();
});

excelFile.addEventListener('change', async (e) => {
    const file = e.target.files[0];
    if (!file) return;

    try {
        // Store original filename (without extension for flexibility)
        originalFileName = file.name.replace(/\.[^/.]+$/, '');
        
        const data = await file.arrayBuffer();
        const workbook = XLSX.read(data, { type: 'array', cellStyles: true });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        
        // Extract cell colors
        cellColors = {};
        for (let cell in worksheet) {
            if (cell[0] === '!') continue;
            if (worksheet[cell].s && worksheet[cell].s.fgColor) {
                cellColors[cell] = '#' + worksheet[cell].s.fgColor.rgb;
            }
        }

        const json = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });

        if (json.length === 0) {
            showStatus(t('errorEmpty'), 'error');
            return;
        }

        headers = json[0].map(h => h !== undefined && h !== null ? String(h) : '');
        workbookData = json.slice(1).map(row => 
            row.map(cell => cell !== undefined && cell !== null ? cell : '')
        );
        history = [];
        activeFilters = {};
        modifiedCells.clear();
        lastAIChange = null;
        
        renderTable();
        addChatMessage('system', t('fileLoaded', { 
            rows: workbookData.length, 
            cols: headers.length,
            filename: file.name 
        }));
        updateButtonStates();
    } catch (error) {
        showStatus(t('errorLoading', { error: error.message }), 'error');
        addChatMessage('system', t('apiError', { error: error.message }));
    }
});

function saveState() {
    history.push({
        headers: JSON.parse(JSON.stringify(headers)),
        data: JSON.parse(JSON.stringify(workbookData)),
        colors: JSON.parse(JSON.stringify(cellColors)),
        modifiedCells: new Set(modifiedCells)
    });
    if (history.length > 20) history.shift();
    updateButtonStates();
}

function renderTable() {
    tableHead.innerHTML = '';
    tableBody.innerHTML = '';

    const headerRow = document.createElement('tr');
    const cornerCell = document.createElement('th');
    cornerCell.className = 'row-header';
    cornerCell.textContent = '#';
    headerRow.appendChild(cornerCell);

    headers.forEach((header, i) => {
        const th = document.createElement('th');
        th.style.width = columnWidths[i] || '150px';
        th.style.position = 'relative';
        
        const headerContent = document.createElement('div');
        headerContent.className = 'header-content';
        
        const headerText = document.createElement('span');
        headerText.className = 'header-text';
        headerText.textContent = header || `Column ${i + 1}`;
        headerContent.appendChild(headerText);
        
        const filterIcon = document.createElement('span');
        filterIcon.className = 'filter-icon' + (activeFilters[i] && activeFilters[i].size > 0 ? ' active' : '');
        filterIcon.textContent = '▼';
        filterIcon.onclick = (e) => {
            e.stopPropagation();
            showFilterDropdown(e, i);
        };
        headerContent.appendChild(filterIcon);
        
        th.appendChild(headerContent);

        const resizeHandle = document.createElement('div');
        resizeHandle.className = 'resize-handle';
        resizeHandle.addEventListener('mousedown', (e) => startResize(e, i, th));
        th.appendChild(resizeHandle);

        headerRow.appendChild(th);
    });
    tableHead.appendChild(headerRow);

    const filteredData = applyFilters();
    filteredData.forEach((row) => {
        const tr = document.createElement('tr');
        
        const rowHeader = document.createElement('td');
        rowHeader.className = 'row-header';
        rowHeader.textContent = row.originalIndex + 1;
        tr.appendChild(rowHeader);

        headers.forEach((_, colIndex) => {
            const td = document.createElement('td');
            const cellRef = XLSX.utils.encode_cell({r: row.originalIndex + 1, c: colIndex});
            
            // Apply original cell color if exists
            if (cellColors[cellRef]) {
                td.style.backgroundColor = cellColors[cellRef];
            }
            
            // Apply AI modification highlight if enabled
            const cellKey = `${row.originalIndex}-${colIndex}`;
            if (highlightEnabled && modifiedCells.has(cellKey)) {
                td.classList.add('ai-modified');
            }

            const input = document.createElement('input');
            input.type = 'text';
            input.value = row.data[colIndex] !== undefined ? row.data[colIndex] : '';
            input.addEventListener('change', (e) => {
                workbookData[row.originalIndex][colIndex] = e.target.value;
            });
            td.appendChild(input);
            tr.appendChild(td);
        });
        tableBody.appendChild(tr);
    });
}

function applyFilters() {
    let filtered = workbookData.map((data, idx) => ({data, originalIndex: idx}));
    
    for (let colIndex in activeFilters) {
        const filterValues = activeFilters[colIndex];
        if (filterValues && filterValues.size > 0) {
            filtered = filtered.filter(row => {
                const cellValue = String(row.data[colIndex] !== undefined ? row.data[colIndex] : '');
                return filterValues.has(cellValue);
            });
        }
    }
    
    return filtered;
}

function showFilterDropdown(e, colIndex) {
    e.stopPropagation();
    
    // Remove existing dropdowns
    document.querySelectorAll('.filter-dropdown').forEach(d => d.remove());

    const dropdown = document.createElement('div');
    dropdown.className = 'filter-dropdown';

    const searchInput = document.createElement('input');
    searchInput.type = 'text';
    searchInput.placeholder = 'Search...';
    dropdown.appendChild(searchInput);

    // Action buttons
    const actionsDiv = document.createElement('div');
    actionsDiv.className = 'filter-actions';
    
    const selectAllBtn = document.createElement('button');
    selectAllBtn.textContent = t('selectAll');
    selectAllBtn.onclick = () => {
        delete activeFilters[colIndex];
        optionsDiv.querySelectorAll('input[type="checkbox"]').forEach(cb => cb.checked = true);
        renderTable();
    };
    actionsDiv.appendChild(selectAllBtn);
    
    const clearBtn = document.createElement('button');
    clearBtn.textContent = t('clearAll');
    clearBtn.onclick = () => {
        activeFilters[colIndex] = new Set();
        optionsDiv.querySelectorAll('input[type="checkbox"]').forEach(cb => cb.checked = false);
        renderTable();
    };
    actionsDiv.appendChild(clearBtn);
    
    dropdown.appendChild(actionsDiv);

    const optionsDiv = document.createElement('div');
    optionsDiv.className = 'filter-options';

    // Get unique values from the column
    const uniqueValues = [...new Set(workbookData.map(row => {
        const val = row[colIndex];
        return val !== undefined && val !== null ? String(val) : '';
    }))].sort((a, b) => {
        // Try numeric sort first
        const numA = parseFloat(a);
        const numB = parseFloat(b);
        if (!isNaN(numA) && !isNaN(numB)) return numA - numB;
        return a.localeCompare(b);
    });
    
    const currentFilter = activeFilters[colIndex];
    const hasFilter = currentFilter && currentFilter.size > 0;

    uniqueValues.forEach(value => {
        const option = document.createElement('div');
        option.className = 'filter-option';
        
        const checkbox = document.createElement('input');
        checkbox.type = 'checkbox';
        checkbox.checked = !hasFilter || currentFilter.has(value);
        
        checkbox.onchange = () => {
            if (!activeFilters[colIndex]) {
                // First filter on this column - start with all selected except this one
                activeFilters[colIndex] = new Set(uniqueValues);
            }
            
            if (checkbox.checked) {
                activeFilters[colIndex].add(value);
            } else {
                activeFilters[colIndex].delete(value);
            }
            
            // If all are selected, remove the filter
            if (activeFilters[colIndex].size === uniqueValues.length) {
                delete activeFilters[colIndex];
            }
            
            renderTable();
        };
        
        const label = document.createElement('span');
        label.textContent = value === '' ? '(blank)' : value;
        
        option.appendChild(checkbox);
        option.appendChild(label);
        optionsDiv.appendChild(option);
    });

    dropdown.appendChild(optionsDiv);
    
    // Position the dropdown
    const th = e.target.closest('th');
    th.appendChild(dropdown);

    // Filter search functionality
    searchInput.oninput = () => {
        const search = searchInput.value.toLowerCase();
        optionsDiv.querySelectorAll('.filter-option').forEach(opt => {
            const text = opt.textContent.toLowerCase();
            opt.style.display = text.includes(search) ? 'flex' : 'none';
        });
    };

    // Close dropdown when clicking outside
    const closeHandler = (e) => {
        if (!dropdown.contains(e.target) && e.target !== dropdown) {
            dropdown.remove();
            document.removeEventListener('click', closeHandler);
        }
    };
    
    setTimeout(() => {
        document.addEventListener('click', closeHandler);
    }, 0);
    
    searchInput.focus();
}

function startResize(e, colIndex, th) {
    e.preventDefault();
    e.stopPropagation();
    
    const startX = e.pageX;
    const startWidth = th.offsetWidth;
    const handle = e.target;
    
    handle.classList.add('resizing');
    document.body.style.cursor = 'col-resize';
    document.body.style.userSelect = 'none';

    function onMouseMove(e) {
        const diff = e.pageX - startX;
        const newWidth = Math.max(60, startWidth + diff);
        columnWidths[colIndex] = newWidth + 'px';
        th.style.width = columnWidths[colIndex];
        
        // Also update the corresponding column cells
        const rows = tableBody.querySelectorAll('tr');
        rows.forEach(row => {
            const cell = row.children[colIndex + 1]; // +1 for row header
            if (cell) {
                cell.style.width = columnWidths[colIndex];
            }
        });
    }

    function onMouseUp() {
        handle.classList.remove('resizing');
        document.body.style.cursor = '';
        document.body.style.userSelect = '';
        document.removeEventListener('mousemove', onMouseMove);
        document.removeEventListener('mouseup', onMouseUp);
    }

    document.addEventListener('mousemove', onMouseMove);
    document.addEventListener('mouseup', onMouseUp);
}

function updateButtonStates() {
    const hasData = workbookData.length > 0;
    const hasHistory = history.length > 0;
    
    aiButton.disabled = !hasData;
    aiPrompt.disabled = !hasData;
    undoBtn.disabled = !hasHistory;
    addRowBtn.disabled = !hasData;
    addColBtn.disabled = !hasData;
    downloadBtn.disabled = !hasData;
    compareBtn.disabled = !hasHistory;
}

// Detect changes between two data sets
function detectChanges(oldHeaders, oldData, newHeaders, newData) {
    const changes = [];
    const modified = new Set();
    
    // Check for header changes
    const maxHeaderLen = Math.max(oldHeaders.length, newHeaders.length);
    for (let i = 0; i < maxHeaderLen; i++) {
        if (oldHeaders[i] !== newHeaders[i]) {
            if (i >= oldHeaders.length) {
                changes.push(`Added column "${newHeaders[i]}"`);
            } else if (i >= newHeaders.length) {
                changes.push(`Removed column "${oldHeaders[i]}"`);
            } else {
                changes.push(`Renamed column "${oldHeaders[i]}" to "${newHeaders[i]}"`);
            }
        }
    }
    
    // Check for row count changes
    if (newData.length > oldData.length) {
        changes.push(`Added ${newData.length - oldData.length} row(s)`);
    } else if (newData.length < oldData.length) {
        changes.push(`Removed ${oldData.length - newData.length} row(s)`);
    }
    
    // Check for cell value changes
    let cellChanges = 0;
    const minRows = Math.min(oldData.length, newData.length);
    const minCols = Math.min(oldHeaders.length, newHeaders.length);
    
    for (let r = 0; r < minRows; r++) {
        for (let c = 0; c < minCols; c++) {
            const oldVal = oldData[r] && oldData[r][c] !== undefined ? String(oldData[r][c]) : '';
            const newVal = newData[r] && newData[r][c] !== undefined ? String(newData[r][c]) : '';
            if (oldVal !== newVal) {
                cellChanges++;
                modified.add(`${r}-${c}`);
            }
        }
    }
    
    // Track new cells in added rows/columns
    for (let r = 0; r < newData.length; r++) {
        for (let c = 0; c < newHeaders.length; c++) {
            if (r >= oldData.length || c >= oldHeaders.length) {
                modified.add(`${r}-${c}`);
            }
        }
    }
    
    if (cellChanges > 0) {
        changes.push(`Modified ${cellChanges} cell(s)`);
    }
    
    return { changes, modified };
}

// ============================================
// SAFETY MECHANISM: Protect against row deletion
// ============================================

// Check if user explicitly asked to delete rows
function userAskedToDeleteRows(prompt) {
    const deleteKeywords = [
        'delete', 'remove', 'erase', 'drop', 'clear', 'wipe', 'get rid', 'throw away',
        'eliminar', 'borrar', 'quitar', 'elimina', 'borra', 'suprimir',
        'esborrar', 'treure', 'elimina', 'esborra', 'suprimeix'
    ];
    const rowKeywords = [
        'row', 'rows', 'line', 'lines', 'record', 'records', 'entry', 'entries', 'item', 'items',
        'fila', 'filas', 'línea', 'líneas', 'registro', 'registros', 'entrada', 'entradas',
        'fila', 'files', 'línia', 'línies', 'registre', 'registres'
    ];
    
    const lowerPrompt = prompt.toLowerCase();
    
    const hasDeleteKeyword = deleteKeywords.some(kw => lowerPrompt.includes(kw));
    const hasRowKeyword = rowKeywords.some(kw => lowerPrompt.includes(kw));
    
    return hasDeleteKeyword && hasRowKeyword;
}

// BULLETPROOF safety function: NEVER lose rows unless explicitly asked
function mergeAIChangesIntoOriginal(oldData, newData, oldHeaders, newHeaders, userPrompt) {
    
    // ONLY allow row deletion if user EXPLICITLY asked for it
    if (userAskedToDeleteRows(userPrompt)) {
        console.log('User requested row deletion - allowing AI to delete rows');
        return { 
            data: newData, 
            headers: newHeaders, 
            restored: 0,
            message: null 
        };
    }
    
    // If AI returned same or MORE rows, that's fine - use AI result
    if (newData.length >= oldData.length) {
        console.log(`AI returned ${newData.length} rows (original: ${oldData.length}) - OK`);
        return { 
            data: newData, 
            headers: newHeaders, 
            restored: 0,
            message: null 
        };
    }
    
    // ============================================
    // ⚠️ AI DELETED ROWS WITHOUT PERMISSION!
    // FORCE keep all original rows, only apply cell changes
    // ============================================
    
    console.warn(`🚨 AI tried to delete rows! Original: ${oldData.length}, AI returned: ${newData.length}`);
    console.warn('Forcing preservation of all original rows...');
    
    const finalHeaders = newHeaders.length >= oldHeaders.length ? [...newHeaders] : [...oldHeaders];
    const numCols = finalHeaders.length;
    
    // ALWAYS start with ALL original rows - deep copy
    const finalData = [];
    for (let i = 0; i < oldData.length; i++) {
        const row = [];
        for (let j = 0; j < numCols; j++) {
            row.push(oldData[i] && oldData[i][j] !== undefined ? oldData[i][j] : '');
        }
        finalData.push(row);
    }
    
    // Now try to apply changes from AI response
    // Strategy: For each row AI returned, find the best matching original row and apply changes
    
    for (let aiRowIdx = 0; aiRowIdx < newData.length; aiRowIdx++) {
        const aiRow = newData[aiRowIdx];
        
        // Find which original row this AI row most likely corresponds to
        let bestMatchIdx = -1;
        let bestMatchScore = -1;
        
        for (let origIdx = 0; origIdx < oldData.length; origIdx++) {
            const origRow = oldData[origIdx];
            let score = 0;
            
            // Count matching cells
            for (let col = 0; col < Math.min(origRow.length, aiRow.length); col++) {
                const origVal = String(origRow[col] ?? '');
                const aiVal = String(aiRow[col] ?? '');
                if (origVal === aiVal) {
                    score += 10; // Exact match
                } else if (origVal !== '' && aiVal !== '' && 
                          (origVal.includes(aiVal) || aiVal.includes(origVal))) {
                    score += 3; // Partial match
                }
            }
            
            // Bonus for same position
            if (origIdx === aiRowIdx) score += 20;
            else if (Math.abs(origIdx - aiRowIdx) === 1) score += 10;
            else if (Math.abs(origIdx - aiRowIdx) <= 3) score += 5;
            
            if (score > bestMatchScore) {
                bestMatchScore = score;
                bestMatchIdx = origIdx;
            }
        }
        
        // Apply changes from AI row to the matched original row
        if (bestMatchIdx !== -1 && bestMatchScore > 0) {
            for (let col = 0; col < aiRow.length; col++) {
                const origVal = String(oldData[bestMatchIdx]?.[col] ?? '');
                const aiVal = String(aiRow[col] ?? '');
                
                // If AI changed this cell, apply the change
                if (origVal !== aiVal) {
                    finalData[bestMatchIdx][col] = aiRow[col];
                }
            }
        }
    }
    
    const protectedCount = oldData.length - newData.length;
    
    return {
        data: finalData,
        headers: finalHeaders,
        restored: protectedCount,
        message: `⚠️ PROTECTED ${protectedCount} row(s) from being deleted! Your edits were applied to the correct rows.`
    };
}

aiButton.addEventListener('click', async () => {
    const prompt = aiPrompt.value.trim();
    if (!prompt) {
        showStatus(t('enterCommand'), 'error');
        return;
    }

    saveState();
    addChatMessage('user', prompt);
    aiPrompt.value = '';
    aiButton.disabled = true;

    addChatMessage('ai', t('processing'));

    try {
        const response = await fetch('https://api.openai.com/v1/chat/completions', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'Authorization': `Bearer ${apiKey}`
            },
            body: JSON.stringify({
                model: 'gpt-4o-mini',
                messages: [
                    {
                        role: 'system',
                        content: `You are a helpful assistant that can answer questions about spreadsheet data AND modify the data when requested.

IMPORTANT: You must determine if the user is:
1. ASKING A QUESTION about the data (e.g., "what's the total?", "how many rows?", "which item has the highest value?")
2. REQUESTING A MODIFICATION (e.g., "add 10%", "sort by", "delete rows where", "change X to Y")

For QUESTIONS: Return JSON with format:
{"type": "question", "answer": "Your detailed answer here based on the data"}

For MODIFICATIONS: Return JSON with format:
{"type": "modification", "headers": [...], "data": [[...], [...]], "description": "Brief description of what was changed"}

CRITICAL RULES FOR MODIFICATIONS:
- You MUST return ALL rows from the original data, not just the modified ones
- If the user asks to modify rows 1-3, you must still include rows 4, 5, 6... and ALL other rows unchanged
- NEVER delete or omit rows unless the user explicitly asks to delete them
- The "data" array must contain EVERY row from the original, with only the requested changes applied
- Preserve all original values that were not explicitly asked to be changed
- The number of rows in your response should match the original unless user asked to add/remove rows

Always analyze the data carefully and provide accurate responses. For calculations, show your work when relevant.`
                    },
                    {
                        role: 'user',
                        content: `Spreadsheet headers: ${JSON.stringify(headers)}
Spreadsheet data (${workbookData.length} rows): ${JSON.stringify(workbookData)}

User request: ${prompt}`
                    }
                ],
                temperature: 0.3,
                max_tokens: 4000
            })
        });

        if (!response.ok) {
            throw new Error(`API error: ${response.status}`);
        }

        const result = await response.json();
        const content = result.choices[0].message.content.trim();
        
        // Remove the "processing" message
        const messages = chatMessages.querySelectorAll('.message.ai');
        const lastAiMsg = messages[messages.length - 1];
        if (lastAiMsg && lastAiMsg.textContent === t('processing')) {
            lastAiMsg.remove();
        }
        
        // Parse response
        const jsonMatch = content.match(/\{[\s\S]*\}/);
        if (!jsonMatch) throw new Error('Invalid response format');
        
        const parsed = JSON.parse(jsonMatch[0]);
        
        if (parsed.type === 'question') {
            // It's a question - just show the answer
            addChatMessage('ai', parsed.answer);
            // Remove the saved state since no changes were made
            history.pop();
        } else if (parsed.type === 'modification' && parsed.headers && parsed.data) {
            // It's a modification
            const oldHeaders = [...headers];
            const oldData = workbookData.map(row => [...row]);
            
            // ⚠️ SAFETY CHECK: Merge AI changes without losing rows
            const mergeResult = mergeAIChangesIntoOriginal(
                oldData, 
                parsed.data, 
                oldHeaders, 
                parsed.headers, 
                prompt
            );
            
            headers = mergeResult.headers;
            workbookData = mergeResult.data;
            
            // 🛡️ FINAL SAFETY CHECK - absolutely never allow fewer rows unless deletion was requested
            if (workbookData.length < oldData.length && !userAskedToDeleteRows(prompt)) {
                console.error('🚨 CRITICAL: Data still has fewer rows after merge! Restoring original...');
                workbookData = oldData.map(row => [...row]);
                addChatMessage('system', '🛡️ Emergency protection activated - restored all original rows.');
            }
            
            // Detect what changed (after merge)
            const { changes, modified } = detectChanges(oldHeaders, oldData, headers, workbookData);
            modifiedCells = modified;
            
            // Store for compare feature
            lastAIChange = {
                before: { headers: oldHeaders, data: oldData },
                after: { headers: headers, data: workbookData }
            };
            
            renderTable();
            
            // Show safety warning if rows were protected
            if (mergeResult.message) {
                addChatMessage('system', mergeResult.message);
            }
            
            // Show description of changes
            const description = parsed.description || changes.join(', ') || 'Changes applied';
            addChatMessage('ai', `✓ ${description}`);
            
            if (changes.length > 0 && !parsed.description) {
                addChatMessage('ai', `Details: ${changes.join(', ')}`);
            }
        } else {
            // Fallback - try old format
            if (parsed.headers && parsed.data) {
                const oldHeaders = [...headers];
                const oldData = workbookData.map(row => [...row]);
                
                // ⚠️ SAFETY CHECK: Merge AI changes without losing rows
                const mergeResult = mergeAIChangesIntoOriginal(
                    oldData, 
                    parsed.data, 
                    oldHeaders, 
                    parsed.headers, 
                    prompt
                );
                
                headers = mergeResult.headers;
                workbookData = mergeResult.data;
                
                // 🛡️ FINAL SAFETY CHECK - absolutely never allow fewer rows unless deletion was requested
                if (workbookData.length < oldData.length && !userAskedToDeleteRows(prompt)) {
                    console.error('🚨 CRITICAL: Data still has fewer rows after merge! Restoring original...');
                    workbookData = oldData.map(row => [...row]);
                    addChatMessage('system', '🛡️ Emergency protection activated - restored all original rows.');
                }
                
                const { changes, modified } = detectChanges(oldHeaders, oldData, headers, workbookData);
                modifiedCells = modified;
                
                lastAIChange = {
                    before: { headers: oldHeaders, data: oldData },
                    after: { headers: headers, data: workbookData }
                };
                
                renderTable();
                
                if (mergeResult.message) {
                    addChatMessage('system', mergeResult.message);
                }
                addChatMessage('ai', t('changesApplied'));
            } else {
                throw new Error('Invalid data structure');
            }
        }
    } catch (error) {
        // Remove processing message on error
        const messages = chatMessages.querySelectorAll('.message.ai');
        const lastAiMsg = messages[messages.length - 1];
        if (lastAiMsg && lastAiMsg.textContent === t('processing')) {
            lastAiMsg.remove();
        }
        
        addChatMessage('ai', t('apiError', { error: error.message }));
        showStatus(t('apiError', { error: error.message }), 'error');
        // Remove saved state on error
        history.pop();
    } finally {
        updateButtonStates();
    }
});

undoBtn.addEventListener('click', () => {
    if (history.length === 0) return;
    
    const previous = history.pop();
    headers = previous.headers;
    workbookData = previous.data;
    cellColors = previous.colors;
    modifiedCells = previous.modifiedCells || new Set();
    renderTable();
    addChatMessage('system', t('undidChange'));
    updateButtonStates();
});

addRowBtn.addEventListener('click', () => {
    saveState();
    const newRow = new Array(headers.length).fill('');
    workbookData.push(newRow);
    renderTable();
    addChatMessage('system', t('addedRow'));
});

addColBtn.addEventListener('click', () => {
    const colName = prompt(t('enterColumnName'));
    if (!colName) return;
    
    saveState();
    headers.push(colName);
    workbookData.forEach(row => row.push(''));
    renderTable();
    addChatMessage('system', t('addedColumn', { name: colName }));
});

// Compare modal functionality
compareBtn.addEventListener('click', () => {
    if (history.length === 0) {
        addChatMessage('system', t('noChangesToCompare'));
        return;
    }
    
    // Get the last state from history (before) and current state (after)
    const beforeState = history[history.length - 1];
    
    renderCompareTable('before', beforeState.headers, beforeState.data);
    renderCompareTable('after', headers, workbookData, beforeState.headers, beforeState.data);
    
    compareModal.classList.remove('hidden');
});

closeCompareModal.addEventListener('click', () => {
    compareModal.classList.add('hidden');
});

compareModal.addEventListener('click', (e) => {
    if (e.target === compareModal) {
        compareModal.classList.add('hidden');
    }
});

function renderCompareTable(prefix, tableHeaders, tableData, compareHeaders = null, compareData = null) {
    const thead = document.getElementById(`${prefix}TableHead`);
    const tbody = document.getElementById(`${prefix}TableBody`);
    
    thead.innerHTML = '';
    tbody.innerHTML = '';
    
    // Headers
    const headerRow = document.createElement('tr');
    const cornerCell = document.createElement('th');
    cornerCell.className = 'row-header';
    cornerCell.textContent = '#';
    headerRow.appendChild(cornerCell);
    
    tableHeaders.forEach((header, i) => {
        const th = document.createElement('th');
        th.textContent = header || `Col ${i + 1}`;
        
        // Highlight changed headers
        if (compareHeaders && compareHeaders[i] !== header) {
            th.classList.add('changed');
        }
        headerRow.appendChild(th);
    });
    thead.appendChild(headerRow);
    
    // Data rows
    tableData.forEach((row, rowIndex) => {
        const tr = document.createElement('tr');
        
        const rowHeader = document.createElement('td');
        rowHeader.className = 'row-header';
        rowHeader.textContent = rowIndex + 1;
        tr.appendChild(rowHeader);
        
        tableHeaders.forEach((_, colIndex) => {
            const td = document.createElement('td');
            const value = row[colIndex] !== undefined ? row[colIndex] : '';
            td.textContent = value;
            
            // Highlight differences
            if (compareData) {
                const compareValue = compareData[rowIndex] && compareData[rowIndex][colIndex] !== undefined 
                    ? compareData[rowIndex][colIndex] : '';
                if (String(value) !== String(compareValue)) {
                    td.classList.add('changed');
                }
            }
            
            tr.appendChild(td);
        });
        tbody.appendChild(tr);
    });
}

downloadBtn.addEventListener('click', () => {
    const ws = XLSX.utils.aoa_to_sheet([headers, ...workbookData]);
    
    // Reapply colors
    for (let cell in cellColors) {
        if (ws[cell]) {
            ws[cell].s = {fgColor: {rgb: cellColors[cell].slice(1)}};
        }
    }
    
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
    
    // Use original filename with "_modified" suffix
    const exportName = originalFileName ? `${originalFileName}_modified.xlsx` : 'modified_spreadsheet.xlsx';
    XLSX.writeFile(wb, exportName);
    addChatMessage('system', t('downloaded'));
});

aiPrompt.addEventListener('keydown', (e) => {
    if (e.key === 'Enter' && !e.shiftKey) {
        e.preventDefault();
        aiButton.click();
    }
});

// Escape key to close modal
document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape' && !compareModal.classList.contains('hidden')) {
        compareModal.classList.add('hidden');
    }
});

updateButtonStates();
