let apiKey = '';
let apiKeyPath = '';
let workbookData = [];
let headers = [];
let history = []; // General history for manual edits
let aiHistory = []; // Separate AI change history for granular undo
let columnWidths = {};
let activeFilters = {};
let cellColors = {};
let originalFileName = '';
let modifiedCells = new Set();
let highlightEnabled = false;
let lastAIChange = null;

// Cell selection state
let selectedCells = new Set();
let selectionStart = null;
let selectionEnd = null;
let isSelecting = false;
let isDraggingFillHandle = false;
let fillHandleStart = null;
let clipboard = null;
let preSelectionState = null; // For ESC cancellation

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
const exportFormatSelect = document.getElementById('exportFormat');
const statusDiv = document.getElementById('status');
const tableHead = document.getElementById('tableHead');
const tableBody = document.getElementById('tableBody');
const chatMessages = document.getElementById('chatMessages');
const highlightToggle = document.getElementById('highlightToggle');
const compareModal = document.getElementById('compareModal');
const closeCompareModal = document.getElementById('closeCompareModal');
const aiHistoryModal = document.getElementById('aiHistoryModal');
const closeAiHistoryModal = document.getElementById('closeAiHistoryModal');
const aiHistoryBtn = document.getElementById('aiHistoryBtn');
const aiHistoryList = document.getElementById('aiHistoryList');

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

// Enhanced file loading with multiple format support
excelFile.addEventListener('change', async (e) => {
    const file = e.target.files[0];
    if (!file) return;

    try {
        originalFileName = file.name.replace(/\.[^/.]+$/, '');
        const fileExt = file.name.split('.').pop().toLowerCase();
        
        let json;
        
        if (fileExt === 'csv') {
            // Parse CSV with better handling
            const text = await file.text();
            json = parseCSV(text);
        } else if (fileExt === 'tsv' || fileExt === 'txt') {
            // Parse TSV/Tab-delimited
            const text = await file.text();
            json = parseTSV(text);
        } else if (fileExt === 'json') {
            // Parse JSON data
            const text = await file.text();
            json = parseJSON(text);
        } else if (fileExt === 'xml') {
            // Parse XML data
            const text = await file.text();
            json = parseXML(text);
        } else if (fileExt === 'ods') {
            // OpenDocument Spreadsheet
            const data = await file.arrayBuffer();
            const workbook = XLSX.read(data, { type: 'array', cellStyles: true });
            const sheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[sheetName];
            extractCellColors(worksheet);
            json = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });
        } else {
            // Excel formats (.xlsx, .xls, .xlsm, .xlsb)
            const data = await file.arrayBuffer();
            const workbook = XLSX.read(data, { type: 'array', cellStyles: true });
            const sheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[sheetName];
            extractCellColors(worksheet);
            json = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });
        }

        if (!json || json.length === 0) {
            showStatus(t('errorEmpty'), 'error');
            return;
        }

        headers = json[0].map(h => h !== undefined && h !== null ? String(h) : '');
        workbookData = json.slice(1).map(row => 
            row.map(cell => cell !== undefined && cell !== null ? cell : '')
        );
        
        // Ensure all rows have same number of columns as headers
        workbookData = workbookData.map(row => {
            while (row.length < headers.length) row.push('');
            return row.slice(0, headers.length);
        });
        
        history = [];
        aiHistory = [];
        activeFilters = {};
        modifiedCells.clear();
        lastAIChange = null;
        selectedCells.clear();
        
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

// Extract cell colors from worksheet
function extractCellColors(worksheet) {
    cellColors = {};
    for (let cell in worksheet) {
        if (cell[0] === '!') continue;
        if (worksheet[cell].s && worksheet[cell].s.fgColor) {
            cellColors[cell] = '#' + worksheet[cell].s.fgColor.rgb;
        }
    }
}

// CSV Parser with quote handling
function parseCSV(text) {
    const lines = [];
    let currentLine = [];
    let currentCell = '';
    let inQuotes = false;
    
    for (let i = 0; i < text.length; i++) {
        const char = text[i];
        const nextChar = text[i + 1];
        
        if (inQuotes) {
            if (char === '"' && nextChar === '"') {
                currentCell += '"';
                i++;
            } else if (char === '"') {
                inQuotes = false;
            } else {
                currentCell += char;
            }
        } else {
            if (char === '"') {
                inQuotes = true;
            } else if (char === ',') {
                currentLine.push(currentCell);
                currentCell = '';
            } else if (char === '\n' || (char === '\r' && nextChar === '\n')) {
                currentLine.push(currentCell);
                lines.push(currentLine);
                currentLine = [];
                currentCell = '';
                if (char === '\r') i++;
            } else if (char !== '\r') {
                currentCell += char;
            }
        }
    }
    
    if (currentCell || currentLine.length > 0) {
        currentLine.push(currentCell);
        lines.push(currentLine);
    }
    
    return lines.filter(line => line.some(cell => cell !== ''));
}

// TSV Parser
function parseTSV(text) {
    const lines = text.split(/\r?\n/).filter(line => line.trim());
    return lines.map(line => line.split('\t'));
}

// JSON Parser - handles arrays of objects or arrays
function parseJSON(text) {
    const data = JSON.parse(text);
    
    if (Array.isArray(data)) {
        if (data.length === 0) return [];
        
        if (typeof data[0] === 'object' && !Array.isArray(data[0])) {
            // Array of objects
            const headers = Object.keys(data[0]);
            const rows = data.map(obj => headers.map(h => obj[h] ?? ''));
            return [headers, ...rows];
        } else if (Array.isArray(data[0])) {
            // Array of arrays
            return data;
        }
    } else if (typeof data === 'object') {
        // Single object or nested structure
        const headers = Object.keys(data);
        const values = headers.map(h => data[h]);
        return [headers, values];
    }
    
    throw new Error('Unsupported JSON structure');
}

// XML Parser - basic table/row/cell structure
function parseXML(text) {
    const parser = new DOMParser();
    const doc = parser.parseFromString(text, 'text/xml');
    
    // Try common XML table structures
    const rows = doc.querySelectorAll('row, tr, record, item');
    if (rows.length === 0) {
        // Try to parse as a flat structure
        const root = doc.documentElement;
        const children = root.children;
        if (children.length > 0) {
            const firstChild = children[0];
            const headers = Array.from(firstChild.children).map(el => el.tagName);
            const data = Array.from(children).map(row => 
                Array.from(row.children).map(cell => cell.textContent)
            );
            return [headers, ...data];
        }
    }
    
    const result = [];
    let headers = null;
    
    rows.forEach((row, idx) => {
        const cells = row.querySelectorAll('cell, td, th, field, *');
        const rowData = Array.from(cells).map(cell => cell.textContent);
        
        if (idx === 0) {
            // Check if first row looks like headers
            const headerCells = row.querySelectorAll('th, header');
            if (headerCells.length > 0 || row.tagName.toLowerCase() === 'header') {
                headers = rowData;
            } else {
                // Generate headers
                headers = rowData.map((_, i) => `Column ${i + 1}`);
                result.push(headers);
            }
            result.push(headers === rowData ? rowData : rowData);
        } else {
            result.push(rowData);
        }
    });
    
    return result;
}

function saveState() {
    history.push({
        headers: JSON.parse(JSON.stringify(headers)),
        data: JSON.parse(JSON.stringify(workbookData)),
        colors: JSON.parse(JSON.stringify(cellColors)),
        modifiedCells: new Set(modifiedCells)
    });
    if (history.length > 50) history.shift();
    updateButtonStates();
}

// Save AI-specific state for granular undo
function saveAIState(description, prompt) {
    aiHistory.push({
        id: Date.now(),
        timestamp: new Date().toLocaleString(),
        description: description,
        prompt: prompt,
        beforeHeaders: JSON.parse(JSON.stringify(history[history.length - 1]?.headers || headers)),
        beforeData: JSON.parse(JSON.stringify(history[history.length - 1]?.data || workbookData)),
        afterHeaders: JSON.parse(JSON.stringify(headers)),
        afterData: JSON.parse(JSON.stringify(workbookData)),
        modifiedCells: new Set(modifiedCells),
        isReverted: false
    });
    if (aiHistory.length > 100) aiHistory.shift();
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
            
            // Apply selection highlight
            if (selectedCells.has(cellKey)) {
                td.classList.add('selected');
            }

            const input = document.createElement('input');
            input.type = 'text';
            input.value = row.data[colIndex] !== undefined ? row.data[colIndex] : '';
            input.dataset.row = row.originalIndex;
            input.dataset.col = colIndex;
            
            // Cell events
            input.addEventListener('change', (e) => {
                saveState();
                workbookData[row.originalIndex][colIndex] = e.target.value;
            });
            
            input.addEventListener('mousedown', (e) => handleCellMouseDown(e, row.originalIndex, colIndex));
            input.addEventListener('mouseenter', (e) => handleCellMouseEnter(e, row.originalIndex, colIndex));
            input.addEventListener('focus', () => handleCellFocus(row.originalIndex, colIndex));
            
            td.appendChild(input);
            
            // Add fill handle to last selected cell
            if (selectedCells.size > 0 && isLastSelectedCell(row.originalIndex, colIndex)) {
                const fillHandle = document.createElement('div');
                fillHandle.className = 'fill-handle';
                fillHandle.addEventListener('mousedown', (e) => startFillHandle(e, row.originalIndex, colIndex));
                td.appendChild(fillHandle);
            }
            
            tr.appendChild(td);
        });
        tableBody.appendChild(tr);
    });
}

// Check if this is the bottom-right cell of selection
function isLastSelectedCell(rowIndex, colIndex) {
    if (selectedCells.size === 0) return false;
    
    let maxRow = -1, maxCol = -1;
    selectedCells.forEach(key => {
        const [r, c] = key.split('-').map(Number);
        if (r > maxRow || (r === maxRow && c > maxCol)) {
            maxRow = r;
            maxCol = c;
        }
    });
    
    return rowIndex === maxRow && colIndex === maxCol;
}

// Cell selection handlers

function handleCellMouseDown(e, rowIndex, colIndex) {
    if (e.button !== 0) return; // Left click only
    
    const cellKey = `${rowIndex}-${colIndex}`;
    
    // Save current state before starting selection (for ESC cancellation)
    preSelectionState = {
        selectedCells: new Set(selectedCells),
        selectionStart: selectionStart ? { ...selectionStart } : null,
        selectionEnd: selectionEnd ? { ...selectionEnd } : null
    };
    
    if (e.shiftKey && selectionStart) {
        // Extend selection from existing start point
        isSelecting = true;
        extendSelection(rowIndex, colIndex);
    } else if (e.ctrlKey || e.metaKey) {
        // Toggle single cell in selection (add/remove from multi-select)
        if (selectedCells.has(cellKey)) {
            selectedCells.delete(cellKey);
        } else {
            selectedCells.add(cellKey);
        }
        // Don't start drag selection for Ctrl+click
        isSelecting = false;
    } else {
        // Start new selection - clear previous and begin fresh
        isSelecting = true;
        selectionStart = { row: rowIndex, col: colIndex };
        selectionEnd = { row: rowIndex, col: colIndex };
        selectedCells.clear();
        selectedCells.add(cellKey);
    }
    
    renderTable();
}

function handleCellMouseEnter(e, rowIndex, colIndex) {
    if (isDraggingFillHandle) {
        updateFillPreview(rowIndex, colIndex);
    } else if (isSelecting && e.buttons === 1) {
        // Only extend selection if mouse button is still held down
        extendSelection(rowIndex, colIndex);
        renderTable();
    }
}

function handleCellFocus(rowIndex, colIndex) {
    // Only update selection on focus if we're not in the middle of a drag selection
    // and the user just clicked (not dragged) on a single cell
    if (!isSelecting) {
        // Don't auto-select on focus - let mousedown handle it
        // This prevents the re-selection issue
    }
}

function extendSelection(endRow, endCol) {
    if (!selectionStart) return;
    
    selectionEnd = { row: endRow, col: endCol };
    selectedCells.clear();
    
    const minRow = Math.min(selectionStart.row, endRow);
    const maxRow = Math.max(selectionStart.row, endRow);
    const minCol = Math.min(selectionStart.col, endCol);
    const maxCol = Math.max(selectionStart.col, endCol);
    
    for (let r = minRow; r <= maxRow; r++) {
        for (let c = minCol; c <= maxCol; c++) {
            selectedCells.add(`${r}-${c}`);
        }
    }
}

// Fill handle functionality (Excel-like drag to fill)
function startFillHandle(e, rowIndex, colIndex) {
    e.preventDefault();
    e.stopPropagation();
    
    isDraggingFillHandle = true;
    fillHandleStart = { row: rowIndex, col: colIndex };
    
    document.addEventListener('mousemove', onFillHandleMove);
    document.addEventListener('mouseup', onFillHandleUp);
}

function onFillHandleMove(e) {
    // Find cell under cursor
    const element = document.elementFromPoint(e.clientX, e.clientY);
    if (element && element.tagName === 'INPUT' && element.dataset.row !== undefined) {
        const row = parseInt(element.dataset.row);
        const col = parseInt(element.dataset.col);
        updateFillPreview(row, col);
    }
}

function updateFillPreview(endRow, endCol) {
    // Visual feedback for fill preview
    document.querySelectorAll('.fill-preview').forEach(el => el.classList.remove('fill-preview'));
    
    if (!fillHandleStart || !selectionStart || !selectionEnd) return;
    
    const sourceMinRow = Math.min(selectionStart.row, selectionEnd.row);
    const sourceMaxRow = Math.max(selectionStart.row, selectionEnd.row);
    const sourceMinCol = Math.min(selectionStart.col, selectionEnd.col);
    const sourceMaxCol = Math.max(selectionStart.col, selectionEnd.col);
    
    // Determine fill direction
    const rowDiff = endRow - sourceMaxRow;
    const colDiff = endCol - sourceMaxCol;
    
    let targetCells = [];
    
    if (Math.abs(rowDiff) > Math.abs(colDiff)) {
        // Fill vertically
        if (rowDiff > 0) {
            for (let r = sourceMaxRow + 1; r <= endRow; r++) {
                for (let c = sourceMinCol; c <= sourceMaxCol; c++) {
                    targetCells.push(`${r}-${c}`);
                }
            }
        } else if (endRow < sourceMinRow) {
            for (let r = endRow; r < sourceMinRow; r++) {
                for (let c = sourceMinCol; c <= sourceMaxCol; c++) {
                    targetCells.push(`${r}-${c}`);
                }
            }
        }
    } else {
        // Fill horizontally
        if (colDiff > 0) {
            for (let r = sourceMinRow; r <= sourceMaxRow; r++) {
                for (let c = sourceMaxCol + 1; c <= endCol; c++) {
                    targetCells.push(`${r}-${c}`);
                }
            }
        } else if (endCol < sourceMinCol) {
            for (let r = sourceMinRow; r <= sourceMaxRow; r++) {
                for (let c = endCol; c < sourceMinCol; c++) {
                    targetCells.push(`${r}-${c}`);
                }
            }
        }
    }
    
    // Add visual preview
    targetCells.forEach(cellKey => {
        const [r, c] = cellKey.split('-').map(Number);
        const input = document.querySelector(`input[data-row="${r}"][data-col="${c}"]`);
        if (input) {
            input.parentElement.classList.add('fill-preview');
        }
    });
}

function onFillHandleUp(e) {
    document.removeEventListener('mousemove', onFillHandleMove);
    document.removeEventListener('mouseup', onFillHandleUp);
    
    const element = document.elementFromPoint(e.clientX, e.clientY);
    if (element && element.tagName === 'INPUT' && element.dataset.row !== undefined) {
        const endRow = parseInt(element.dataset.row);
        const endCol = parseInt(element.dataset.col);
        performFill(endRow, endCol);
    }
    
    isDraggingFillHandle = false;
    fillHandleStart = null;
    document.querySelectorAll('.fill-preview').forEach(el => el.classList.remove('fill-preview'));
    renderTable();
}

function performFill(endRow, endCol) {
    if (!selectionStart || !selectionEnd) return;
    
    saveState();
    
    const sourceMinRow = Math.min(selectionStart.row, selectionEnd.row);
    const sourceMaxRow = Math.max(selectionStart.row, selectionEnd.row);
    const sourceMinCol = Math.min(selectionStart.col, selectionEnd.col);
    const sourceMaxCol = Math.max(selectionStart.col, selectionEnd.col);
    
    const sourceHeight = sourceMaxRow - sourceMinRow + 1;
    const sourceWidth = sourceMaxCol - sourceMinCol + 1;
    
    // Get source values
    const sourceValues = [];
    for (let r = sourceMinRow; r <= sourceMaxRow; r++) {
        const row = [];
        for (let c = sourceMinCol; c <= sourceMaxCol; c++) {
            row.push(workbookData[r][c]);
        }
        sourceValues.push(row);
    }
    
    // Detect patterns for smart fill
    const patterns = detectPatterns(sourceValues);
    
    const rowDiff = endRow - sourceMaxRow;
    const colDiff = endCol - sourceMaxCol;
    
    if (Math.abs(rowDiff) > Math.abs(colDiff)) {
        // Fill vertically
        if (rowDiff > 0) {
            for (let r = sourceMaxRow + 1; r <= endRow; r++) {
                const sourceRowIdx = (r - sourceMinRow) % sourceHeight;
                const repeatNum = Math.floor((r - sourceMinRow) / sourceHeight);
                for (let c = sourceMinCol; c <= sourceMaxCol; c++) {
                    const sourceColIdx = c - sourceMinCol;
                    workbookData[r][c] = getFilledValue(sourceValues[sourceRowIdx][sourceColIdx], patterns[sourceColIdx], repeatNum, r - sourceMinRow);
                }
            }
        } else if (endRow < sourceMinRow) {
            for (let r = sourceMinRow - 1; r >= endRow; r--) {
                const sourceRowIdx = (sourceMinRow - r - 1) % sourceHeight;
                const repeatNum = -Math.floor((sourceMinRow - r) / sourceHeight) - 1;
                for (let c = sourceMinCol; c <= sourceMaxCol; c++) {
                    const sourceColIdx = c - sourceMinCol;
                    workbookData[r][c] = getFilledValue(sourceValues[sourceRowIdx][sourceColIdx], patterns[sourceColIdx], repeatNum, r - sourceMinRow);
                }
            }
        }
    } else {
        // Fill horizontally
        if (colDiff > 0) {
            for (let r = sourceMinRow; r <= sourceMaxRow; r++) {
                const sourceRowIdx = r - sourceMinRow;
                for (let c = sourceMaxCol + 1; c <= endCol; c++) {
                    const sourceColIdx = (c - sourceMinCol) % sourceWidth;
                    const repeatNum = Math.floor((c - sourceMinCol) / sourceWidth);
                    workbookData[r][c] = getFilledValue(sourceValues[sourceRowIdx][sourceColIdx], patterns[sourceColIdx], repeatNum, c - sourceMinCol);
                }
            }
        } else if (endCol < sourceMinCol) {
            for (let r = sourceMinRow; r <= sourceMaxRow; r++) {
                const sourceRowIdx = r - sourceMinRow;
                for (let c = sourceMinCol - 1; c >= endCol; c--) {
                    const sourceColIdx = (sourceMinCol - c - 1) % sourceWidth;
                    const repeatNum = -Math.floor((sourceMinCol - c) / sourceWidth) - 1;
                    workbookData[r][c] = getFilledValue(sourceValues[sourceRowIdx][sourceColIdx], patterns[sourceColIdx], repeatNum, c - sourceMinCol);
                }
            }
        }
    }
    
    addChatMessage('system', t('fillApplied'));
}

// Detect numeric patterns for smart fill
function detectPatterns(values) {
    const patterns = [];
    
    for (let c = 0; c < values[0].length; c++) {
        const colValues = values.map(row => row[c]);
        
        if (colValues.length >= 2) {
            // Check for numeric sequence
            const numbers = colValues.map(v => parseFloat(v)).filter(n => !isNaN(n));
            if (numbers.length === colValues.length && numbers.length >= 2) {
                const diffs = [];
                for (let i = 1; i < numbers.length; i++) {
                    diffs.push(numbers[i] - numbers[i-1]);
                }
                // Check if differences are constant (arithmetic sequence)
                const allSame = diffs.every(d => Math.abs(d - diffs[0]) < 0.0001);
                if (allSame) {
                    patterns.push({ type: 'arithmetic', step: diffs[0], start: numbers[0] });
                    continue;
                }
            }
            
            // Check for date patterns
            const dates = colValues.map(v => new Date(v)).filter(d => !isNaN(d.getTime()));
            if (dates.length === colValues.length && dates.length >= 2) {
                const dayDiff = (dates[1] - dates[0]) / (1000 * 60 * 60 * 24);
                patterns.push({ type: 'date', dayStep: dayDiff, start: dates[0] });
                continue;
            }
        }
        
        // Default: repeat pattern
        patterns.push({ type: 'repeat' });
    }
    
    return patterns;
}

function getFilledValue(sourceValue, pattern, repeatNum, position) {
    if (!pattern || pattern.type === 'repeat') {
        return sourceValue;
    }
    
    if (pattern.type === 'arithmetic') {
        const num = parseFloat(sourceValue);
        if (!isNaN(num)) {
            return String(num + pattern.step * position);
        }
    }
    
    if (pattern.type === 'date') {
        const date = new Date(sourceValue);
        if (!isNaN(date.getTime())) {
            const newDate = new Date(date.getTime() + pattern.dayStep * position * 24 * 60 * 60 * 1000);
            return newDate.toISOString().split('T')[0];
        }
    }
    
    return sourceValue;
}

// Mouse up handler for ending selection
document.addEventListener('mouseup', (e) => {
    if (isSelecting) {
        isSelecting = false;
        // Selection is now finalized - clear the pre-selection state
        // Keep preSelectionState until next mousedown for potential ESC cancellation
    }
});

// Keyboard shortcuts
document.addEventListener('keydown', (e) => {
    // Copy (Ctrl+C)
    if ((e.ctrlKey || e.metaKey) && e.key === 'c' && selectedCells.size > 0) {
        e.preventDefault();
        copySelectedCells();
    }
    
    // Cut (Ctrl+X)
    if ((e.ctrlKey || e.metaKey) && e.key === 'x' && selectedCells.size > 0) {
        e.preventDefault();
        cutSelectedCells();
    }
    
    // Paste (Ctrl+V)
    if ((e.ctrlKey || e.metaKey) && e.key === 'v' && clipboard) {
        e.preventDefault();
        pasteClipboard();
    }
    
    // Delete (Delete key)
    if (e.key === 'Delete' && selectedCells.size > 0) {
        e.preventDefault();
        deleteSelectedCells();
    }
    
    // Fill down (Ctrl+D)
    if ((e.ctrlKey || e.metaKey) && e.key === 'd' && selectedCells.size > 0) {
        e.preventDefault();
        fillDown();
    }
    
    // Fill right (Ctrl+R)
    if ((e.ctrlKey || e.metaKey) && e.key === 'r' && selectedCells.size > 0) {
        e.preventDefault();
        fillRight();
    }
    
    // Select all (Ctrl+A) when in table
    if ((e.ctrlKey || e.metaKey) && e.key === 'a') {
        const activeEl = document.activeElement;
        if (activeEl && activeEl.tagName === 'INPUT' && activeEl.dataset.row !== undefined) {
            e.preventDefault();
            selectAllCells();
        }
    }
    
    // Escape to cancel selection or close modals
    if (e.key === 'Escape') {
        if (!compareModal.classList.contains('hidden')) {
            compareModal.classList.add('hidden');
        } else if (!aiHistoryModal.classList.contains('hidden')) {
            aiHistoryModal.classList.add('hidden');
        } else if (isSelecting) {
            // Cancel ongoing selection and restore previous state
            isSelecting = false;
            if (preSelectionState) {
                selectedCells = preSelectionState.selectedCells;
                selectionStart = preSelectionState.selectionStart;
                selectionEnd = preSelectionState.selectionEnd;
                preSelectionState = null;
            }
            renderTable();
        } else if (selectedCells.size > 0) {
            // Clear selection if not in the middle of selecting
            selectedCells.clear();
            selectionStart = null;
            selectionEnd = null;
            preSelectionState = null;
            renderTable();
        }
    }
});

function copySelectedCells() {
    if (selectedCells.size === 0) return;
    
    const cells = Array.from(selectedCells).map(key => key.split('-').map(Number));
    const minRow = Math.min(...cells.map(c => c[0]));
    const maxRow = Math.max(...cells.map(c => c[0]));
    const minCol = Math.min(...cells.map(c => c[1]));
    const maxCol = Math.max(...cells.map(c => c[1]));
    
    clipboard = {
        data: [],
        width: maxCol - minCol + 1,
        height: maxRow - minRow + 1,
        isCut: false
    };
    
    for (let r = minRow; r <= maxRow; r++) {
        const row = [];
        for (let c = minCol; c <= maxCol; c++) {
            row.push(workbookData[r] ? workbookData[r][c] || '' : '');
        }
        clipboard.data.push(row);
    }
    
    // Also copy to system clipboard as TSV
    const tsv = clipboard.data.map(row => row.join('\t')).join('\n');
    navigator.clipboard.writeText(tsv).catch(() => {});
    
    addChatMessage('system', t('copiedCells', { count: selectedCells.size }));
}

function cutSelectedCells() {
    copySelectedCells();
    if (clipboard) {
        clipboard.isCut = true;
        clipboard.sourceStart = {
            row: Math.min(...Array.from(selectedCells).map(k => parseInt(k.split('-')[0]))),
            col: Math.min(...Array.from(selectedCells).map(k => parseInt(k.split('-')[1])))
        };
    }
}

function pasteClipboard() {
    if (!clipboard || selectedCells.size === 0) return;
    
    saveState();
    
    const targetCells = Array.from(selectedCells).map(key => key.split('-').map(Number));
    const targetMinRow = Math.min(...targetCells.map(c => c[0]));
    const targetMinCol = Math.min(...targetCells.map(c => c[1]));
    
    // Paste data
    for (let r = 0; r < clipboard.height; r++) {
        for (let c = 0; c < clipboard.width; c++) {
            const targetRow = targetMinRow + r;
            const targetCol = targetMinCol + c;
            
            if (targetRow < workbookData.length && targetCol < headers.length) {
                workbookData[targetRow][targetCol] = clipboard.data[r][c];
            }
        }
    }
    
    // If cut, clear source cells
    if (clipboard.isCut && clipboard.sourceStart) {
        for (let r = 0; r < clipboard.height; r++) {
            for (let c = 0; c < clipboard.width; c++) {
                const srcRow = clipboard.sourceStart.row + r;
                const srcCol = clipboard.sourceStart.col + c;
                if (srcRow !== targetMinRow + r || srcCol !== targetMinCol + c) {
                    if (srcRow < workbookData.length && srcCol < headers.length) {
                        workbookData[srcRow][srcCol] = '';
                    }
                }
            }
        }
        clipboard.isCut = false;
    }
    
    renderTable();
    addChatMessage('system', t('pastedCells'));
}

function deleteSelectedCells() {
    if (selectedCells.size === 0) return;
    
    saveState();
    
    selectedCells.forEach(cellKey => {
        const [r, c] = cellKey.split('-').map(Number);
        if (workbookData[r]) {
            workbookData[r][c] = '';
        }
    });
    
    renderTable();
    addChatMessage('system', t('deletedCells', { count: selectedCells.size }));
}

function fillDown() {
    if (selectedCells.size === 0) return;
    
    saveState();
    
    const cells = Array.from(selectedCells).map(key => key.split('-').map(Number));
    const minRow = Math.min(...cells.map(c => c[0]));
    const maxRow = Math.max(...cells.map(c => c[0]));
    const cols = [...new Set(cells.map(c => c[1]))].sort((a,b) => a-b);
    
    cols.forEach(col => {
        const sourceValue = workbookData[minRow][col];
        for (let r = minRow + 1; r <= maxRow; r++) {
            workbookData[r][col] = sourceValue;
        }
    });
    
    renderTable();
    addChatMessage('system', t('fillDownApplied'));
}

function fillRight() {
    if (selectedCells.size === 0) return;
    
    saveState();
    
    const cells = Array.from(selectedCells).map(key => key.split('-').map(Number));
    const minCol = Math.min(...cells.map(c => c[1]));
    const maxCol = Math.max(...cells.map(c => c[1]));
    const rows = [...new Set(cells.map(c => c[0]))].sort((a,b) => a-b);
    
    rows.forEach(row => {
        const sourceValue = workbookData[row][minCol];
        for (let c = minCol + 1; c <= maxCol; c++) {
            workbookData[row][c] = sourceValue;
        }
    });
    
    renderTable();
    addChatMessage('system', t('fillRightApplied'));
}

function selectAllCells() {
    selectedCells.clear();
    selectionStart = { row: 0, col: 0 };
    selectionEnd = { row: workbookData.length - 1, col: headers.length - 1 };
    
    for (let r = 0; r < workbookData.length; r++) {
        for (let c = 0; c < headers.length; c++) {
            selectedCells.add(`${r}-${c}`);
        }
    }
    
    renderTable();
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
                activeFilters[colIndex] = new Set(uniqueValues);
            }
            
            if (checkbox.checked) {
                activeFilters[colIndex].add(value);
            } else {
                activeFilters[colIndex].delete(value);
            }
            
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
    
    const th = e.target.closest('th');
    th.appendChild(dropdown);

    searchInput.oninput = () => {
        const search = searchInput.value.toLowerCase();
        optionsDiv.querySelectorAll('.filter-option').forEach(opt => {
            const text = opt.textContent.toLowerCase();
            opt.style.display = text.includes(search) ? 'flex' : 'none';
        });
    };

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
        
        const rows = tableBody.querySelectorAll('tr');
        rows.forEach(row => {
            const cell = row.children[colIndex + 1];
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
    const hasAIHistory = aiHistory.length > 0;
    
    aiButton.disabled = !hasData;
    aiPrompt.disabled = !hasData;
    undoBtn.disabled = !hasHistory;
    addRowBtn.disabled = !hasData;
    addColBtn.disabled = !hasData;
    downloadBtn.disabled = !hasData;
    compareBtn.disabled = !hasHistory;
    aiHistoryBtn.disabled = !hasAIHistory;
}

// Detect changes between two data sets
function detectChanges(oldHeaders, oldData, newHeaders, newData) {
    const changes = [];
    const modified = new Set();
    
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
    
    if (newData.length > oldData.length) {
        changes.push(`Added ${newData.length - oldData.length} row(s)`);
    } else if (newData.length < oldData.length) {
        changes.push(`Removed ${oldData.length - newData.length} row(s)`);
    }
    
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

// Safety mechanisms
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

function mergeAIChangesIntoOriginal(oldData, newData, oldHeaders, newHeaders, userPrompt) {
    if (userAskedToDeleteRows(userPrompt)) {
        return { 
            data: newData, 
            headers: newHeaders, 
            restored: 0,
            message: null 
        };
    }
    
    if (newData.length >= oldData.length) {
        return { 
            data: newData, 
            headers: newHeaders, 
            restored: 0,
            message: null 
        };
    }
    
    const finalHeaders = newHeaders.length >= oldHeaders.length ? [...newHeaders] : [...oldHeaders];
    const numCols = finalHeaders.length;
    
    const finalData = [];
    for (let i = 0; i < oldData.length; i++) {
        const row = [];
        for (let j = 0; j < numCols; j++) {
            row.push(oldData[i] && oldData[i][j] !== undefined ? oldData[i][j] : '');
        }
        finalData.push(row);
    }
    
    for (let aiRowIdx = 0; aiRowIdx < newData.length; aiRowIdx++) {
        const aiRow = newData[aiRowIdx];
        
        let bestMatchIdx = -1;
        let bestMatchScore = -1;
        
        for (let origIdx = 0; origIdx < oldData.length; origIdx++) {
            const origRow = oldData[origIdx];
            let score = 0;
            
            for (let col = 0; col < Math.min(origRow.length, aiRow.length); col++) {
                const origVal = String(origRow[col] ?? '');
                const aiVal = String(aiRow[col] ?? '');
                if (origVal === aiVal) {
                    score += 10;
                } else if (origVal !== '' && aiVal !== '' && 
                          (origVal.includes(aiVal) || aiVal.includes(origVal))) {
                    score += 3;
                }
            }
            
            if (origIdx === aiRowIdx) score += 20;
            else if (Math.abs(origIdx - aiRowIdx) === 1) score += 10;
            else if (Math.abs(origIdx - aiRowIdx) <= 3) score += 5;
            
            if (score > bestMatchScore) {
                bestMatchScore = score;
                bestMatchIdx = origIdx;
            }
        }
        
        if (bestMatchIdx !== -1 && bestMatchScore > 0) {
            for (let col = 0; col < aiRow.length; col++) {
                const origVal = String(oldData[bestMatchIdx]?.[col] ?? '');
                const aiVal = String(aiRow[col] ?? '');
                
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

// AI Button handler
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
            addChatMessage('ai', parsed.answer);
            history.pop();
        } else if (parsed.type === 'modification' && parsed.headers && parsed.data) {
            const oldHeaders = [...headers];
            const oldData = workbookData.map(row => [...row]);
            
            const mergeResult = mergeAIChangesIntoOriginal(
                oldData, 
                parsed.data, 
                oldHeaders, 
                parsed.headers, 
                prompt
            );
            
            headers = mergeResult.headers;
            workbookData = mergeResult.data;
            
            if (workbookData.length < oldData.length && !userAskedToDeleteRows(prompt)) {
                workbookData = oldData.map(row => [...row]);
                addChatMessage('system', '🛡️ Emergency protection activated - restored all original rows.');
            }
            
            const { changes, modified } = detectChanges(oldHeaders, oldData, headers, workbookData);
            modifiedCells = modified;
            
            // Store for compare feature
            lastAIChange = {
                before: { headers: oldHeaders, data: oldData },
                after: { headers: headers, data: workbookData }
            };
            
            // Save to AI history for granular undo
            const description = parsed.description || changes.join(', ') || 'Changes applied';
            saveAIState(description, prompt);
            
            renderTable();
            
            if (mergeResult.message) {
                addChatMessage('system', mergeResult.message);
            }
            
            addChatMessage('ai', `✓ ${description}`);
            
            if (changes.length > 0 && !parsed.description) {
                addChatMessage('ai', `Details: ${changes.join(', ')}`);
            }
        } else {
            // Fallback
            if (parsed.headers && parsed.data) {
                const oldHeaders = [...headers];
                const oldData = workbookData.map(row => [...row]);
                
                const mergeResult = mergeAIChangesIntoOriginal(
                    oldData, 
                    parsed.data, 
                    oldHeaders, 
                    parsed.headers, 
                    prompt
                );
                
                headers = mergeResult.headers;
                workbookData = mergeResult.data;
                
                if (workbookData.length < oldData.length && !userAskedToDeleteRows(prompt)) {
                    workbookData = oldData.map(row => [...row]);
                    addChatMessage('system', '🛡️ Emergency protection activated - restored all original rows.');
                }
                
                const { changes, modified } = detectChanges(oldHeaders, oldData, headers, workbookData);
                modifiedCells = modified;
                
                lastAIChange = {
                    before: { headers: oldHeaders, data: oldData },
                    after: { headers: headers, data: workbookData }
                };
                
                saveAIState('Changes applied', prompt);
                
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
        const messages = chatMessages.querySelectorAll('.message.ai');
        const lastAiMsg = messages[messages.length - 1];
        if (lastAiMsg && lastAiMsg.textContent === t('processing')) {
            lastAiMsg.remove();
        }
        
        addChatMessage('ai', t('apiError', { error: error.message }));
        showStatus(t('apiError', { error: error.message }), 'error');
        history.pop();
    } finally {
        updateButtonStates();
    }
});

// Standard undo
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

// AI History Modal
aiHistoryBtn.addEventListener('click', () => {
    renderAIHistory();
    aiHistoryModal.classList.remove('hidden');
});

closeAiHistoryModal.addEventListener('click', () => {
    aiHistoryModal.classList.add('hidden');
});

aiHistoryModal.addEventListener('click', (e) => {
    if (e.target === aiHistoryModal) {
        aiHistoryModal.classList.add('hidden');
    }
});

function renderAIHistory() {
    aiHistoryList.innerHTML = '';
    
    if (aiHistory.length === 0) {
        aiHistoryList.innerHTML = '<p class="no-history">' + t('noAIHistory') + '</p>';
        return;
    }
    
    // Reverse to show most recent first
    [...aiHistory].reverse().forEach((entry, idx) => {
        const realIdx = aiHistory.length - 1 - idx;
        const item = document.createElement('div');
        item.className = 'ai-history-item' + (entry.isReverted ? ' reverted' : '');
        
        const header = document.createElement('div');
        header.className = 'ai-history-header';
        
        const timestamp = document.createElement('span');
        timestamp.className = 'ai-history-timestamp';
        timestamp.textContent = entry.timestamp;
        header.appendChild(timestamp);
        
        const status = document.createElement('span');
        status.className = 'ai-history-status';
        status.textContent = entry.isReverted ? '↩️ Reverted' : '✓ Applied';
        header.appendChild(status);
        
        item.appendChild(header);
        
        const prompt = document.createElement('div');
        prompt.className = 'ai-history-prompt';
        prompt.textContent = `"${entry.prompt}"`;
        item.appendChild(prompt);
        
        const description = document.createElement('div');
        description.className = 'ai-history-description';
        description.textContent = entry.description;
        item.appendChild(description);
        
        const actions = document.createElement('div');
        actions.className = 'ai-history-actions';
        
        if (!entry.isReverted) {
            const revertBtn = document.createElement('button');
            revertBtn.className = 'ai-history-revert';
            revertBtn.textContent = t('revertChange');
            revertBtn.onclick = () => revertAIChange(realIdx);
            actions.appendChild(revertBtn);
        } else {
            const reapplyBtn = document.createElement('button');
            reapplyBtn.className = 'ai-history-reapply';
            reapplyBtn.textContent = t('reapplyChange');
            reapplyBtn.onclick = () => reapplyAIChange(realIdx);
            actions.appendChild(reapplyBtn);
        }
        
        const viewBtn = document.createElement('button');
        viewBtn.className = 'ai-history-view';
        viewBtn.textContent = t('viewDiff');
        viewBtn.onclick = () => viewAIDiff(realIdx);
        actions.appendChild(viewBtn);
        
        item.appendChild(actions);
        aiHistoryList.appendChild(item);
    });
}

function revertAIChange(index) {
    const entry = aiHistory[index];
    if (!entry || entry.isReverted) return;
    
    saveState();
    
    // Calculate the changes this entry made
    const beforeData = entry.beforeData;
    const afterData = entry.afterData;
    const beforeHeaders = entry.beforeHeaders;
    const afterHeaders = entry.afterHeaders;
    
    // We need to "subtract" this change from current data
    // This is complex because subsequent changes may have modified the same cells
    // Simple approach: revert cells that were changed by this AI call AND haven't been changed since
    
    for (let r = 0; r < Math.min(beforeData.length, workbookData.length); r++) {
        for (let c = 0; c < Math.min(beforeHeaders.length, headers.length); c++) {
            const beforeVal = String(beforeData[r]?.[c] ?? '');
            const afterVal = String(afterData[r]?.[c] ?? '');
            const currentVal = String(workbookData[r]?.[c] ?? '');
            
            // If this cell was changed by this AI call, and current value matches the AI's change
            if (beforeVal !== afterVal && currentVal === afterVal) {
                workbookData[r][c] = beforeData[r][c];
            }
        }
    }
    
    entry.isReverted = true;
    
    renderTable();
    renderAIHistory();
    addChatMessage('system', t('revertedAIChange', { description: entry.description }));
    updateButtonStates();
}

function reapplyAIChange(index) {
    const entry = aiHistory[index];
    if (!entry || !entry.isReverted) return;
    
    saveState();
    
    const beforeData = entry.beforeData;
    const afterData = entry.afterData;
    const beforeHeaders = entry.beforeHeaders;
    const afterHeaders = entry.afterHeaders;
    
    // Re-apply the changes
    for (let r = 0; r < Math.min(afterData.length, workbookData.length); r++) {
        for (let c = 0; c < Math.min(afterHeaders.length, headers.length); c++) {
            const beforeVal = String(beforeData[r]?.[c] ?? '');
            const afterVal = String(afterData[r]?.[c] ?? '');
            const currentVal = String(workbookData[r]?.[c] ?? '');
            
            // If this cell was changed by this AI call, and current value matches the original
            if (beforeVal !== afterVal && currentVal === beforeVal) {
                workbookData[r][c] = afterData[r][c];
            }
        }
    }
    
    entry.isReverted = false;
    
    renderTable();
    renderAIHistory();
    addChatMessage('system', t('reappliedAIChange', { description: entry.description }));
    updateButtonStates();
}

function viewAIDiff(index) {
    const entry = aiHistory[index];
    if (!entry) return;
    
    renderCompareTable('before', entry.beforeHeaders, entry.beforeData);
    renderCompareTable('after', entry.afterHeaders, entry.afterData, entry.beforeHeaders, entry.beforeData);
    
    aiHistoryModal.classList.add('hidden');
    compareModal.classList.remove('hidden');
}

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
    
    const headerRow = document.createElement('tr');
    const cornerCell = document.createElement('th');
    cornerCell.className = 'row-header';
    cornerCell.textContent = '#';
    headerRow.appendChild(cornerCell);
    
    tableHeaders.forEach((header, i) => {
        const th = document.createElement('th');
        th.textContent = header || `Col ${i + 1}`;
        
        if (compareHeaders && compareHeaders[i] !== header) {
            th.classList.add('changed');
        }
        headerRow.appendChild(th);
    });
    thead.appendChild(headerRow);
    
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

// Enhanced download with format selection
downloadBtn.addEventListener('click', () => {
    const format = exportFormatSelect.value;
    const exportName = originalFileName || 'spreadsheet';
    
    switch (format) {
        case 'xlsx':
            exportXLSX(exportName);
            break;
        case 'xls':
            exportXLS(exportName);
            break;
        case 'csv':
            exportCSV(exportName);
            break;
        case 'tsv':
            exportTSV(exportName);
            break;
        case 'json':
            exportJSON(exportName);
            break;
        case 'xml':
            exportXML(exportName);
            break;
        case 'html':
            exportHTML(exportName);
            break;
        case 'ods':
            exportODS(exportName);
            break;
        default:
            exportXLSX(exportName);
    }
    
    addChatMessage('system', t('downloaded'));
});

function exportXLSX(filename) {
    const ws = XLSX.utils.aoa_to_sheet([headers, ...workbookData]);
    
    for (let cell in cellColors) {
        if (ws[cell]) {
            ws[cell].s = {fgColor: {rgb: cellColors[cell].slice(1)}};
        }
    }
    
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
    XLSX.writeFile(wb, `${filename}_modified.xlsx`);
}

function exportXLS(filename) {
    const ws = XLSX.utils.aoa_to_sheet([headers, ...workbookData]);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
    XLSX.writeFile(wb, `${filename}_modified.xls`, { bookType: 'xls' });
}

function exportCSV(filename) {
    const ws = XLSX.utils.aoa_to_sheet([headers, ...workbookData]);
    const csv = XLSX.utils.sheet_to_csv(ws);
    downloadText(csv, `${filename}_modified.csv`, 'text/csv');
}

function exportTSV(filename) {
    const lines = [headers.join('\t')];
    workbookData.forEach(row => {
        lines.push(row.join('\t'));
    });
    downloadText(lines.join('\n'), `${filename}_modified.tsv`, 'text/tab-separated-values');
}

function exportJSON(filename) {
    const data = workbookData.map(row => {
        const obj = {};
        headers.forEach((h, i) => {
            obj[h || `Column${i+1}`] = row[i];
        });
        return obj;
    });
    downloadText(JSON.stringify(data, null, 2), `${filename}_modified.json`, 'application/json');
}

function exportXML(filename) {
    let xml = '<?xml version="1.0" encoding="UTF-8"?>\n<data>\n';
    workbookData.forEach((row, rowIdx) => {
        xml += '  <record>\n';
        headers.forEach((h, i) => {
            const tagName = (h || `column${i+1}`).replace(/[^a-zA-Z0-9_]/g, '_');
            const value = escapeXML(String(row[i] ?? ''));
            xml += `    <${tagName}>${value}</${tagName}>\n`;
        });
        xml += '  </record>\n';
    });
    xml += '</data>';
    downloadText(xml, `${filename}_modified.xml`, 'application/xml');
}

function escapeXML(str) {
    return str
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&apos;');
}

function exportHTML(filename) {
    let html = `<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>${filename}</title>
    <style>
        table { border-collapse: collapse; width: 100%; }
        th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
        th { background-color: #2d3e50; color: white; }
        tr:nth-child(even) { background-color: #f2f2f2; }
    </style>
</head>
<body>
    <table>
        <thead>
            <tr>${headers.map(h => `<th>${escapeHTML(h)}</th>`).join('')}</tr>
        </thead>
        <tbody>
`;
    workbookData.forEach(row => {
        html += `            <tr>${row.map(cell => `<td>${escapeHTML(String(cell ?? ''))}</td>`).join('')}</tr>\n`;
    });
    html += `        </tbody>
    </table>
</body>
</html>`;
    downloadText(html, `${filename}_modified.html`, 'text/html');
}

function escapeHTML(str) {
    return str
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;');
}

function exportODS(filename) {
    const ws = XLSX.utils.aoa_to_sheet([headers, ...workbookData]);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
    XLSX.writeFile(wb, `${filename}_modified.ods`, { bookType: 'ods' });
}

function downloadText(content, filename, mimeType) {
    const blob = new Blob([content], { type: mimeType + ';charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
}

aiPrompt.addEventListener('keydown', (e) => {
    if (e.key === 'Enter' && !e.shiftKey) {
        e.preventDefault();
        aiButton.click();
    }
});

updateButtonStates();
