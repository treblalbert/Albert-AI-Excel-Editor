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

// Processing control
let isProcessing = false;
let shouldStopProcessing = false;
let preProcessingState = null;

// Selection state - simple bounds-based
let selectionStart = null;
let selectionEnd = null;
let isSelecting = false;
let isDraggingFill = false;
let fillEnd = null;
let editingCell = null;
let clipboard = [];

// DOM Elements
const apiSetup = document.getElementById('apiSetup');
const mainApp = document.getElementById('mainApp');
const apiKeyFile = document.getElementById('apiKeyFile');
const apiKeyFileLater = document.getElementById('apiKeyFileLater');
const skipAiBtn = document.getElementById('skipAiBtn');
const noAiOverlay = document.getElementById('noAiOverlay');
const excelFile = document.getElementById('excelFile');
const aiPrompt = document.getElementById('aiPrompt');
const aiButton = document.getElementById('aiButton');
const undoBtn = document.getElementById('undoBtn');
const stopBtn = document.getElementById('stopBtn');
const addRowBtn = document.getElementById('addRowBtn');
const addColBtn = document.getElementById('addColBtn');
const compareBtn = document.getElementById('compareBtn');
const downloadBtn = document.getElementById('downloadBtn');
const exportFormat = document.getElementById('exportFormat');
const statusDiv = document.getElementById('status');
const tableHead = document.getElementById('tableHead');
const tableBody = document.getElementById('tableBody');
const chatMessages = document.getElementById('chatMessages');
const highlightToggle = document.getElementById('highlightToggle');
const compareModal = document.getElementById('compareModal');
const closeCompareModal = document.getElementById('closeCompareModal');
const themeToggle = document.getElementById('themeToggle');

// AI Mode flag
let aiModeEnabled = false;

// Theme
function initTheme() {
    const savedTheme = localStorage.getItem('theme') || 'light';
    document.documentElement.setAttribute('data-theme', savedTheme);
    updateThemeIcon(savedTheme);
}

function toggleTheme() {
    const current = document.documentElement.getAttribute('data-theme');
    const next = current === 'dark' ? 'light' : 'dark';
    document.documentElement.setAttribute('data-theme', next);
    localStorage.setItem('theme', next);
    updateThemeIcon(next);
}

function updateThemeIcon(theme) {
    const icon = document.getElementById('themeIcon');
    const text = document.getElementById('themeText');
    if (icon) icon.textContent = theme === 'dark' ? '☀️' : '🌙';
    if (text) text.textContent = theme === 'dark' ? t('lightModeLabel') : t('darkModeLabel');
}

// Init
window.addEventListener('load', () => {
    initTheme();
    const savedHighlight = localStorage.getItem('highlightEnabled');
    const savedFormat = localStorage.getItem('exportFormat');
    if (savedHighlight === 'true') { highlightToggle.checked = true; highlightEnabled = true; }
    if (savedFormat && exportFormat) exportFormat.value = savedFormat;
    const savedPath = localStorage.getItem('apiKeyPath');
    if (savedPath) {
        addChatMessage('system', t('attemptingLoad'));
        addChatMessage('system', t('previouslyUsed', { path: savedPath }));
    }
});

if (themeToggle) themeToggle.addEventListener('click', toggleTheme);
highlightToggle.addEventListener('change', (e) => {
    highlightEnabled = e.target.checked;
    localStorage.setItem('highlightEnabled', highlightEnabled);
    renderTable();
});
if (exportFormat) exportFormat.addEventListener('change', (e) => localStorage.setItem('exportFormat', e.target.value));

function showStatus(msg, type = 'info') {
    statusDiv.className = `status ${type}`;
    statusDiv.textContent = msg;
    statusDiv.style.display = 'block';
    setTimeout(() => statusDiv.style.display = 'none', 5000);
}

function addChatMessage(type, content) {
    const div = document.createElement('div');
    div.className = `message ${type}`;
    div.textContent = content;
    chatMessages.appendChild(div);
    chatMessages.scrollTop = chatMessages.scrollHeight;
}

function showToast(text) {
    let toast = document.querySelector('.toast-indicator');
    if (!toast) { toast = document.createElement('div'); toast.className = 'toast-indicator'; document.body.appendChild(toast); }
    toast.textContent = text; toast.classList.add('show');
    setTimeout(() => toast.classList.remove('show'), 1500);
}

// File handling
apiKeyFile.addEventListener('change', async (e) => {
    const file = e.target.files[0];
    if (!file) return;
    apiKey = (await file.text()).trim();
    apiKeyPath = file.name;
    localStorage.setItem('apiKeyPath', file.name);
    addChatMessage('system', t('apiKeyLoaded'));
    aiModeEnabled = true;
    apiSetup.classList.add('hidden');
    mainApp.classList.remove('hidden');
    noAiOverlay.classList.add('hidden');
    updateButtonStates();
});

// Skip AI - continue without AI features
skipAiBtn.addEventListener('click', () => {
    aiModeEnabled = false;
    apiSetup.classList.add('hidden');
    mainApp.classList.remove('hidden');
    noAiOverlay.classList.remove('hidden');
    addChatMessage('system', t('noAiModeActive') || 'Editor mode without AI features. You can load an API key anytime to enable AI.');
    updateButtonStates();
});

// Load API key later (from chat overlay)
apiKeyFileLater.addEventListener('change', async (e) => {
    const file = e.target.files[0];
    if (!file) return;
    apiKey = (await file.text()).trim();
    apiKeyPath = file.name;
    localStorage.setItem('apiKeyPath', file.name);
    aiModeEnabled = true;
    noAiOverlay.classList.add('hidden');
    addChatMessage('system', t('apiKeyLoaded'));
    updateButtonStates();
});

excelFile.addEventListener('change', async (e) => {
    const file = e.target.files[0];
    if (!file) return;
    try {
        originalFileName = file.name.replace(/\.[^/.]+$/, '');
        const data = await file.arrayBuffer();
        const wb = XLSX.read(data, { type: 'array', cellStyles: true });
        const ws = wb.Sheets[wb.SheetNames[0]];
        cellColors = {};
        for (let c in ws) {
            if (c[0] !== '!' && ws[c].s && ws[c].s.fgColor) cellColors[c] = '#' + ws[c].s.fgColor.rgb;
        }
        const json = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
        if (!json.length) { showStatus(t('errorEmpty'), 'error'); return; }
        headers = json[0].map(h => h != null ? String(h) : '');
        workbookData = json.slice(1).map(r => r.map(c => c != null ? c : ''));
        history = []; activeFilters = {}; modifiedCells.clear(); lastAIChange = null;
        selectionStart = selectionEnd = editingCell = null;
        renderTable();
        addChatMessage('system', t('fileLoaded', { rows: workbookData.length, cols: headers.length, filename: file.name }));
        updateButtonStates();
    } catch (err) {
        showStatus(t('errorLoading', { error: err.message }), 'error');
        addChatMessage('system', t('apiError', { error: err.message }));
    }
});

function saveState(actionType = 'general') {
    history.push({ 
        headers: JSON.parse(JSON.stringify(headers)), 
        data: JSON.parse(JSON.stringify(workbookData)), 
        colors: JSON.parse(JSON.stringify(cellColors)), 
        modifiedCells: new Set(modifiedCells),
        actionType: actionType
    });
    if (history.length > 50) history.shift();
    markUnsavedChanges();
    updateButtonStates();
}

function getSelectionBounds() {
    if (!selectionStart) return null;
    const e = selectionEnd || selectionStart;
    return { r1: Math.min(selectionStart.row, e.row), r2: Math.max(selectionStart.row, e.row), c1: Math.min(selectionStart.col, e.col), c2: Math.max(selectionStart.col, e.col) };
}
function renderTable() {
    tableHead.innerHTML = ''; tableBody.innerHTML = '';
    const hr = document.createElement('tr');
    const corner = document.createElement('th'); corner.className = 'row-header'; corner.textContent = '#'; hr.appendChild(corner);
    headers.forEach((h, i) => {
        const th = document.createElement('th');
        th.style.width = columnWidths[i] || '120px';
        const hc = document.createElement('div'); hc.className = 'header-content';
        
        // Header text with sort indicator
        const ht = document.createElement('span'); 
        ht.className = 'header-text sortable'; 
        let headerText = h || `Col ${i+1}`;
        if (sortState.column === i) {
            headerText += sortState.direction === 'asc' ? ' ↑' : ' ↓';
        }
        ht.textContent = headerText;
        ht.onclick = (e) => { e.stopPropagation(); sortByColumn(i); };
        ht.title = t('clickToSort') || 'Click to sort';
        hc.appendChild(ht);
        
        const fi = document.createElement('span'); fi.className = 'filter-icon' + (hasActiveFilter(i) ? ' active' : ''); fi.textContent = '▼';
        fi.onclick = (e) => { e.stopPropagation(); showFilterDropdown(e, i); }; hc.appendChild(fi);
        th.appendChild(hc);
        const rh = document.createElement('div'); rh.className = 'resize-handle';
        rh.addEventListener('mousedown', (e) => startResize(e, i, th)); th.appendChild(rh);
        hr.appendChild(th);
    });
    tableHead.appendChild(hr);

    const bounds = getSelectionBounds();
    applyFilters().forEach((row) => {
        const tr = document.createElement('tr');
        const rh = document.createElement('td'); rh.className = 'row-header'; rh.textContent = row.originalIndex + 1; tr.appendChild(rh);
        headers.forEach((_, ci) => {
            const ri = row.originalIndex;
            const td = document.createElement('td');
            td.dataset.row = ri; td.dataset.col = ci;
            const cellRef = XLSX.utils.encode_cell({r: ri + 1, c: ci});
            if (cellColors[cellRef]) td.style.backgroundColor = cellColors[cellRef];
            if (highlightEnabled && modifiedCells.has(`${ri}-${ci}`)) td.classList.add('ai-modified');
            if (bounds && ri >= bounds.r1 && ri <= bounds.r2 && ci >= bounds.c1 && ci <= bounds.c2) td.classList.add('selected');
            
            if (editingCell?.row === ri && editingCell?.col === ci) {
                const inp = document.createElement('input'); inp.type = 'text'; inp.className = 'cell-input';
                inp.value = row.data[ci] ?? '';
                inp.addEventListener('blur', () => finishEditing(ri, ci, inp.value));
                inp.addEventListener('keydown', (e) => handleEditKey(e, ri, ci, inp));
                td.appendChild(inp);
                setTimeout(() => { inp.focus(); inp.select(); }, 0);
            } else {
                const d = document.createElement('div'); d.className = 'cell-display'; d.textContent = row.data[ci] ?? ''; td.appendChild(d);
            }
            if (bounds && ri === bounds.r2 && ci === bounds.c2 && !editingCell) {
                const fh = document.createElement('div'); fh.className = 'fill-handle'; td.appendChild(fh);
            }
            tr.appendChild(td);
        });
        tableBody.appendChild(tr);
    });
    tableBody.onmousedown = onTableMouseDown;
    tableBody.ondblclick = onTableDblClick;
}

function onTableMouseDown(e) {
    const td = e.target.closest('td[data-row]');
    if (!td) return;
    const r = +td.dataset.row, c = +td.dataset.col;
    if (e.target.classList.contains('fill-handle')) { isDraggingFill = true; fillEnd = {row:r,col:c}; document.addEventListener('mousemove', onDocMouseMove); document.addEventListener('mouseup', onDocMouseUp); e.preventDefault(); return; }
    if (e.target.classList.contains('cell-input')) return;
    if (editingCell) { const inp = document.querySelector('.cell-input'); if (inp) finishEditing(editingCell.row, editingCell.col, inp.value); }
    e.preventDefault();
    if (e.shiftKey && selectionStart) { selectionEnd = {row:r,col:c}; } else { selectionStart = {row:r,col:c}; selectionEnd = {row:r,col:c}; }
    isSelecting = true;
    updateSelectionVisual();
    document.addEventListener('mousemove', onDocMouseMove);
    document.addEventListener('mouseup', onDocMouseUp);
}

function onDocMouseMove(e) {
    const td = document.elementFromPoint(e.clientX, e.clientY)?.closest('td[data-row]');
    if (!td) return;
    const r = +td.dataset.row, c = +td.dataset.col;
    if (isSelecting) { selectionEnd = {row:r,col:c}; updateSelectionVisual(); }
    else if (isDraggingFill) { fillEnd = {row:r,col:c}; updateFillPreview(); }
}

function onDocMouseUp() {
    document.removeEventListener('mousemove', onDocMouseMove);
    document.removeEventListener('mouseup', onDocMouseUp);
    if (isDraggingFill) { performFill(); isDraggingFill = false; clearFillPreview(); }
    isSelecting = false;
}

function onTableDblClick(e) {
    const td = e.target.closest('td[data-row]');
    if (td) startEditing(+td.dataset.row, +td.dataset.col);
}

function updateSelectionVisual() {
    document.querySelectorAll('td.selected').forEach(t => t.classList.remove('selected'));
    document.querySelectorAll('.fill-handle').forEach(f => f.remove());
    const b = getSelectionBounds(); if (!b) return;
    for (let r = b.r1; r <= b.r2; r++) for (let c = b.c1; c <= b.c2; c++) {
        const td = document.querySelector(`td[data-row="${r}"][data-col="${c}"]`);
        if (td) td.classList.add('selected');
    }
    if (!editingCell) {
        const last = document.querySelector(`td[data-row="${b.r2}"][data-col="${b.c2}"]`);
        if (last && !last.querySelector('.fill-handle')) { const fh = document.createElement('div'); fh.className = 'fill-handle'; last.appendChild(fh); }
    }
}

function updateFillPreview() {
    clearFillPreview();
    const b = getSelectionBounds(); if (!b || !fillEnd) return;
    const tr = fillEnd.row, tc = fillEnd.col;
    if (tc >= b.c1 && tc <= b.c2) {
        if (tr > b.r2) for (let r = b.r2+1; r <= tr; r++) for (let c = b.c1; c <= b.c2; c++) document.querySelector(`td[data-row="${r}"][data-col="${c}"]`)?.classList.add('fill-preview');
        if (tr < b.r1) for (let r = tr; r < b.r1; r++) for (let c = b.c1; c <= b.c2; c++) document.querySelector(`td[data-row="${r}"][data-col="${c}"]`)?.classList.add('fill-preview');
    }
    if (tr >= b.r1 && tr <= b.r2) {
        if (tc > b.c2) for (let r = b.r1; r <= b.r2; r++) for (let c = b.c2+1; c <= tc; c++) document.querySelector(`td[data-row="${r}"][data-col="${c}"]`)?.classList.add('fill-preview');
        if (tc < b.c1) for (let r = b.r1; r <= b.r2; r++) for (let c = tc; c < b.c1; c++) document.querySelector(`td[data-row="${r}"][data-col="${c}"]`)?.classList.add('fill-preview');
    }
}

function clearFillPreview() { document.querySelectorAll('td.fill-preview').forEach(t => t.classList.remove('fill-preview')); }

function performFill() {
    const b = getSelectionBounds(); if (!b || !fillEnd) return;
    const src = [];
    for (let r = b.r1; r <= b.r2; r++) { const row = []; for (let c = b.c1; c <= b.c2; c++) row.push(workbookData[r]?.[c] ?? ''); src.push(row); }
    
    // Save state BEFORE making changes - enables Ctrl+Z for fill operations
    saveState('fill');
    
    let cnt = 0;
    const tr = fillEnd.row, tc = fillEnd.col;
    if (tr > b.r2 && tc >= b.c1 && tc <= b.c2) {
        for (let r = b.r2+1; r <= tr; r++) for (let c = b.c1; c <= b.c2; c++) {
            const si = (r - b.r2 - 1) % src.length, ci = c - b.c1;
            if (!workbookData[r]) workbookData[r] = new Array(headers.length).fill('');
            workbookData[r][c] = detectPattern(src, ci, r - b.r1); cnt++;
        }
    }
    if (tr < b.r1 && tc >= b.c1 && tc <= b.c2) {
        for (let r = b.r1-1; r >= tr; r--) for (let c = b.c1; c <= b.c2; c++) {
            const si = src.length - 1 - ((b.r1 - 1 - r) % src.length), ci = c - b.c1;
            if (!workbookData[r]) workbookData[r] = new Array(headers.length).fill('');
            workbookData[r][c] = src[si][ci]; cnt++;
        }
    }
    if (tc > b.c2 && tr >= b.r1 && tr <= b.r2) {
        for (let c = b.c2+1; c <= tc; c++) for (let r = b.r1; r <= b.r2; r++) {
            const si = (c - b.c2 - 1) % (b.c2 - b.c1 + 1), ri = r - b.r1;
            workbookData[r][c] = src[ri][si]; cnt++;
        }
    }
    if (tc < b.c1 && tr >= b.r1 && tr <= b.r2) {
        for (let c = b.c1-1; c >= tc; c--) for (let r = b.r1; r <= b.r2; r++) {
            const si = (b.c1 - 1 - c) % (b.c2 - b.c1 + 1), ri = r - b.r1;
            workbookData[r][c] = src[ri][b.c2 - b.c1 - si]; cnt++;
        }
    }
    
    if (cnt) { 
        showToast(t('cellsFilled', { count: cnt })); 
        renderTable(); 
    } else {
        // No changes made, remove saved state
        history.pop();
        updateButtonStates();
    }
}

function detectPattern(src, ci, offset) {
    if (src.length >= 2) {
        const v1 = parseFloat(src[0][ci]), v2 = parseFloat(src[1][ci]);
        if (!isNaN(v1) && !isNaN(v2)) return v1 + (v2 - v1) * offset;
    }
    return src[offset % src.length][ci];
}

function startEditing(r, c) { editingCell = {row:r,col:c}; selectionStart = {row:r,col:c}; selectionEnd = {row:r,col:c}; renderTable(); }

function finishEditing(r, c, val) {
    if (!editingCell || editingCell.row !== r || editingCell.col !== c) return;
    if (String(workbookData[r][c]) !== String(val)) { saveState('edit'); workbookData[r][c] = val; }
    editingCell = null; renderTable();
}

function handleEditKey(e, r, c, inp) {
    if (e.key === 'Enter') { e.preventDefault(); finishEditing(r, c, inp.value); if (r < workbookData.length - 1) { selectionStart = selectionEnd = {row:r+1,col:c}; } renderTable(); }
    else if (e.key === 'Escape') { editingCell = null; renderTable(); }
    else if (e.key === 'Tab') { e.preventDefault(); finishEditing(r, c, inp.value); const nc = e.shiftKey ? c-1 : c+1; if (nc >= 0 && nc < headers.length) { selectionStart = selectionEnd = {row:r,col:nc}; } renderTable(); }
}
// Keyboard shortcuts
document.addEventListener('keydown', (e) => {
    if (editingCell || e.target.tagName === 'INPUT' || e.target.tagName === 'TEXTAREA') return;
    const b = getSelectionBounds();
    if (['ArrowUp','ArrowDown','ArrowLeft','ArrowRight'].includes(e.key) && b) {
        e.preventDefault();
        let {row,col} = selectionStart;
        if (e.key === 'ArrowUp') row = Math.max(0, row-1);
        if (e.key === 'ArrowDown') row = Math.min(workbookData.length-1, row+1);
        if (e.key === 'ArrowLeft') col = Math.max(0, col-1);
        if (e.key === 'ArrowRight') col = Math.min(headers.length-1, col+1);
        if (e.shiftKey) selectionEnd = {row,col}; else selectionStart = selectionEnd = {row,col};
        updateSelectionVisual();
        document.querySelector(`td[data-row="${row}"][data-col="${col}"]`)?.scrollIntoView({block:'nearest',inline:'nearest'});
        return;
    }
    if ((e.key === 'Delete' || e.key === 'Backspace') && b) { e.preventDefault(); saveState('delete'); for (let r = b.r1; r <= b.r2; r++) for (let c = b.c1; c <= b.c2; c++) workbookData[r][c] = ''; renderTable(); return; }
    if (e.key === 'Enter' && b && b.r1 === b.r2 && b.c1 === b.c2) { e.preventDefault(); startEditing(b.r1, b.c1); return; }
    if (b && b.r1 === b.r2 && b.c1 === b.c2 && e.key.length === 1 && !e.ctrlKey && !e.metaKey) { workbookData[b.r1][b.c1] = ''; startEditing(b.r1, b.c1); return; }
    if ((e.ctrlKey || e.metaKey) && e.key === 'c' && b) { e.preventDefault(); clipboard = []; for (let r = b.r1; r <= b.r2; r++) { const row = []; for (let c = b.c1; c <= b.c2; c++) row.push(workbookData[r]?.[c] ?? ''); clipboard.push(row); } navigator.clipboard.writeText(clipboard.map(r=>r.join('\t')).join('\n')).catch(()=>{}); showToast(t('cellsCopied',{count:(b.r2-b.r1+1)*(b.c2-b.c1+1)})); return; }
    if ((e.ctrlKey || e.metaKey) && e.key === 'x' && b) { e.preventDefault(); clipboard = []; saveState('cut'); for (let r = b.r1; r <= b.r2; r++) { const row = []; for (let c = b.c1; c <= b.c2; c++) { row.push(workbookData[r]?.[c] ?? ''); workbookData[r][c] = ''; } clipboard.push(row); } renderTable(); showToast(t('copied')); return; }
    if ((e.ctrlKey || e.metaKey) && e.key === 'v' && b && clipboard.length) { e.preventDefault(); saveState('paste'); for (let r = 0; r < clipboard.length; r++) for (let c = 0; c < clipboard[r].length; c++) { const tr = b.r1+r, tc = b.c1+c; if (tr < workbookData.length && tc < headers.length) workbookData[tr][tc] = clipboard[r][c]; } renderTable(); showToast(t('pasted')); return; }
    if ((e.ctrlKey || e.metaKey) && e.key === 'a' && workbookData.length) { e.preventDefault(); selectionStart = {row:0,col:0}; selectionEnd = {row:workbookData.length-1,col:headers.length-1}; updateSelectionVisual(); return; }
    if ((e.ctrlKey || e.metaKey) && e.key === 'z' && history.length) { e.preventDefault(); undoBtn.click(); return; }
    if (e.key === 'Escape') { 
        if (!compareModal.classList.contains('hidden')) compareModal.classList.add('hidden'); 
        else if (searchDialog && !searchDialog.classList.contains('hidden')) closeSearch(); 
        else if (replaceDialog && !replaceDialog.classList.contains('hidden')) closeReplace(); 
        else { selectionStart = selectionEnd = null; updateSelectionVisual(); } 
        return; 
    }
    if ((e.ctrlKey || e.metaKey) && e.key === 'f') { e.preventDefault(); openSearch(); return; }
    if ((e.ctrlKey || e.metaKey) && e.key === 'h') { e.preventDefault(); openReplace(); return; }
    if ((e.ctrlKey || e.metaKey) && e.key === 's') { e.preventDefault(); openSaveDialog(); return; }
});

// =====================================
// ADVANCED FILTER SYSTEM
// =====================================

function hasActiveFilter(colIndex) {
    const f = activeFilters[colIndex];
    if (!f) return false;
    if (f.type === 'values' && f.values?.size > 0) return true;
    if (f.type === 'text' && f.value) return true;
    if (f.type === 'number' && (f.value1 !== '' || f.value2 !== '')) return true;
    return false;
}

function applyFilters() {
    let f = workbookData.map((d,i)=>({data:d,originalIndex:i}));
    for (let ci in activeFilters) {
        const filter = activeFilters[ci];
        if (!filter) continue;
        
        if (filter.type === 'values' && filter.values?.size > 0) {
            f = f.filter(r => filter.values.has(String(r.data[ci] ?? '')));
        } else if (filter.type === 'text' && filter.value) {
            const val = filter.value.toLowerCase();
            f = f.filter(r => {
                const cellVal = String(r.data[ci] ?? '').toLowerCase();
                switch(filter.operator) {
                    case 'contains': return cellVal.includes(val);
                    case 'notcontains': return !cellVal.includes(val);
                    case 'startswith': return cellVal.startsWith(val);
                    case 'endswith': return cellVal.endsWith(val);
                    case 'equals': return cellVal === val;
                    case 'notequals': return cellVal !== val;
                    case 'empty': return cellVal === '';
                    case 'notempty': return cellVal !== '';
                    default: return true;
                }
            });
        } else if (filter.type === 'number') {
            const v1 = parseFloat(filter.value1);
            const v2 = parseFloat(filter.value2);
            f = f.filter(r => {
                const cellVal = parseFloat(r.data[ci]);
                if (isNaN(cellVal)) return filter.operator === 'empty';
                switch(filter.operator) {
                    case 'equals': return !isNaN(v1) && cellVal === v1;
                    case 'notequals': return !isNaN(v1) && cellVal !== v1;
                    case 'greater': return !isNaN(v1) && cellVal > v1;
                    case 'greaterequal': return !isNaN(v1) && cellVal >= v1;
                    case 'less': return !isNaN(v1) && cellVal < v1;
                    case 'lessequal': return !isNaN(v1) && cellVal <= v1;
                    case 'between': return !isNaN(v1) && !isNaN(v2) && cellVal >= v1 && cellVal <= v2;
                    case 'empty': return false;
                    case 'notempty': return true;
                    default: return true;
                }
            });
        }
    }
    return f;
}

function showFilterDropdown(e, ci) {
    e.stopPropagation(); 
    document.querySelectorAll('.filter-dropdown').forEach(d => d.remove());
    
    const dd = document.createElement('div'); 
    dd.className = 'filter-dropdown filter-dropdown-advanced';
    
    // Filter type tabs
    const tabs = document.createElement('div');
    tabs.className = 'filter-tabs';
    const currentFilter = activeFilters[ci];
    const currentType = currentFilter?.type || 'values';
    
    const tabData = [
        { id: 'values', label: t('filterByValues') || 'Values' },
        { id: 'text', label: t('filterByText') || 'Text' },
        { id: 'number', label: t('filterByNumber') || 'Number' }
    ];
    
    tabData.forEach(tab => {
        const btn = document.createElement('button');
        btn.className = 'filter-tab' + (currentType === tab.id ? ' active' : '');
        btn.textContent = tab.label;
        btn.onclick = (ev) => { ev.stopPropagation(); switchFilterTab(dd, ci, tab.id); };
        tabs.appendChild(btn);
    });
    dd.appendChild(tabs);
    
    const content = document.createElement('div');
    content.className = 'filter-content';
    content.id = 'filterContent';
    dd.appendChild(content);
    
    const clearBtn = document.createElement('button');
    clearBtn.className = 'filter-clear-btn';
    clearBtn.textContent = t('clearFilter') || 'Clear Filter';
    clearBtn.onclick = (ev) => {
        ev.stopPropagation();
        delete activeFilters[ci];
        renderTable();
        dd.remove();
    };
    dd.appendChild(clearBtn);
    
    e.target.closest('th').appendChild(dd);
    renderFilterContent(content, ci, currentType);
    
    const closeHandler = (evt) => {
        if (!dd.contains(evt.target) && !evt.target.classList.contains('filter-icon')) {
            dd.remove();
            document.removeEventListener('click', closeHandler);
        }
    };
    setTimeout(() => document.addEventListener('click', closeHandler), 0);
}

function switchFilterTab(dd, ci, type) {
    dd.querySelectorAll('.filter-tab').forEach(t => t.classList.remove('active'));
    const index = type === 'values' ? 1 : type === 'text' ? 2 : 3;
    dd.querySelector(`.filter-tab:nth-child(${index})`).classList.add('active');
    renderFilterContent(dd.querySelector('#filterContent'), ci, type);
}

function renderFilterContent(container, ci, type) {
    container.innerHTML = '';
    if (type === 'values') renderValuesFilter(container, ci);
    else if (type === 'text') renderTextFilter(container, ci);
    else if (type === 'number') renderNumberFilter(container, ci);
}

function renderValuesFilter(container, ci) {
    const si = document.createElement('input'); 
    si.type = 'text'; 
    si.placeholder = t('searchPlaceholder') || 'Search...'; 
    si.className = 'filter-search';
    container.appendChild(si);
    
    const ad = document.createElement('div'); 
    ad.className = 'filter-actions';
    const ab = document.createElement('button'); 
    ab.textContent = t('selectAll'); 
    ab.onclick = (ev) => { 
        ev.stopPropagation();
        delete activeFilters[ci]; 
        od.querySelectorAll('input[type=checkbox]').forEach(cb => cb.checked = true); 
        renderTable(); 
    }; 
    ad.appendChild(ab);
    const cb = document.createElement('button'); 
    cb.textContent = t('clearAll'); 
    cb.onclick = (ev) => { 
        ev.stopPropagation();
        activeFilters[ci] = { type: 'values', values: new Set() }; 
        od.querySelectorAll('input[type=checkbox]').forEach(cb => cb.checked = false); 
        renderTable(); 
    }; 
    ad.appendChild(cb);
    container.appendChild(ad);
    
    const od = document.createElement('div'); 
    od.className = 'filter-options';
    const uv = [...new Set(workbookData.map(r => String(r[ci] ?? '')))].sort((a, b) => {
        const na = parseFloat(a), nb = parseFloat(b);
        return !isNaN(na) && !isNaN(nb) ? na - nb : a.localeCompare(b);
    });
    
    const cf = activeFilters[ci];
    const hf = cf?.type === 'values' && cf.values?.size > 0;
    
    uv.forEach(v => {
        const o = document.createElement('div'); 
        o.className = 'filter-option';
        const chk = document.createElement('input'); 
        chk.type = 'checkbox'; 
        chk.checked = !hf || cf.values.has(v);
        chk.onchange = () => { 
            if (!activeFilters[ci] || activeFilters[ci].type !== 'values') {
                activeFilters[ci] = { type: 'values', values: new Set(uv) };
            }
            if (chk.checked) activeFilters[ci].values.add(v); 
            else activeFilters[ci].values.delete(v); 
            if (activeFilters[ci].values.size === uv.length) delete activeFilters[ci]; 
            renderTable(); 
        };
        const l = document.createElement('span'); 
        l.textContent = v === '' ? '(blank)' : v;
        o.appendChild(chk); 
        o.appendChild(l); 
        od.appendChild(o);
    });
    container.appendChild(od);
    
    si.oninput = () => { 
        const s = si.value.toLowerCase(); 
        od.querySelectorAll('.filter-option').forEach(o => {
            o.style.display = o.textContent.toLowerCase().includes(s) ? 'flex' : 'none';
        }); 
    };
}

function renderTextFilter(container, ci) {
    const currentFilter = activeFilters[ci]?.type === 'text' ? activeFilters[ci] : null;
    
    const selectDiv = document.createElement('div');
    selectDiv.className = 'filter-select-group';
    const select = document.createElement('select');
    select.className = 'filter-select';
    const options = [
        { value: 'contains', label: t('filterContains') || 'Contains' },
        { value: 'notcontains', label: t('filterNotContains') || 'Does not contain' },
        { value: 'startswith', label: t('filterStartsWith') || 'Starts with' },
        { value: 'endswith', label: t('filterEndsWith') || 'Ends with' },
        { value: 'equals', label: t('filterEquals') || 'Equals' },
        { value: 'notequals', label: t('filterNotEquals') || 'Does not equal' },
        { value: 'empty', label: t('filterEmpty') || 'Is empty' },
        { value: 'notempty', label: t('filterNotEmpty') || 'Is not empty' }
    ];
    options.forEach(opt => {
        const o = document.createElement('option');
        o.value = opt.value;
        o.textContent = opt.label;
        if (currentFilter?.operator === opt.value) o.selected = true;
        select.appendChild(o);
    });
    selectDiv.appendChild(select);
    container.appendChild(selectDiv);
    
    const input = document.createElement('input');
    input.type = 'text';
    input.className = 'filter-input';
    input.placeholder = t('filterValuePlaceholder') || 'Enter value...';
    input.value = currentFilter?.value || '';
    if (select.value === 'empty' || select.value === 'notempty') input.style.display = 'none';
    container.appendChild(input);
    
    select.onchange = () => {
        input.style.display = (select.value === 'empty' || select.value === 'notempty') ? 'none' : 'block';
        applyTextFilter();
    };
    
    const applyBtn = document.createElement('button');
    applyBtn.className = 'filter-apply-btn';
    applyBtn.textContent = t('applyFilter') || 'Apply';
    applyBtn.onclick = (ev) => { ev.stopPropagation(); applyTextFilter(); };
    container.appendChild(applyBtn);
    
    function applyTextFilter() {
        const op = select.value;
        const val = input.value;
        if (op === 'empty' || op === 'notempty' || val) {
            activeFilters[ci] = { type: 'text', operator: op, value: val };
            renderTable();
        }
    }
    
    input.addEventListener('keydown', (e) => { if (e.key === 'Enter') applyTextFilter(); });
}

function renderNumberFilter(container, ci) {
    const currentFilter = activeFilters[ci]?.type === 'number' ? activeFilters[ci] : null;
    
    const selectDiv = document.createElement('div');
    selectDiv.className = 'filter-select-group';
    const select = document.createElement('select');
    select.className = 'filter-select';
    const options = [
        { value: 'equals', label: t('filterNumEquals') || 'Equals' },
        { value: 'notequals', label: t('filterNumNotEquals') || 'Does not equal' },
        { value: 'greater', label: t('filterGreater') || 'Greater than' },
        { value: 'greaterequal', label: t('filterGreaterEqual') || 'Greater or equal' },
        { value: 'less', label: t('filterLess') || 'Less than' },
        { value: 'lessequal', label: t('filterLessEqual') || 'Less or equal' },
        { value: 'between', label: t('filterBetween') || 'Between' },
        { value: 'empty', label: t('filterEmpty') || 'Is empty' },
        { value: 'notempty', label: t('filterNotEmpty') || 'Is not empty' }
    ];
    options.forEach(opt => {
        const o = document.createElement('option');
        o.value = opt.value;
        o.textContent = opt.label;
        if (currentFilter?.operator === opt.value) o.selected = true;
        select.appendChild(o);
    });
    selectDiv.appendChild(select);
    container.appendChild(selectDiv);
    
    const inputsDiv = document.createElement('div');
    inputsDiv.className = 'filter-inputs';
    
    const input1 = document.createElement('input');
    input1.type = 'number';
    input1.className = 'filter-input';
    input1.placeholder = t('filterValue') || 'Value';
    input1.value = currentFilter?.value1 || '';
    inputsDiv.appendChild(input1);
    
    const input2 = document.createElement('input');
    input2.type = 'number';
    input2.className = 'filter-input';
    input2.placeholder = t('filterValueTo') || 'To';
    input2.value = currentFilter?.value2 || '';
    input2.style.display = select.value === 'between' ? 'block' : 'none';
    inputsDiv.appendChild(input2);
    
    if (select.value === 'empty' || select.value === 'notempty') inputsDiv.style.display = 'none';
    container.appendChild(inputsDiv);
    
    select.onchange = () => {
        input2.style.display = select.value === 'between' ? 'block' : 'none';
        inputsDiv.style.display = (select.value === 'empty' || select.value === 'notempty') ? 'none' : 'block';
        applyNumberFilter();
    };
    
    const applyBtn = document.createElement('button');
    applyBtn.className = 'filter-apply-btn';
    applyBtn.textContent = t('applyFilter') || 'Apply';
    applyBtn.onclick = (ev) => { ev.stopPropagation(); applyNumberFilter(); };
    container.appendChild(applyBtn);
    
    function applyNumberFilter() {
        const op = select.value;
        const v1 = input1.value;
        const v2 = input2.value;
        if (op === 'empty' || op === 'notempty' || v1 !== '' || (op === 'between' && v2 !== '')) {
            activeFilters[ci] = { type: 'number', operator: op, value1: v1, value2: v2 };
            renderTable();
        }
    }
    
    input1.addEventListener('keydown', (e) => { if (e.key === 'Enter') applyNumberFilter(); });
    input2.addEventListener('keydown', (e) => { if (e.key === 'Enter') applyNumberFilter(); });
}
// =====================================
// SEARCH DIALOG (Ctrl+F)
// =====================================
let searchDialog = null, searchMatches = [], currentSearchIndex = -1;

function createSearchDialog() {
    if (searchDialog) return;
    searchDialog = document.createElement('div'); 
    searchDialog.className = 'search-dialog hidden';
    searchDialog.innerHTML = `
        <div class="search-content">
            <input type="text" id="searchInput" placeholder="${t('findPlaceholder') || 'Find...'}">
            <span id="searchCount"></span>
            <button id="searchPrev" title="Previous (Shift+Enter)">▲</button>
            <button id="searchNext" title="Next (Enter)">▼</button>
            <button id="searchClose" title="Close (Esc)">×</button>
        </div>
    `;
    document.body.appendChild(searchDialog);
    document.getElementById('searchInput').addEventListener('input', performSearch);
    document.getElementById('searchInput').addEventListener('keydown', (e) => { 
        if (e.key === 'Enter') { 
            e.shiftKey ? navigateSearchResult(-1) : navigateSearchResult(1); 
            e.preventDefault(); 
        } else if (e.key === 'Escape') closeSearch(); 
    });
    document.getElementById('searchPrev').addEventListener('click', () => navigateSearchResult(-1));
    document.getElementById('searchNext').addEventListener('click', () => navigateSearchResult(1));
    document.getElementById('searchClose').addEventListener('click', closeSearch);
}

function openSearch() { 
    closeReplace();
    createSearchDialog(); 
    searchDialog.classList.remove('hidden'); 
    document.getElementById('searchInput').focus(); 
    document.getElementById('searchInput').select(); 
}

function closeSearch() { 
    if (searchDialog) searchDialog.classList.add('hidden'); 
    searchMatches = []; 
    currentSearchIndex = -1; 
    const countEl = document.getElementById('searchCount');
    if (countEl) countEl.textContent = ''; 
    document.querySelectorAll('.search-match,.search-current').forEach(e => e.classList.remove('search-match','search-current')); 
}

function performSearch() {
    const q = document.getElementById('searchInput').value.toLowerCase().trim(); 
    searchMatches = []; 
    currentSearchIndex = -1;
    document.querySelectorAll('.search-match,.search-current').forEach(e => e.classList.remove('search-match','search-current'));
    if (!q) { document.getElementById('searchCount').textContent = ''; return; }
    workbookData.forEach((r, ri) => r.forEach((c, ci) => { 
        if (String(c).toLowerCase().includes(q)) searchMatches.push({ row: ri, col: ci }); 
    }));
    document.getElementById('searchCount').textContent = searchMatches.length ? `${searchMatches.length} ${t('found') || 'found'}` : t('noMatches') || 'No matches';
    searchMatches.forEach(m => document.querySelector(`td[data-row="${m.row}"][data-col="${m.col}"]`)?.classList.add('search-match'));
    if (searchMatches.length) navigateSearchResult(0, true);
}

function navigateSearchResult(dir, init = false) {
    if (!searchMatches.length) return;
    if (currentSearchIndex >= 0) { 
        const p = searchMatches[currentSearchIndex]; 
        document.querySelector(`td[data-row="${p.row}"][data-col="${p.col}"]`)?.classList.replace('search-current', 'search-match'); 
    }
    currentSearchIndex = init ? 0 : (currentSearchIndex + dir + searchMatches.length) % searchMatches.length;
    const c = searchMatches[currentSearchIndex], td = document.querySelector(`td[data-row="${c.row}"][data-col="${c.col}"]`);
    if (td) { td.classList.replace('search-match', 'search-current'); td.scrollIntoView({ behavior: 'smooth', block: 'center' }); }
    document.getElementById('searchCount').textContent = `${currentSearchIndex + 1} ${t('of') || 'of'} ${searchMatches.length}`;
}

// =====================================
// FIND & REPLACE DIALOG (Ctrl+H)
// =====================================
let replaceDialog = null;
let replaceMatches = [];
let currentReplaceIndex = -1;

function createReplaceDialog() {
    if (replaceDialog) return;
    replaceDialog = document.createElement('div');
    replaceDialog.className = 'replace-dialog hidden';
    replaceDialog.innerHTML = `
        <div class="replace-header">
            <span>${t('findAndReplace') || 'Find and Replace'}</span>
            <button id="replaceClose" class="replace-close-btn">×</button>
        </div>
        <div class="replace-content">
            <div class="replace-row">
                <label>${t('findLabel') || 'Find:'}</label>
                <input type="text" id="replaceFindInput" placeholder="${t('findPlaceholder') || 'Find...'}">
            </div>
            <div class="replace-row">
                <label>${t('replaceLabel') || 'Replace:'}</label>
                <input type="text" id="replaceWithInput" placeholder="${t('replacePlaceholder') || 'Replace with...'}">
            </div>
            <div class="replace-options">
                <label class="replace-option">
                    <input type="checkbox" id="replaceMatchCase">
                    <span>${t('matchCase') || 'Match case'}</span>
                </label>
                <label class="replace-option">
                    <input type="checkbox" id="replaceWholeCell">
                    <span>${t('matchWholeCell') || 'Match entire cell'}</span>
                </label>
            </div>
            <div class="replace-count" id="replaceCount"></div>
            <div class="replace-buttons">
                <button id="replaceFindPrev" title="Previous">◀ ${t('previous') || 'Prev'}</button>
                <button id="replaceFindNext" title="Next">${t('next') || 'Next'} ▶</button>
                <button id="replaceOne">${t('replace') || 'Replace'}</button>
                <button id="replaceAll">${t('replaceAll') || 'Replace All'}</button>
            </div>
        </div>
    `;
    document.body.appendChild(replaceDialog);
    
    document.getElementById('replaceFindInput').addEventListener('input', performReplaceSearch);
    document.getElementById('replaceFindInput').addEventListener('keydown', (e) => {
        if (e.key === 'Enter') {
            e.shiftKey ? navigateReplaceResult(-1) : navigateReplaceResult(1);
            e.preventDefault();
        } else if (e.key === 'Escape') closeReplace();
    });
    document.getElementById('replaceWithInput').addEventListener('keydown', (e) => {
        if (e.key === 'Enter') { e.preventDefault(); replaceOne(); } 
        else if (e.key === 'Escape') closeReplace();
    });
    document.getElementById('replaceMatchCase').addEventListener('change', performReplaceSearch);
    document.getElementById('replaceWholeCell').addEventListener('change', performReplaceSearch);
    document.getElementById('replaceFindPrev').addEventListener('click', () => navigateReplaceResult(-1));
    document.getElementById('replaceFindNext').addEventListener('click', () => navigateReplaceResult(1));
    document.getElementById('replaceOne').addEventListener('click', replaceOne);
    document.getElementById('replaceAll').addEventListener('click', replaceAllMatches);
    document.getElementById('replaceClose').addEventListener('click', closeReplace);
}

function openReplace() {
    closeSearch();
    createReplaceDialog();
    replaceDialog.classList.remove('hidden');
    document.getElementById('replaceFindInput').focus();
    document.getElementById('replaceFindInput').select();
}

function closeReplace() {
    if (replaceDialog) replaceDialog.classList.add('hidden');
    replaceMatches = [];
    currentReplaceIndex = -1;
    const countEl = document.getElementById('replaceCount');
    if (countEl) countEl.textContent = '';
    document.querySelectorAll('.search-match,.search-current').forEach(e => e.classList.remove('search-match', 'search-current'));
}

function performReplaceSearch() {
    const findText = document.getElementById('replaceFindInput').value;
    const matchCase = document.getElementById('replaceMatchCase').checked;
    const wholeCell = document.getElementById('replaceWholeCell').checked;
    
    replaceMatches = [];
    currentReplaceIndex = -1;
    document.querySelectorAll('.search-match,.search-current').forEach(e => e.classList.remove('search-match', 'search-current'));
    
    if (!findText) {
        document.getElementById('replaceCount').textContent = '';
        return;
    }
    
    const searchVal = matchCase ? findText : findText.toLowerCase();
    
    workbookData.forEach((r, ri) => r.forEach((c, ci) => {
        const cellVal = matchCase ? String(c) : String(c).toLowerCase();
        let isMatch = wholeCell ? cellVal === searchVal : cellVal.includes(searchVal);
        if (isMatch) replaceMatches.push({ row: ri, col: ci });
    }));
    
    document.getElementById('replaceCount').textContent = replaceMatches.length 
        ? `${replaceMatches.length} ${t('matchesFound') || 'match(es) found'}` 
        : t('noMatches') || 'No matches';
    
    replaceMatches.forEach(m => {
        document.querySelector(`td[data-row="${m.row}"][data-col="${m.col}"]`)?.classList.add('search-match');
    });
    
    if (replaceMatches.length) navigateReplaceResult(0, true);
}

function navigateReplaceResult(dir, init = false) {
    if (!replaceMatches.length) return;
    
    if (currentReplaceIndex >= 0 && currentReplaceIndex < replaceMatches.length) {
        const p = replaceMatches[currentReplaceIndex];
        document.querySelector(`td[data-row="${p.row}"][data-col="${p.col}"]`)?.classList.replace('search-current', 'search-match');
    }
    
    currentReplaceIndex = init ? 0 : (currentReplaceIndex + dir + replaceMatches.length) % replaceMatches.length;
    const c = replaceMatches[currentReplaceIndex];
    const td = document.querySelector(`td[data-row="${c.row}"][data-col="${c.col}"]`);
    
    if (td) {
        td.classList.replace('search-match', 'search-current');
        td.scrollIntoView({ behavior: 'smooth', block: 'center' });
    }
    
    document.getElementById('replaceCount').textContent = `${currentReplaceIndex + 1} ${t('of') || 'of'} ${replaceMatches.length}`;
}

function replaceOne() {
    if (!replaceMatches.length || currentReplaceIndex < 0) return;
    
    const findText = document.getElementById('replaceFindInput').value;
    const replaceText = document.getElementById('replaceWithInput').value;
    const matchCase = document.getElementById('replaceMatchCase').checked;
    const wholeCell = document.getElementById('replaceWholeCell').checked;
    
    if (!findText) return;
    
    saveState('replace');
    
    const match = replaceMatches[currentReplaceIndex];
    const cellVal = String(workbookData[match.row][match.col]);
    
    if (wholeCell) {
        workbookData[match.row][match.col] = replaceText;
    } else {
        const regex = new RegExp(escapeRegex(findText), matchCase ? 'g' : 'gi');
        workbookData[match.row][match.col] = cellVal.replace(regex, replaceText);
    }
    
    renderTable();
    performReplaceSearch();
    showToast(t('replaced') || 'Replaced 1 match');
}

function replaceAllMatches() {
    const findText = document.getElementById('replaceFindInput').value;
    const replaceText = document.getElementById('replaceWithInput').value;
    const matchCase = document.getElementById('replaceMatchCase').checked;
    const wholeCell = document.getElementById('replaceWholeCell').checked;
    
    if (!findText || !replaceMatches.length) return;
    
    saveState('replaceAll');
    
    let count = 0;
    const searchVal = matchCase ? findText : findText.toLowerCase();
    
    workbookData.forEach((r, ri) => r.forEach((c, ci) => {
        const cellVal = String(c);
        const compareVal = matchCase ? cellVal : cellVal.toLowerCase();
        
        if (wholeCell) {
            if (compareVal === searchVal) {
                workbookData[ri][ci] = replaceText;
                count++;
            }
        } else {
            if (compareVal.includes(searchVal)) {
                const regex = new RegExp(escapeRegex(findText), matchCase ? 'g' : 'gi');
                workbookData[ri][ci] = cellVal.replace(regex, replaceText);
                count++;
            }
        }
    }));
    
    renderTable();
    performReplaceSearch();
    showToast(t('replacedAll', { count }) || `Replaced ${count} match(es)`);
}

function escapeRegex(string) {
    return string.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}
// =====================================
// RESIZE COLUMNS
// =====================================
function startResize(e, colIndex, th) {
    e.preventDefault();
    const startX = e.clientX;
    const startWidth = th.offsetWidth;
    
    function onMouseMove(e) {
        const newWidth = Math.max(50, startWidth + e.clientX - startX);
        th.style.width = newWidth + 'px';
        columnWidths[colIndex] = newWidth + 'px';
    }
    
    function onMouseUp() {
        document.removeEventListener('mousemove', onMouseMove);
        document.removeEventListener('mouseup', onMouseUp);
    }
    
    document.addEventListener('mousemove', onMouseMove);
    document.addEventListener('mouseup', onMouseUp);
}

// =====================================
// BUTTON STATES
// =====================================
function updateButtonStates() {
    const hasData = workbookData.length > 0;
    const hasApiKey = apiKey !== '' && aiModeEnabled;
    addRowBtn.disabled = !hasData;
    addColBtn.disabled = !hasData;
    compareBtn.disabled = !history.length;
    downloadBtn.disabled = !hasData;
    aiPrompt.disabled = !hasData || !hasApiKey;
    aiButton.disabled = !hasData || !hasApiKey;
    undoBtn.disabled = !history.length;
    
    // Update chat input area appearance
    const chatInputArea = document.querySelector('.chat-input-area');
    if (chatInputArea) {
        if (!aiModeEnabled) {
            chatInputArea.classList.add('disabled');
        } else {
            chatInputArea.classList.remove('disabled');
        }
    }
}

// =====================================
// AI PROCESSING
// =====================================
stopBtn.addEventListener('click', () => {
    if (isProcessing) {
        shouldStopProcessing = true;
        stopBtn.disabled = true;
        stopBtn.querySelector('span').textContent = t('cancellingProcess');
    }
});

aiButton.addEventListener('click', async () => {
    const prompt = aiPrompt.value.trim();
    if (!prompt) { showStatus(t('enterCommand'), 'error'); return; }
    addChatMessage('user', prompt);
    aiPrompt.value = '';
    
    isProcessing = true;
    shouldStopProcessing = false;
    preProcessingState = { headers: [...headers], data: workbookData.map(r => [...r]), colors: {...cellColors}, modifiedCells: new Set(modifiedCells) };
    
    stopBtn.classList.remove('hidden');
    stopBtn.disabled = false;
    stopBtn.querySelector('span').textContent = t('stopText');
    aiButton.disabled = true;
    
    saveState('ai');
    
    try {
        // Use smart optimized processing
        await processOptimizedRequest(prompt);
    } catch (err) {
        removeProgressMessage();
        addChatMessage('system', t('apiError', { error: err.message }));
    } finally {
        isProcessing = false;
        shouldStopProcessing = false;
        preProcessingState = null;
        stopBtn.classList.add('hidden');
        aiButton.disabled = false;
        updateButtonStates();
    }
});

function userAskedToDeleteRows(prompt) {
    const deleteKeywords = ['delete', 'remove', 'eliminate', 'drop', 'erase', 'clear rows', 'empty rows', 'blank rows', 'eliminar', 'borrar', 'elimina', 'esborra'];
    const p = prompt.toLowerCase();
    return deleteKeywords.some(k => p.includes(k));
}

// =====================================
// SMART DATA OPTIMIZATION
// =====================================

function detectRelevantColumns(prompt) {
    const promptLower = prompt.toLowerCase();
    const relevantCols = new Set();
    
    // Check for column letter references (A, B, C, etc.)
    const colLetterMatch = promptLower.match(/\bcolumn\s*([a-z])\b|\bcol\s*([a-z])\b|\bcolumna\s*([a-z])\b/gi);
    if (colLetterMatch) {
        colLetterMatch.forEach(match => {
            const letter = match.match(/[a-z]$/i)[0].toUpperCase();
            const colIndex = letter.charCodeAt(0) - 65; // A=0, B=1, etc.
            if (colIndex >= 0 && colIndex < headers.length) {
                relevantCols.add(colIndex);
            }
        });
    }
    
    // Check for column name references
    headers.forEach((header, index) => {
        if (!header) return;
        const headerLower = header.toLowerCase();
        const headerWords = headerLower.split(/[\s_-]+/);
        
        // Check if header name or any word in header appears in prompt
        if (promptLower.includes(headerLower)) {
            relevantCols.add(index);
        } else {
            // Check individual words (for multi-word headers)
            headerWords.forEach(word => {
                if (word.length > 2 && promptLower.includes(word)) {
                    relevantCols.add(index);
                }
            });
        }
    });
    
    // Check for "all columns" or similar
    if (promptLower.includes('all column') || promptLower.includes('every column') || 
        promptLower.includes('todas las columna') || promptLower.includes('totes les column')) {
        return null; // null means all columns
    }
    
    return relevantCols.size > 0 ? Array.from(relevantCols) : null;
}

function detectRelevantRows(prompt) {
    const promptLower = prompt.toLowerCase();
    
    // Check for "all rows" or similar
    if (promptLower.includes('all row') || promptLower.includes('every row') || 
        promptLower.includes('todas las fila') || promptLower.includes('totes les file')) {
        return null; // null means all rows
    }
    
    // Check for row range (e.g., "rows 1-10", "rows 1 to 50")
    const rangeMatch = promptLower.match(/rows?\s*(\d+)\s*[-to]+\s*(\d+)/i);
    if (rangeMatch) {
        const start = parseInt(rangeMatch[1]) - 1; // Convert to 0-indexed
        const end = parseInt(rangeMatch[2]);
        return { type: 'range', start: Math.max(0, start), end: Math.min(end, workbookData.length) };
    }
    
    // Check for "first N rows" or "last N rows"
    const firstMatch = promptLower.match(/first\s*(\d+)\s*rows?|primer[ao]s?\s*(\d+)\s*fila/i);
    if (firstMatch) {
        const n = parseInt(firstMatch[1] || firstMatch[2]);
        return { type: 'range', start: 0, end: Math.min(n, workbookData.length) };
    }
    
    const lastMatch = promptLower.match(/last\s*(\d+)\s*rows?|[úu]ltim[ao]s?\s*(\d+)\s*fila/i);
    if (lastMatch) {
        const n = parseInt(lastMatch[1] || lastMatch[2]);
        return { type: 'range', start: Math.max(0, workbookData.length - n), end: workbookData.length };
    }
    
    // Check for conditional filters in prompt
    // e.g., "where price > 100", "rows where name contains 'John'"
    const conditionPatterns = [
        /where\s+(\w+)\s*(>|<|>=|<=|=|==|!=|contains|equals?|is)\s*["']?([^"'\s]+)["']?/i,
        /(?:filas?|rows?)\s+(?:donde|where|con|with)\s+(\w+)\s*(>|<|>=|<=|=|==|!=|contien[ea]|igual)\s*["']?([^"'\s]+)["']?/i
    ];
    
    for (const pattern of conditionPatterns) {
        const match = promptLower.match(pattern);
        if (match) {
            const columnName = match[1];
            const operator = match[2];
            const value = match[3];
            
            // Find column index
            const colIndex = headers.findIndex(h => 
                h && h.toLowerCase().includes(columnName.toLowerCase())
            );
            
            if (colIndex !== -1) {
                return { type: 'condition', column: colIndex, operator, value };
            }
        }
    }
    
    return null; // null means all rows
}

function filterRowsByCondition(condition) {
    if (!condition || condition.type === 'range') {
        return condition;
    }
    
    if (condition.type === 'condition') {
        const { column, operator, value } = condition;
        const matchingIndices = [];
        
        workbookData.forEach((row, index) => {
            const cellValue = row[column];
            const cellStr = String(cellValue).toLowerCase();
            const valueLower = value.toLowerCase();
            const cellNum = parseFloat(cellValue);
            const valueNum = parseFloat(value);
            
            let matches = false;
            switch (operator) {
                case '>':
                    matches = !isNaN(cellNum) && !isNaN(valueNum) && cellNum > valueNum;
                    break;
                case '<':
                    matches = !isNaN(cellNum) && !isNaN(valueNum) && cellNum < valueNum;
                    break;
                case '>=':
                    matches = !isNaN(cellNum) && !isNaN(valueNum) && cellNum >= valueNum;
                    break;
                case '<=':
                    matches = !isNaN(cellNum) && !isNaN(valueNum) && cellNum <= valueNum;
                    break;
                case '=':
                case '==':
                case 'equals':
                case 'equal':
                case 'is':
                case 'igual':
                    matches = cellStr === valueLower || cellNum === valueNum;
                    break;
                case '!=':
                    matches = cellStr !== valueLower && cellNum !== valueNum;
                    break;
                case 'contains':
                case 'contiene':
                case 'conté':
                    matches = cellStr.includes(valueLower);
                    break;
            }
            
            if (matches) matchingIndices.push(index);
        });
        
        return { type: 'indices', indices: matchingIndices };
    }
    
    return null;
}

function getOptimizedData(colIndices, rowFilter) {
    // Determine which rows to include
    let rowIndices = [];
    if (!rowFilter) {
        rowIndices = workbookData.map((_, i) => i);
    } else if (rowFilter.type === 'range') {
        for (let i = rowFilter.start; i < rowFilter.end; i++) {
            rowIndices.push(i);
        }
    } else if (rowFilter.type === 'indices') {
        rowIndices = rowFilter.indices;
    }
    
    // Determine which columns to include
    const colsToUse = colIndices || headers.map((_, i) => i);
    
    // Always include at least one identifier column (first column or first non-empty)
    if (colIndices && colIndices.length > 0 && !colIndices.includes(0)) {
        // Add first column as identifier if not already included
        colsToUse.unshift(0);
    }
    
    // Build optimized headers
    const optHeaders = colsToUse.map(ci => headers[ci]);
    
    // Build optimized data with row index tracking
    const optData = rowIndices.map(ri => {
        const row = colsToUse.map(ci => workbookData[ri][ci]);
        return row;
    });
    
    return {
        headers: optHeaders,
        data: optData,
        colMapping: colsToUse,
        rowMapping: rowIndices,
        isOptimized: colIndices !== null || rowFilter !== null
    };
}

function mergeOptimizedResults(originalData, aiResult, optimization) {
    if (!optimization.isOptimized) {
        // Not optimized, use standard merge
        return aiResult;
    }
    
    const { colMapping, rowMapping } = optimization;
    
    // Create a copy of original data
    let newHeaders = [...headers];
    let newData = originalData.map(r => [...r]);
    
    // Check if AI added new columns
    const aiHeadersLower = aiResult.headers.map(h => String(h).toLowerCase().trim());
    const newColsFromAI = [];
    
    aiResult.headers.forEach((h, ai) => {
        const hLower = String(h).toLowerCase().trim();
        const existingIndex = headers.findIndex(oh => String(oh).toLowerCase().trim() === hLower);
        if (existingIndex === -1 && !colMapping.some(ci => headers[ci]?.toLowerCase().trim() === hLower)) {
            newColsFromAI.push({ header: h, aiIndex: ai });
            newHeaders.push(h);
            newData.forEach(r => r.push(''));
        }
    });
    
    // Map AI results back to original positions
    aiResult.data.forEach((aiRow, aiRowIndex) => {
        const originalRowIndex = rowMapping[aiRowIndex];
        if (originalRowIndex === undefined || originalRowIndex >= newData.length) return;
        
        aiRow.forEach((value, aiColIndex) => {
            // Find the original column index
            const aiHeader = aiResult.headers[aiColIndex];
            const aiHeaderLower = String(aiHeader).toLowerCase().trim();
            
            // Check if it's a mapped column
            let originalColIndex = -1;
            if (aiColIndex < colMapping.length) {
                const mappedCol = colMapping[aiColIndex];
                if (headers[mappedCol]?.toLowerCase().trim() === aiHeaderLower) {
                    originalColIndex = mappedCol;
                }
            }
            
            // If not found in mapping, search in all headers
            if (originalColIndex === -1) {
                originalColIndex = newHeaders.findIndex(h => String(h).toLowerCase().trim() === aiHeaderLower);
            }
            
            if (originalColIndex !== -1 && value !== undefined) {
                newData[originalRowIndex][originalColIndex] = value;
            }
        });
    });
    
    return { headers: newHeaders, data: newData };
}

async function processOptimizedRequest(prompt) {
    addChatMessage('ai', t('processing'));
    
    // Detect relevant columns and rows
    const relevantCols = detectRelevantColumns(prompt);
    const rowFilter = filterRowsByCondition(detectRelevantRows(prompt));
    
    // Get optimized dataset
    const optimization = getOptimizedData(relevantCols, rowFilter);
    
    // Show optimization info if applicable
    if (optimization.isOptimized) {
        const colCount = optimization.headers.length;
        const rowCount = optimization.data.length;
        const totalCols = headers.length;
        const totalRows = workbookData.length;
        
        if (colCount < totalCols || rowCount < totalRows) {
            const saved = Math.round((1 - (colCount * rowCount) / (totalCols * totalRows)) * 100);
            if (saved > 10) {
                updateProgressMessage(`${t('processing')} (${t('optimized') || 'optimized'}: ${colCount}/${totalCols} cols, ${rowCount}/${totalRows} rows)`);
            }
        }
    }
    
    // Determine if we need batching
    const dataSize = optimization.data.length;
    
    if (dataSize > 100) {
        await processLargeDatasetOptimized(prompt, optimization);
    } else {
        await processSingleRequestOptimized(prompt, optimization);
    }
}

async function processSingleRequestOptimized(prompt, optimization) {
    const res = await callAI(optimization.headers, optimization.data, prompt, 1, 1);
    
    if (shouldStopProcessing) {
        if (preProcessingState) {
            headers = preProcessingState.headers;
            workbookData = preProcessingState.data;
            cellColors = preProcessingState.colors;
            modifiedCells = preProcessingState.modifiedCells;
        }
        removeProgressMessage();
        addChatMessage('system', t('processingCancelled'));
        renderTable();
        return;
    }
    
    removeProgressMessage();
    
    if (res.type === 'question') {
        addChatMessage('ai', res.answer);
        history.pop();
    } else if (res.type === 'modification' && res.headers && res.data) {
        const oh = [...headers], od = workbookData.map(r => [...r]);
        
        // Merge optimized results back
        let mergedResult;
        if (optimization.isOptimized) {
            mergedResult = mergeOptimizedResults(od, res, optimization);
        } else {
            mergedResult = mergeAIChangesIntoOriginal(od, res.data, oh, res.headers, prompt);
        }
        
        headers = mergedResult.headers || oh;
        workbookData = mergedResult.data || mergedResult;
        
        if (workbookData.length < od.length && !userAskedToDeleteRows(prompt)) {
            workbookData = od.map(r => [...r]);
            addChatMessage('system', '🛡️ Emergency protection activated.');
        }
        
        const { changes, modified } = detectChanges(oh, od, headers, workbookData);
        modifiedCells = modified;
        lastAIChange = { before: { headers: oh, data: od }, after: { headers, data: workbookData } };
        renderTable();
        
        if (mergedResult.message) addChatMessage('system', mergedResult.message);
        addChatMessage('ai', `✓ ${res.description || changes.join(', ') || 'Changes applied'}`);
    } else if (res.headers && res.data) {
        // Handle non-optimized full replacement
        if (!optimization.isOptimized) {
            headers = res.headers;
            workbookData = res.data;
        } else {
            const mergedResult = mergeOptimizedResults(workbookData, res, optimization);
            headers = mergedResult.headers;
            workbookData = mergedResult.data;
        }
        renderTable();
        addChatMessage('ai', t('changesApplied'));
    } else {
        throw new Error('Invalid response');
    }
}

async function processLargeDatasetOptimized(prompt, optimization) {
    const bs = calculateDynamicBatchSize();
    const tb = Math.ceil(optimization.data.length / bs);
    addChatMessage('system', t('largeDataset', { rows: optimization.data.length, batches: tb }));
    
    const oh = [...headers], od = workbookData.map(r => [...r]);
    const am = new Set();
    
    for (let i = 0; i < tb; i++) {
        if (shouldStopProcessing) {
            headers = preProcessingState.headers;
            workbookData = preProcessingState.data;
            cellColors = preProcessingState.colors;
            modifiedCells = preProcessingState.modifiedCells;
            removeProgressMessage();
            addChatMessage('system', t('processingCancelled'));
            renderTable();
            return;
        }
        
        const start = i * bs, end = Math.min(start + bs, optimization.data.length);
        const batchData = optimization.data.slice(start, end);
        const batchRowMapping = optimization.rowMapping.slice(start, end);
        
        updateProgressMessage(t('processingBatch', { current: i + 1, total: tb, start: start + 1, end }));
        
        try {
            const res = await callAI(optimization.headers, batchData, prompt, i + 1, tb);
            if (shouldStopProcessing) continue;
            
            if (res.type === 'modification' && res.data) {
                res.data.forEach((row, ri) => {
                    const originalRowIndex = batchRowMapping[ri];
                    if (originalRowIndex === undefined || originalRowIndex >= workbookData.length) return;
                    
                    row.forEach((val, ci) => {
                        // Map back to original column
                        const originalColIndex = optimization.colMapping[ci];
                        if (originalColIndex !== undefined && originalColIndex < headers.length) {
                            if (String(workbookData[originalRowIndex][originalColIndex]) !== String(val)) {
                                workbookData[originalRowIndex][originalColIndex] = val;
                                am.add(`${originalRowIndex}-${originalColIndex}`);
                            }
                        }
                    });
                });
            } else if (res.type === 'question' && i === 0) {
                removeProgressMessage();
                addChatMessage('ai', res.answer);
                history.pop();
                return;
            }
            
            if (i < tb - 1) await new Promise(r => setTimeout(r, 300));
        } catch (be) {
            addChatMessage('system', `⚠️ Batch ${i + 1} failed: ${be.message}`);
        }
    }
    
    removeProgressMessage();
    modifiedCells = am;
    lastAIChange = { before: { headers: oh, data: od }, after: { headers, data: workbookData } };
    renderTable();
    addChatMessage('ai', t('batchesComplete', { batches: tb, cells: am.size }));
}

function mergeAIChangesIntoOriginal(originalData, aiData, originalHeaders, aiHeaders, prompt) {
    let newHeaders = [...originalHeaders];
    let newData = originalData.map(r => [...r]);
    let message = null;
    
    const aiHeadersLower = aiHeaders.map(h => String(h).toLowerCase().trim());
    const originalHeadersLower = originalHeaders.map(h => String(h).toLowerCase().trim());
    
    aiHeaders.forEach((h, i) => {
        const hLower = String(h).toLowerCase().trim();
        if (!originalHeadersLower.includes(hLower)) { newHeaders.push(h); newData.forEach(r => r.push('')); }
    });
    
    const headerMapping = aiHeaders.map(h => {
        const hLower = String(h).toLowerCase().trim();
        let idx = originalHeadersLower.indexOf(hLower);
        if (idx === -1) idx = newHeaders.findIndex(nh => String(nh).toLowerCase().trim() === hLower);
        return idx;
    });
    
    const matchRows = (origRow, aiRows) => {
        for (let i = 0; i < aiRows.length; i++) {
            let matches = 0, comparisons = 0;
            for (let j = 0; j < Math.min(3, aiHeaders.length); j++) {
                const origIdx = headerMapping[j];
                if (origIdx !== -1 && origRow[origIdx] != null && aiRows[i][j] != null) {
                    comparisons++;
                    if (String(origRow[origIdx]).trim() === String(aiRows[i][j]).trim()) matches++;
                }
            }
            if (comparisons > 0 && matches / comparisons >= 0.5) return i;
        }
        return -1;
    };
    
    const aiDataCopy = aiData.map(r => [...r]);
    newData.forEach((row, ri) => {
        const aiRowIdx = matchRows(row, aiDataCopy);
        if (aiRowIdx !== -1) {
            aiHeaders.forEach((_, ai) => {
                const ti = headerMapping[ai];
                if (ti !== -1 && aiDataCopy[aiRowIdx][ai] !== undefined) row[ti] = aiDataCopy[aiRowIdx][ai];
            });
            aiDataCopy.splice(aiRowIdx, 1);
        }
    });
    
    if (aiDataCopy.length > 0 && (prompt.toLowerCase().includes('add') || prompt.toLowerCase().includes('new') || prompt.toLowerCase().includes('añad') || prompt.toLowerCase().includes('afeg'))) {
        aiDataCopy.forEach(ar => {
            const nr = new Array(newHeaders.length).fill('');
            aiHeaders.forEach((_, ai) => { const ti = headerMapping[ai]; if (ti !== -1) nr[ti] = ar[ai]; });
            newData.push(nr);
        });
        message = `Added ${aiDataCopy.length} new row(s)`;
    }
    
    return { headers: newHeaders, data: newData, message };
}

function detectChanges(oldHeaders, oldData, newHeaders, newData) {
    const changes = [], modified = new Set();
    if (newHeaders.length > oldHeaders.length) changes.push(t('columnsAdded', { count: newHeaders.length - oldHeaders.length }));
    if (newData.length > oldData.length) changes.push(t('rowsAdded', { count: newData.length - oldData.length }));
    if (newData.length < oldData.length) changes.push(t('rowsRemoved', { count: oldData.length - newData.length }));
    let mc = 0;
    for (let r = 0; r < Math.min(oldData.length, newData.length); r++) {
        for (let c = 0; c < Math.min(oldHeaders.length, newHeaders.length); c++) {
            if (String(oldData[r]?.[c] ?? '') !== String(newData[r]?.[c] ?? '')) { mc++; modified.add(`${r}-${c}`); }
        }
    }
    if (mc > 0) changes.push(t('cellsModified', { count: mc }));
    return { changes, modified };
}

async function processLargeDataset(prompt) {
    const bs = calculateDynamicBatchSize();
    const tb = Math.ceil(workbookData.length / bs);
    addChatMessage('system', t('largeDataset', { rows: workbookData.length, batches: tb }));
    
    const oh = [...headers], od = workbookData.map(r => [...r]);
    const am = new Set();
    
    for (let i = 0; i < tb; i++) {
        if (shouldStopProcessing) { headers = preProcessingState.headers; workbookData = preProcessingState.data; cellColors = preProcessingState.colors; modifiedCells = preProcessingState.modifiedCells; removeProgressMessage(); addChatMessage('system', t('processingCancelled')); renderTable(); return; }
        const start = i * bs, end = Math.min(start + bs, workbookData.length);
        const bd = workbookData.slice(start, end);
        updateProgressMessage(t('processingBatch', { current: i+1, total: tb, start: start+1, end }));
        
        try {
            const res = await callAI(headers, bd, prompt, i+1, tb);
            if (shouldStopProcessing) continue;
            if (res.type === 'modification' && res.data) {
                res.data.forEach((row, ri) => {
                    const ti = start + ri;
                    if (ti < workbookData.length) {
                        row.forEach((val, ci) => {
                            if (ci < headers.length && String(workbookData[ti][ci]) !== String(val)) { workbookData[ti][ci] = val; am.add(`${ti}-${ci}`); }
                        });
                    }
                });
            }
            else if (res.type === 'question' && i === 0) { removeProgressMessage(); addChatMessage('ai', res.answer); history.pop(); return; }
            if (i < tb - 1) await new Promise(r => setTimeout(r, 300));
        } catch (be) { addChatMessage('system', `⚠️ Batch ${i+1} failed: ${be.message}`); }
    }
    removeProgressMessage();
    modifiedCells = am; lastAIChange = { before: { headers: oh, data: od }, after: { headers, data: workbookData } };
    renderTable(); addChatMessage('ai', t('batchesComplete', { batches: tb, cells: am.size }));
}

async function processSingleRequest(prompt) {
    addChatMessage('ai', t('processing'));
    const res = await callAI(headers, workbookData, prompt, 1, 1);
    if (shouldStopProcessing) { if (preProcessingState) { headers = preProcessingState.headers; workbookData = preProcessingState.data; cellColors = preProcessingState.colors; modifiedCells = preProcessingState.modifiedCells; } removeProgressMessage(); addChatMessage('system', t('processingCancelled')); renderTable(); return; }
    removeProgressMessage();
    if (res.type === 'question') { addChatMessage('ai', res.answer); history.pop(); }
    else if (res.type === 'modification' && res.headers && res.data) {
        const oh = [...headers], od = workbookData.map(r => [...r]);
        const mr = mergeAIChangesIntoOriginal(od, res.data, oh, res.headers, prompt);
        headers = mr.headers; workbookData = mr.data;
        if (workbookData.length < od.length && !userAskedToDeleteRows(prompt)) { workbookData = od.map(r => [...r]); addChatMessage('system', '🛡️ Emergency protection activated.'); }
        const { changes, modified } = detectChanges(oh, od, headers, workbookData);
        modifiedCells = modified; lastAIChange = { before: { headers: oh, data: od }, after: { headers, data: workbookData } };
        renderTable();
        if (mr.message) addChatMessage('system', mr.message);
        addChatMessage('ai', `✓ ${res.description || changes.join(', ') || 'Changes applied'}`);
    } else if (res.headers && res.data) { headers = res.headers; workbookData = res.data; renderTable(); addChatMessage('ai', t('changesApplied')); }
    else throw new Error('Invalid response');
}

async function callAI(h, d, p, bn, tb) {
    const sp = `You are a helpful assistant for spreadsheet data. Determine if user is ASKING A QUESTION or REQUESTING A MODIFICATION.
For QUESTIONS: {"type": "question", "answer": "..."}
For MODIFICATIONS: {"type": "modification", "headers": [...], "data": [[...]], "description": "..."}
CRITICAL: Return ALL rows, NEVER omit unless explicitly asked to delete.${tb > 1 ? ` This is batch ${bn}/${tb}.` : ''} Respond with valid JSON only.`;
    const resp = await fetch('https://api.openai.com/v1/chat/completions', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', 'Authorization': `Bearer ${apiKey}` },
        body: JSON.stringify({ model: 'gpt-4o-mini', messages: [{ role: 'system', content: sp }, { role: 'user', content: `Headers: ${JSON.stringify(h)}\nData (${d.length} rows): ${JSON.stringify(d)}\n\nRequest: ${p}` }], temperature: 0.3, max_tokens: 8000 })
    });
    if (!resp.ok) { const e = await resp.json().catch(()=>({})); throw new Error(`API error: ${resp.status} - ${e.error?.message || 'Unknown'}`); }
    const r = await resp.json(), c = r.choices[0].message.content.trim(), m = c.match(/\{[\s\S]*\}/);
    return m ? JSON.parse(m[0]) : { type: 'question', answer: c };
}

function calculateDynamicBatchSize() {
    const cols = headers.length;
    let tc = 0, cc = 0;
    for (let i = 0; i < Math.min(50, workbookData.length); i++) for (let j = 0; j < cols; j++) { if (workbookData[i]?.[j] != null) { tc += String(workbookData[i][j]).length; cc++; } }
    const acs = cc ? tc / cc : 10, ars = acs * cols;
    let bs = Math.floor(12000 / (ars * 2));
    bs = Math.max(10, Math.min(200, bs));
    if (cols > 30) bs = Math.min(bs, 30); else if (cols > 20) bs = Math.min(bs, 50); else if (cols > 10) bs = Math.min(bs, 80);
    return bs;
}

function updateProgressMessage(txt) {
    const msgs = chatMessages.querySelectorAll('.message.ai'), last = msgs[msgs.length - 1];
    if (last && (last.textContent.includes('Processing') || last.textContent.includes('batch'))) last.textContent = txt;
    else addChatMessage('ai', txt);
}

function removeProgressMessage() {
    const msgs = chatMessages.querySelectorAll('.message.ai'), last = msgs[msgs.length - 1];
    if (last && (last.textContent === t('processing') || last.textContent.includes('Processing') || last.textContent.includes('batch'))) last.remove();
}
// =====================================
// SAVE DIALOG (Ctrl+S) - Native Windows Dialog
// =====================================
let hasUnsavedChanges = false;
let pendingExitAfterSave = false;

// Check if running in Electron (electronAPI exposed via preload)
const isElectron = typeof window.electronAPI !== 'undefined';

// Track changes
function markUnsavedChanges() {
    hasUnsavedChanges = true;
    if (isElectron) {
        window.electronAPI.setUnsavedChanges(true);
    }
}

function clearUnsavedChanges() {
    hasUnsavedChanges = false;
    if (isElectron) {
        window.electronAPI.setUnsavedChanges(false);
    }
}

// Listen for save-before-exit trigger from main process
if (isElectron) {
    window.electronAPI.onTriggerSaveBeforeExit(() => {
        pendingExitAfterSave = true;
        openSaveDialog();
    });
}

async function openSaveDialog() {
    if (!workbookData.length) {
        showToast(t('noDataToSave') || 'No data to save');
        return;
    }
    
    if (isElectron) {
        // Use native Windows save dialog via Electron
        const defaultName = originalFileName || 'spreadsheet';
        const currentFormat = exportFormat?.value || 'xlsx';
        
        try {
            const result = await window.electronAPI.showSaveDialog(defaultName, currentFormat);
            
            if (!result.canceled && result.filePath) {
                await performNativeSave(result.filePath);
            } else {
                // User cancelled
                pendingExitAfterSave = false;
            }
        } catch (error) {
            console.error('Save dialog error:', error);
            showToast(t('saveError') || 'Error saving file');
            pendingExitAfterSave = false;
        }
    } else {
        // Fallback for browser - use download
        performBrowserDownload();
    }
}

async function performNativeSave(filePath) {
    try {
        // Determine format from file extension
        const ext = filePath.split('.').pop().toLowerCase();
        const format = ext === 'csv' ? 'csv' : ext === 'xls' ? 'xls' : 'xlsx';
        
        // Create workbook
        const ws = XLSX.utils.aoa_to_sheet([headers, ...workbookData]);
        for (let c in cellColors) {
            if (ws[c]) ws[c].s = { fgColor: { rgb: cellColors[c].slice(1) } };
        }
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
        
        // Write to buffer
        const wbout = XLSX.write(wb, { 
            bookType: format, 
            type: 'array'
        });
        
        // Save via IPC
        const result = await window.electronAPI.writeFile(filePath, wbout);
        
        if (result.success) {
            clearUnsavedChanges();
            const fileName = filePath.split(/[\\/]/).pop();
            showToast(t('fileSaved') || 'File saved successfully!');
            addChatMessage('system', t('fileSavedAs', { filename: fileName }) || `File saved as ${fileName}`);
            
            // If we were saving before exit, now close the app
            if (pendingExitAfterSave) {
                pendingExitAfterSave = false;
                window.electronAPI.saveCompletedExit();
            }
        } else {
            throw new Error(result.error);
        }
    } catch (error) {
        console.error('Save error:', error);
        showToast(t('saveError') || 'Error saving file: ' + error.message);
        pendingExitAfterSave = false;
    }
}

function performBrowserDownload() {
    // Fallback for browser environment
    const fmt = exportFormat?.value || 'xlsx';
    const ws = XLSX.utils.aoa_to_sheet([headers, ...workbookData]);
    for (let c in cellColors) if (ws[c]) ws[c].s = { fgColor: { rgb: cellColors[c].slice(1) } };
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
    XLSX.writeFile(wb, `${originalFileName || 'spreadsheet'}.${fmt}`, fmt === 'csv' ? { bookType: 'csv' } : fmt === 'xls' ? { bookType: 'xls' } : {});
    clearUnsavedChanges();
    showToast(t('fileSaved') || 'File saved successfully!');
}

function closeSaveDialog() {
    // No longer needed with native dialog, but keep for compatibility
}

// Browser close/refresh warning (for when not in Electron)
window.addEventListener('beforeunload', (e) => {
    if (hasUnsavedChanges && workbookData.length > 0 && !isElectron) {
        e.preventDefault();
        e.returnValue = '';
        return '';
    }
});

// =====================================
// COLUMN SORTING
// =====================================
let sortState = { column: null, direction: null }; // null, 'asc', 'desc'

function sortByColumn(colIndex) {
    if (!workbookData.length) return;
    
    saveState('sort');
    markUnsavedChanges();
    
    // Determine sort direction
    if (sortState.column === colIndex) {
        if (sortState.direction === 'asc') {
            sortState.direction = 'desc';
        } else if (sortState.direction === 'desc') {
            sortState.direction = null; // Reset to original
            sortState.column = null;
            // Restore original order - we can't easily do this without storing original indices
            // So we just keep desc as the last state
            sortState.direction = 'asc';
        } else {
            sortState.direction = 'asc';
        }
    } else {
        sortState.column = colIndex;
        sortState.direction = 'asc';
    }
    
    workbookData.sort((a, b) => {
        let valA = a[colIndex];
        let valB = b[colIndex];
        
        // Handle empty values - put them at the end
        if (valA === '' || valA === null || valA === undefined) return 1;
        if (valB === '' || valB === null || valB === undefined) return -1;
        
        // Try numeric comparison first
        const numA = parseFloat(valA);
        const numB = parseFloat(valB);
        
        if (!isNaN(numA) && !isNaN(numB)) {
            return sortState.direction === 'asc' ? numA - numB : numB - numA;
        }
        
        // Try date comparison
        const dateA = new Date(valA);
        const dateB = new Date(valB);
        if (!isNaN(dateA.getTime()) && !isNaN(dateB.getTime())) {
            return sortState.direction === 'asc' ? dateA - dateB : dateB - dateA;
        }
        
        // String comparison (case-insensitive)
        const strA = String(valA).toLowerCase();
        const strB = String(valB).toLowerCase();
        const comparison = strA.localeCompare(strB);
        return sortState.direction === 'asc' ? comparison : -comparison;
    });
    
    renderTable();
    const directionText = sortState.direction === 'asc' ? '↑' : '↓';
    showToast(t('sortedBy', { column: headers[colIndex], direction: directionText }) || `Sorted by ${headers[colIndex]} ${directionText}`);
}

// =====================================
// HELP MODAL
// =====================================
const helpBtn = document.getElementById('helpBtn');
const helpModal = document.getElementById('helpModal');
const closeHelpModal = document.getElementById('closeHelpModal');

if (helpBtn) {
    helpBtn.addEventListener('click', () => {
        helpModal.classList.remove('hidden');
    });
}

if (closeHelpModal) {
    closeHelpModal.addEventListener('click', () => {
        helpModal.classList.add('hidden');
    });
}

if (helpModal) {
    helpModal.addEventListener('click', (e) => {
        if (e.target === helpModal) helpModal.classList.add('hidden');
    });
}

// =====================================
// UI BUTTON HANDLERS
// =====================================
undoBtn.addEventListener('click', () => { 
    if (!history.length) return; 
    const p = history.pop(); 
    headers = p.headers; 
    workbookData = p.data; 
    cellColors = p.colors; 
    modifiedCells = p.modifiedCells || new Set(); 
    renderTable(); 
    addChatMessage('system', t('undidChange')); 
    updateButtonStates(); 
});

addRowBtn.addEventListener('click', () => { 
    saveState('addRow'); 
    workbookData.push(new Array(headers.length).fill('')); 
    renderTable(); 
    addChatMessage('system', t('addedRow')); 
});

addColBtn.addEventListener('click', () => { 
    const n = prompt(t('enterColumnName')); 
    if (!n) return; 
    saveState('addColumn'); 
    headers.push(n); 
    workbookData.forEach(r => r.push('')); 
    renderTable(); 
    addChatMessage('system', t('addedColumn', { name: n })); 
});

// =====================================
// COMPARE MODAL
// =====================================
compareBtn.addEventListener('click', () => { 
    if (!history.length) { addChatMessage('system', t('noChangesToCompare')); return; } 
    const bs = history[history.length - 1]; 
    renderCompareTable('before', bs.headers, bs.data); 
    renderCompareTable('after', headers, workbookData, bs.headers, bs.data); 
    compareModal.classList.remove('hidden'); 
});

closeCompareModal.addEventListener('click', () => compareModal.classList.add('hidden'));
compareModal.addEventListener('click', (e) => { if (e.target === compareModal) compareModal.classList.add('hidden'); });

function renderCompareTable(pf, th, td, ch = null, cd = null) {
    const thead = document.getElementById(`${pf}TableHead`), tbody = document.getElementById(`${pf}TableBody`);
    thead.innerHTML = ''; tbody.innerHTML = '';
    const hr = document.createElement('tr'), cc = document.createElement('th'); cc.className = 'row-header'; cc.textContent = '#'; hr.appendChild(cc);
    th.forEach((h, i) => { const t = document.createElement('th'); t.textContent = h || `Col ${i+1}`; if (ch && ch[i] !== h) t.classList.add('changed'); hr.appendChild(t); });
    thead.appendChild(hr);
    td.forEach((r, ri) => {
        const tr = document.createElement('tr'), rh = document.createElement('td'); rh.className = 'row-header'; rh.textContent = ri + 1; tr.appendChild(rh);
        th.forEach((_, ci) => { const t = document.createElement('td'), v = r[ci] ?? ''; t.textContent = v; if (cd && String(v) !== String(cd[ri]?.[ci] ?? '')) t.classList.add('changed'); tr.appendChild(t); });
        tbody.appendChild(tr);
    });
}

// =====================================
// DOWNLOAD
// =====================================
downloadBtn.addEventListener('click', () => {
    const fmt = exportFormat?.value || 'xlsx';
    const ws = XLSX.utils.aoa_to_sheet([headers, ...workbookData]);
    for (let c in cellColors) if (ws[c]) ws[c].s = { fgColor: { rgb: cellColors[c].slice(1) } };
    const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
    XLSX.writeFile(wb, `${originalFileName || 'spreadsheet'}_modified.${fmt}`, fmt === 'csv' ? { bookType: 'csv' } : fmt === 'xls' ? { bookType: 'xls' } : {});
    addChatMessage('system', t('downloaded'));
});

// =====================================
// AI PROMPT ENTER KEY
// =====================================
aiPrompt.addEventListener('keydown', (e) => { if (e.key === 'Enter' && !e.shiftKey) { e.preventDefault(); aiButton.click(); } });

// =====================================
// INITIALIZE
// =====================================
updateButtonStates();
