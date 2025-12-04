const translations = {
    en: {
        mainTitle: "🤖 Albert's AI Excel Editor",
        subtitle: "Intelligent spreadsheet editing powered by AI",
        welcomeTitle: "Welcome! Let's Get Started",
        welcomeText: "First, load a text file containing your OpenAI API key to unlock AI-powered features.",
        securityNote: "Your API key is stored securely in your browser and never leaves your device.",
        loadFileLabel: "Load File (.xlsx, .xls, .xlsm, .csv, .tsv, .json, .xml, .ods)",
        addRowText: "➕ Add Row",
        addColText: "➕ Add Column",
        compareText: "🔄 Compare Changes",
        downloadText: "⬇️ Download",
        chatHeaderText: "💬 AI Assistant",
        sendText: "Send ✨",
        undoText: "↶ Undo",
        highlightLabel: "Highlight AI changes",
        promptPlaceholder: "Ask questions about your data or request modifications...\ne.g., 'What's the total of column B?' or 'Add 10% to all prices'",
        apiKeyLoaded: "✓ API Key loaded successfully!",
        fileLoaded: "✓ Loaded {rows} rows and {cols} columns from \"{filename}\"",
        errorEmpty: "Error: File is empty",
        errorLoading: "Error loading file: {error}",
        enterCommand: "Please enter a message",
        processing: "🤔 Processing your request...",
        changesApplied: "✓ Changes applied successfully!",
        apiError: "✗ Error: {error}",
        undidChange: "↶ Undid last change",
        addedRow: "✓ Added new row",
        addedColumn: "✓ Added column \"{name}\"",
        downloaded: "✓ Downloaded successfully!",
        enterColumnName: "Enter column name:",
        previouslyUsed: "Previously used: {path}. Please load it again.",
        attemptingLoad: "Attempting to load saved API key...",
        compareModalTitle: "Compare Changes",
        beforeLabel: "Before",
        afterLabel: "After (Current)",
        noChangesToCompare: "No previous version to compare. Make some AI changes first!",
        selectAll: "All",
        clearAll: "Clear",
        changesDescription: "I made the following changes:\n{changes}",
        noChangesNeeded: "No changes were needed based on your request.",
        questionResponse: "Based on your data: {response}",
        cellsModified: "{count} cell(s) were modified",
        rowsAdded: "{count} row(s) added",
        rowsRemoved: "{count} row(s) removed",
        columnsAdded: "{count} column(s) added",
        columnsRemoved: "{count} column(s) removed",
        // New translations for enhanced features
        aiHistoryText: "📜 AI History",
        aiHistoryModalTitle: "📜 AI Change History",
        aiHistoryDescription: "Manage individual AI changes. You can revert specific changes while keeping others.",
        noAIHistory: "No AI changes recorded yet. Use the AI assistant to modify your data.",
        revertChange: "↩️ Revert",
        reapplyChange: "↪️ Reapply",
        viewDiff: "👁️ View Diff",
        revertedAIChange: "↩️ Reverted: {description}",
        reappliedAIChange: "↪️ Reapplied: {description}",
        shortcutsHint: "Ctrl+C/X/V: Copy/Cut/Paste | Ctrl+D: Fill Down | Ctrl+R: Fill Right | Del: Clear | Drag corner to fill",
        copiedCells: "📋 Copied {count} cell(s)",
        pastedCells: "📋 Pasted from clipboard",
        deletedCells: "🗑️ Cleared {count} cell(s)",
        fillApplied: "✓ Fill applied",
        fillDownApplied: "✓ Fill down applied",
        fillRightApplied: "✓ Fill right applied",
        exportFormat: "Export format",
        exportXLSX: "Excel (.xlsx)",
        exportXLS: "Excel 97-2003 (.xls)",
        exportCSV: "CSV (.csv)",
        exportTSV: "TSV (.tsv)",
        exportJSON: "JSON (.json)",
        exportXML: "XML (.xml)",
        exportHTML: "HTML (.html)",
        exportODS: "OpenDocument (.ods)"
    },
    es: {
        mainTitle: "🤖 Editor Excel IA de Albert",
        subtitle: "Edición inteligente de hojas de cálculo con IA",
        welcomeTitle: "¡Bienvenido! Empecemos",
        welcomeText: "Primero, carga un archivo de texto que contenga tu clave API de OpenAI para desbloquear las funciones de IA.",
        securityNote: "Tu clave API se almacena de forma segura en tu navegador y nunca sale de tu dispositivo.",
        loadFileLabel: "Cargar archivo (.xlsx, .xls, .xlsm, .csv, .tsv, .json, .xml, .ods)",
        addRowText: "➕ Añadir fila",
        addColText: "➕ Añadir columna",
        compareText: "🔄 Comparar Cambios",
        downloadText: "⬇️ Descargar",
        chatHeaderText: "💬 Asistente IA",
        sendText: "Enviar ✨",
        undoText: "↶ Deshacer",
        highlightLabel: "Resaltar cambios de IA",
        promptPlaceholder: "Haz preguntas sobre tus datos o solicita modificaciones...\nej., '¿Cuál es el total de la columna B?' o 'Añade un 10% a todos los precios'",
        apiKeyLoaded: "✓ ¡Clave API cargada con éxito!",
        fileLoaded: "✓ Cargadas {rows} filas y {cols} columnas de \"{filename}\"",
        errorEmpty: "Error: El archivo está vacío",
        errorLoading: "Error al cargar el archivo: {error}",
        enterCommand: "Por favor, introduce un mensaje",
        processing: "🤔 Procesando tu solicitud...",
        changesApplied: "✓ ¡Cambios aplicados con éxito!",
        apiError: "✗ Error: {error}",
        undidChange: "↶ Cambio deshecho",
        addedRow: "✓ Nueva fila añadida",
        addedColumn: "✓ Columna \"{name}\" añadida",
        downloaded: "✓ ¡Descargado con éxito!",
        enterColumnName: "Introduce el nombre de la columna:",
        previouslyUsed: "Usado anteriormente: {path}. Por favor, cárgalo de nuevo.",
        attemptingLoad: "Intentando cargar la clave API guardada...",
        compareModalTitle: "Comparar Cambios",
        beforeLabel: "Antes",
        afterLabel: "Después (Actual)",
        noChangesToCompare: "No hay versión anterior para comparar. ¡Haz algunos cambios con IA primero!",
        selectAll: "Todos",
        clearAll: "Limpiar",
        changesDescription: "Realicé los siguientes cambios:\n{changes}",
        noChangesNeeded: "No se necesitaron cambios según tu solicitud.",
        questionResponse: "Según tus datos: {response}",
        cellsModified: "{count} celda(s) fueron modificadas",
        rowsAdded: "{count} fila(s) añadida(s)",
        rowsRemoved: "{count} fila(s) eliminada(s)",
        columnsAdded: "{count} columna(s) añadida(s)",
        columnsRemoved: "{count} columna(s) eliminada(s)",
        // New translations
        aiHistoryText: "📜 Historial IA",
        aiHistoryModalTitle: "📜 Historial de Cambios IA",
        aiHistoryDescription: "Gestiona cambios individuales de IA. Puedes revertir cambios específicos manteniendo los demás.",
        noAIHistory: "No hay cambios de IA registrados. Usa el asistente IA para modificar tus datos.",
        revertChange: "↩️ Revertir",
        reapplyChange: "↪️ Reaplicar",
        viewDiff: "👁️ Ver Diferencias",
        revertedAIChange: "↩️ Revertido: {description}",
        reappliedAIChange: "↪️ Reaplicado: {description}",
        shortcutsHint: "Ctrl+C/X/V: Copiar/Cortar/Pegar | Ctrl+D: Rellenar Abajo | Ctrl+R: Rellenar Derecha | Supr: Borrar | Arrastra esquina para rellenar",
        copiedCells: "📋 {count} celda(s) copiada(s)",
        pastedCells: "📋 Pegado desde portapapeles",
        deletedCells: "🗑️ {count} celda(s) borrada(s)",
        fillApplied: "✓ Relleno aplicado",
        fillDownApplied: "✓ Relleno hacia abajo aplicado",
        fillRightApplied: "✓ Relleno hacia derecha aplicado",
        exportFormat: "Formato de exportación",
        exportXLSX: "Excel (.xlsx)",
        exportXLS: "Excel 97-2003 (.xls)",
        exportCSV: "CSV (.csv)",
        exportTSV: "TSV (.tsv)",
        exportJSON: "JSON (.json)",
        exportXML: "XML (.xml)",
        exportHTML: "HTML (.html)",
        exportODS: "OpenDocument (.ods)"
    },
    ca: {
        mainTitle: "🤖 Editor Excel IA d'Albert",
        subtitle: "Edició intel·ligent de fulls de càlcul amb IA",
        welcomeTitle: "Benvingut! Comencem",
        welcomeText: "Primer, carrega un arxiu de text que contingui la teva clau API d'OpenAI per desbloquejar les funcions d'IA.",
        securityNote: "La teva clau API s'emmagatzema de forma segura al teu navegador i mai surt del teu dispositiu.",
        loadFileLabel: "Carregar arxiu (.xlsx, .xls, .xlsm, .csv, .tsv, .json, .xml, .ods)",
        addRowText: "➕ Afegir fila",
        addColText: "➕ Afegir columna",
        compareText: "🔄 Comparar Canvis",
        downloadText: "⬇️ Descarregar",
        chatHeaderText: "💬 Assistent IA",
        sendText: "Enviar ✨",
        undoText: "↶ Desfer",
        highlightLabel: "Ressaltar canvis d'IA",
        promptPlaceholder: "Fes preguntes sobre les teves dades o sol·licita modificacions...\nex., 'Quin és el total de la columna B?' o 'Afegeix un 10% a tots els preus'",
        apiKeyLoaded: "✓ Clau API carregada amb èxit!",
        fileLoaded: "✓ Carregades {rows} files i {cols} columnes de \"{filename}\"",
        errorEmpty: "Error: L'arxiu està buit",
        errorLoading: "Error en carregar l'arxiu: {error}",
        enterCommand: "Si us plau, introdueix un missatge",
        processing: "🤔 Processant la teva sol·licitud...",
        changesApplied: "✓ Canvis aplicats amb èxit!",
        apiError: "✗ Error: {error}",
        undidChange: "↶ Canvi desfet",
        addedRow: "✓ Nova fila afegida",
        addedColumn: "✓ Columna \"{name}\" afegida",
        downloaded: "✓ Descarregat amb èxit!",
        enterColumnName: "Introdueix el nom de la columna:",
        previouslyUsed: "Usat anteriorment: {path}. Si us plau, carrega'l de nou.",
        attemptingLoad: "Intentant carregar la clau API desada...",
        compareModalTitle: "Comparar Canvis",
        beforeLabel: "Abans",
        afterLabel: "Després (Actual)",
        noChangesToCompare: "No hi ha versió anterior per comparar. Fes alguns canvis amb IA primer!",
        selectAll: "Tots",
        clearAll: "Netejar",
        changesDescription: "He fet els següents canvis:\n{changes}",
        noChangesNeeded: "No calen canvis segons la teva sol·licitud.",
        questionResponse: "Segons les teves dades: {response}",
        cellsModified: "{count} cel·la(es) modificada(es)",
        rowsAdded: "{count} fila(es) afegida(es)",
        rowsRemoved: "{count} fila(es) eliminada(es)",
        columnsAdded: "{count} columna(es) afegida(es)",
        columnsRemoved: "{count} columna(es) eliminada(es)",
        // New translations
        aiHistoryText: "📜 Historial IA",
        aiHistoryModalTitle: "📜 Historial de Canvis IA",
        aiHistoryDescription: "Gestiona canvis individuals d'IA. Pots revertir canvis específics mantenint els altres.",
        noAIHistory: "No hi ha canvis d'IA registrats. Usa l'assistent IA per modificar les teves dades.",
        revertChange: "↩️ Revertir",
        reapplyChange: "↪️ Reaplicar",
        viewDiff: "👁️ Veure Diferències",
        revertedAIChange: "↩️ Revertit: {description}",
        reappliedAIChange: "↪️ Reaplicat: {description}",
        shortcutsHint: "Ctrl+C/X/V: Copiar/Tallar/Enganxar | Ctrl+D: Omplir Avall | Ctrl+R: Omplir Dreta | Supr: Esborrar | Arrossega cantonada per omplir",
        copiedCells: "📋 {count} cel·la(es) copiada(es)",
        pastedCells: "📋 Enganxat des del porta-retalls",
        deletedCells: "🗑️ {count} cel·la(es) esborrada(es)",
        fillApplied: "✓ Emplenat aplicat",
        fillDownApplied: "✓ Emplenat cap avall aplicat",
        fillRightApplied: "✓ Emplenat cap a la dreta aplicat",
        exportFormat: "Format d'exportació",
        exportXLSX: "Excel (.xlsx)",
        exportXLS: "Excel 97-2003 (.xls)",
        exportCSV: "CSV (.csv)",
        exportTSV: "TSV (.tsv)",
        exportJSON: "JSON (.json)",
        exportXML: "XML (.xml)",
        exportHTML: "HTML (.html)",
        exportODS: "OpenDocument (.ods)"
    }
};

let currentLang = 'en';

function t(key, replacements = {}) {
    let text = translations[currentLang][key] || translations.en[key] || key;
    for (let [placeholder, value] of Object.entries(replacements)) {
        text = text.replace(`{${placeholder}}`, value);
    }
    return text;
}

function updateLanguage(lang) {
    currentLang = lang;
    localStorage.setItem('preferredLanguage', lang);
    
    // Update all text elements
    const elements = {
        'mainTitle': 'mainTitle',
        'subtitle': 'subtitle',
        'welcomeTitle': 'welcomeTitle',
        'welcomeText': 'welcomeText',
        'securityNote': 'securityNote',
        'loadFileLabel': 'loadFileLabel',
        'addRowText': 'addRowText',
        'addColText': 'addColText',
        'compareText': 'compareText',
        'downloadText': 'downloadText',
        'chatHeaderText': 'chatHeaderText',
        'sendText': 'sendText',
        'undoText': 'undoText',
        'highlightLabel': 'highlightLabel',
        'compareModalTitle': 'compareModalTitle',
        'beforeLabel': 'beforeLabel',
        'afterLabel': 'afterLabel',
        'aiHistoryText': 'aiHistoryText',
        'aiHistoryModalTitle': 'aiHistoryModalTitle',
        'aiHistoryDescription': 'aiHistoryDescription',
        'shortcutsHint': 'shortcutsHint'
    };

    for (const [id, key] of Object.entries(elements)) {
        const element = document.getElementById(id);
        if (element) {
            element.textContent = t(key);
        }
    }

    // Update placeholder
    const aiPrompt = document.getElementById('aiPrompt');
    if (aiPrompt) {
        aiPrompt.placeholder = t('promptPlaceholder');
    }
    
    // Update active button
    document.querySelectorAll('.lang-btn').forEach(btn => {
        btn.classList.toggle('active', btn.dataset.lang === lang);
    });
}

// Initialize language on load
window.addEventListener('DOMContentLoaded', () => {
    const savedLang = localStorage.getItem('preferredLanguage') || 'en';
    updateLanguage(savedLang);
    
    // Add click handlers to language buttons
    document.querySelectorAll('.lang-btn').forEach(btn => {
        btn.addEventListener('click', () => {
            updateLanguage(btn.dataset.lang);
        });
    });
});
