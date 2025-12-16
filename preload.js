const { contextBridge, ipcRenderer } = require('electron');

// Expose protected methods to the renderer process
contextBridge.exposeInMainWorld('electronAPI', {
    // Save dialog
    showSaveDialog: (defaultName, format) => ipcRenderer.invoke('show-save-dialog', defaultName, format),
    
    // Write file
    writeFile: (filePath, data) => ipcRenderer.invoke('write-file', filePath, data),
    
    // Set unsaved changes flag
    setUnsavedChanges: (value) => ipcRenderer.send('set-unsaved-changes', value),
    
    // Notify save completed and exit
    saveCompletedExit: () => ipcRenderer.send('save-completed-exit'),
    
    // Listen for trigger-save-before-exit
    onTriggerSaveBeforeExit: (callback) => {
        ipcRenderer.on('trigger-save-before-exit', callback);
    },
    
    // Remove listener
    removeAllListeners: (channel) => {
        ipcRenderer.removeAllListeners(channel);
    }
});
