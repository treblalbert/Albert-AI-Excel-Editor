const { app, BrowserWindow, ipcMain, dialog } = require('electron');
const path = require('path');
const fs = require('fs');

let mainWindow = null;
let hasUnsavedChanges = false;

function createWindow() {
    mainWindow = new BrowserWindow({
        width: 1400,
        height: 900,
        minWidth: 800,
        minHeight: 600,
        icon: path.join(__dirname, 'build', 'icon.ico'),
        webPreferences: {
            nodeIntegration: false,
            contextIsolation: true,
            preload: path.join(__dirname, 'preload.js')
        },
        title: "Albert's AI Excel Editor"
    });

    mainWindow.loadFile('src/index.html');
    mainWindow.setMenuBarVisibility(false);
    
    // Handle close event - show confirmation if unsaved changes
    mainWindow.on('close', async (e) => {
        if (hasUnsavedChanges) {
            e.preventDefault();
            
            const result = await dialog.showMessageBox(mainWindow, {
                type: 'question',
                buttons: ['Save', 'Don\'t Save', 'Cancel'],
                defaultId: 0,
                cancelId: 2,
                title: 'Unsaved Changes',
                message: 'Do you want to save your changes before exiting?',
                detail: 'Your changes will be lost if you don\'t save them.'
            });
            
            if (result.response === 0) {
                // Save - trigger save dialog
                mainWindow.webContents.send('trigger-save-before-exit');
            } else if (result.response === 1) {
                // Don't Save - close without saving
                hasUnsavedChanges = false;
                mainWindow.close();
            }
            // Cancel (response === 2) - do nothing, stay in app
        }
    });
}

// Handle save dialog request from renderer
ipcMain.handle('show-save-dialog', async (event, defaultName, format) => {
    const filters = [
        { name: 'Excel Workbook', extensions: ['xlsx'] },
        { name: 'Excel 97-2003 Workbook', extensions: ['xls'] },
        { name: 'CSV (Comma delimited)', extensions: ['csv'] },
        { name: 'All Files', extensions: ['*'] }
    ];
    
    const result = await dialog.showSaveDialog(mainWindow, {
        title: 'Save File',
        defaultPath: defaultName || 'spreadsheet',
        filters: filters
    });
    
    return result;
});

// Handle file write request from renderer
ipcMain.handle('write-file', async (event, filePath, data) => {
    try {
        fs.writeFileSync(filePath, Buffer.from(data));
        return { success: true };
    } catch (error) {
        return { success: false, error: error.message };
    }
});

// Handle unsaved changes flag from renderer
ipcMain.on('set-unsaved-changes', (event, value) => {
    hasUnsavedChanges = value;
});

// Handle save completed - can now close
ipcMain.on('save-completed-exit', () => {
    hasUnsavedChanges = false;
    mainWindow.close();
});

app.whenReady().then(createWindow);

app.on('window-all-closed', () => {
    if (process.platform !== 'darwin') {
        app.quit();
    }
});

app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) {
        createWindow();
    }
});
