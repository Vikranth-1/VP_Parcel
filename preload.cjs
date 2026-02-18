const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('electronAPI', {
    savePDF: (buffer) => ipcRenderer.invoke('save-pdf', buffer),
});
