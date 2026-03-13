const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('electron', {
  ipcRenderer: {
    invoke: (channel, ...args) => ipcRenderer.invoke(channel, ...args),
    on: (channel, func) => ipcRenderer.on(channel, (event, ...args) => func(...args)),
    send: (channel, ...args) => ipcRenderer.send(channel, ...args)
  },
  reportChecker: {
    processReports: (payload) => ipcRenderer.invoke('report-checker:process-reports', payload),
    showOutputInFolder: (filePath) => ipcRenderer.invoke('report-checker:show-output-in-folder', filePath),
    exportRules: (rulePath) => ipcRenderer.invoke('report-checker:export-rules', rulePath)
  }
});
