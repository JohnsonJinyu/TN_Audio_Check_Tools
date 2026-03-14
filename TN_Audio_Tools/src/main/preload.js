const { contextBridge, ipcRenderer } = require('electron');

function subscribeToChannel(channel, listener) {
  if (typeof listener !== 'function') {
    return () => {};
  }

  const wrappedListener = (_, payload) => listener(payload);
  ipcRenderer.on(channel, wrappedListener);

  return () => {
    ipcRenderer.removeListener(channel, wrappedListener);
  };
}

contextBridge.exposeInMainWorld('electron', {
  ipcRenderer: {
    invoke: (channel, ...args) => ipcRenderer.invoke(channel, ...args),
    on: (channel, func) => ipcRenderer.on(channel, (event, ...args) => func(...args)),
    send: (channel, ...args) => ipcRenderer.send(channel, ...args)
  },
  reportChecker: {
    processReports: (payload) => ipcRenderer.invoke('report-checker:process-reports', payload),
    onProgress: (listener) => subscribeToChannel('report-checker:progress', listener),
    showOutputInFolder: (filePath) => ipcRenderer.invoke('report-checker:show-output-in-folder', filePath),
    exportRules: (rulePath) => ipcRenderer.invoke('report-checker:export-rules', rulePath)
  }
});
