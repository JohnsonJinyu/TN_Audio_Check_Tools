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
  appInfo: {
    getVersion: () => ipcRenderer.invoke('get-version')
  },
  settings: {
    get: () => ipcRenderer.invoke('app-settings:get'),
    getDefaults: () => ipcRenderer.invoke('app-settings:defaults'),
    save: (payload) => ipcRenderer.invoke('app-settings:save', payload),
    reset: () => ipcRenderer.invoke('app-settings:reset'),
    chooseOutputDirectory: () => ipcRenderer.invoke('app-settings:choose-output-directory'),
    clearCache: () => ipcRenderer.invoke('app-settings:clear-cache'),
    onChanged: (listener) => subscribeToChannel('app-settings:changed', listener)
  },
  updates: {
    getState: () => ipcRenderer.invoke('app-update:get-state'),
    checkForUpdates: () => ipcRenderer.invoke('app-update:check-for-updates'),
    downloadUpdate: () => ipcRenderer.invoke('app-update:download-update'),
    openExternalDownload: (payload) => ipcRenderer.invoke('app-update:open-external-download', payload),
    quitAndInstall: () => ipcRenderer.invoke('app-update:quit-and-install'),
    onStateChanged: (listener) => subscribeToChannel('app-update:state-changed', listener)
  },
  testDataCollection: {
    processReports: (payload) => ipcRenderer.invoke('report-checker:process-reports', payload),
    onProgress: (listener) => subscribeToChannel('report-checker:progress', listener),
    showOutputInFolder: (filePath) => ipcRenderer.invoke('report-checker:show-output-in-folder', filePath),
    getChecklistReportOptions: (checklistPath) => ipcRenderer.invoke('report-checker:get-checklist-report-options', checklistPath),
    inspectReportContext: (payload) => ipcRenderer.invoke('report-checker:inspect-report-context', payload),
    exportRules: (rulePath) => ipcRenderer.invoke('report-checker:export-rules', rulePath)
  },
  reportReview: {
    reviewWordReport: (payload) => ipcRenderer.invoke('report-review:review-word-report', payload),
    uploadWordReport: (payload) => ipcRenderer.invoke('report-review:upload-word-report', payload)
  }
});
