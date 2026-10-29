const { contextBridge, ipcRenderer, webUtils } = require('electron')

contextBridge.exposeInMainWorld('jarvisdoc', {
  onMenu: (fn) => {
    ipcRenderer.on('menu', (_e, payload) => {
      if (typeof payload === 'string') fn({ action: payload })
      else fn(payload || {})
    })
  },
  onLoad: (fn) => {
    ipcRenderer.on('doc:load', (_e, payload) => fn(payload))
  },
  onSaved: (fn) => {
    ipcRenderer.on('doc:saved', (_e, payload) => fn(payload))
  },
  onToast: (fn) => {
    ipcRenderer.on('doc:toast', (_e, message) => fn(message))
  },
  ready: () => ipcRenderer.invoke('doc:ready'),
  changed: (dirty) => ipcRenderer.send('doc:changed', Boolean(dirty)),
  new: () => ipcRenderer.invoke('doc:new'),
  open: () => ipcRenderer.invoke('doc:open'),
  save: (html) => ipcRenderer.invoke('doc:save', html),
  saveAs: (html, forcedExt) => ipcRenderer.invoke('doc:saveAs', { html, forcedExt }),
  insertImage: () => ipcRenderer.invoke('doc:insertImage'),
  openPath: (filePath) => ipcRenderer.invoke('doc:openPath', filePath),
  mdReport: (html) => ipcRenderer.invoke('md:report', html),
  platform: process.platform,
  pathForFile: (file) => {
    try {
      return webUtils.getPathForFile(file) || ''
    } catch {
      return ''
    }
  },
  onVersionsChanged: (fn) => {
    ipcRenderer.on('versions:changed', () => fn())
  },
  versions: {
    list: () => ipcRenderer.invoke('versions:list'),
    preview: (id) => ipcRenderer.invoke('versions:preview', id),
    restore: (id, html) => ipcRenderer.invoke('versions:restore', { id, html }),
    remove: (id) => ipcRenderer.invoke('versions:delete', id),
    save: (label, html) => ipcRenderer.invoke('versions:save', { label, html }),
    exportTo: (id) => ipcRenderer.invoke('versions:export', id)
  }
})
