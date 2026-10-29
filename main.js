const { app, BrowserWindow, Menu, MenuItem, dialog, ipcMain, nativeTheme, shell } = require('electron')
const { access, mkdtemp, readFile, rm, writeFile } = require('fs/promises')
const fs = require('fs')
const { constants } = require('fs')
const os = require('os')
const path = require('path')
const { fileToHtml, htmlToFileBuffer, toMarkdown, formatOf, ensureExt, isMarkdown, isDocx, DEFAULT_PAGE } = require('./doc-io')
const { markdownReport } = require('./md-fidelity')
const { convertToPdf } = require('./export-pdf')
const versions = require('./versions')

const AUTO_VERSION_MS = 5 * 60 * 1000

/** @type {BrowserWindow | null} */
const IS_MAC = process.platform === 'darwin'
const IS_WIN = process.platform === 'win32'
const IS_LINUX = process.platform === 'linux'

let mainWindow = null
/** @type {string | null} */
let currentPath = null
/** @type {Buffer | null} */
let sourceDocx = null
let currentPage = { ...DEFAULT_PAGE }
let currentHeader = ''
let currentFooter = ''
let keepWordOpen = false
let tipsEnabled = true
let readOnly = false
let dirty = false
let autoVersionTimer = null

const WORD_FILTER = { name: 'Word Document', extensions: ['docx'] }
const PDF_FILTER = { name: 'PDF', extensions: ['pdf'] }
const MARKDOWN_FILTER = { name: 'Markdown', extensions: ['md', 'markdown'] }

function fileTitle(filePath) {
  return filePath ? path.basename(filePath) : 'Untitled'
}

function currentFormat() {
  return formatOf(currentPath) || 'docx'
}

function refreshTitle() {
  if (!mainWindow) return
  const mark = dirty ? ' — Edited' : ''
  mainWindow.setTitle(`${fileTitle(currentPath)}${mark}`)
  mainWindow.setRepresentedFilename(currentPath || '')
  mainWindow.setDocumentEdited(dirty)
}

function askHtml() {
  return mainWindow?.webContents.executeJavaScript('window.jarvisdocEditor.getHtml()') || Promise.resolve('')
}

function sendState() {
  mainWindow?.webContents.send('doc:state', {
    title: fileTitle(currentPath),
    format: currentFormat(),
    path: currentPath
  })
}

async function confirmDiscard() {
  if (!dirty) return true
  const { response } = await dialog.showMessageBox(mainWindow, {
    type: 'warning',
    buttons: ['Save', 'Discard', 'Cancel'],
    defaultId: 0,
    cancelId: 2,
    message: 'Save changes to this document?',
    detail: `${fileTitle(currentPath)} has unsaved changes.`
  })
  if (response === 2) return false
  if (response === 1) return true
  const html = await askHtml()
  return save(html, false)
}

function prefsFile() {
  return path.join(app.getPath('userData'), 'prefs.json')
}

function loadPrefs() {
  try {
    return JSON.parse(fs.readFileSync(prefsFile(), 'utf8'))
  } catch {
    return {}
  }
}

function savePrefs() {
  try {
    fs.writeFileSync(prefsFile(), JSON.stringify({ tips: tipsEnabled }))
  } catch {
    /* preferences are best effort */
  }
}

function stateFile() {
  return path.join(app.getPath('userData'), 'window-state.json')
}

function loadWindowState() {
  try {
    const saved = JSON.parse(fs.readFileSync(stateFile(), 'utf8'))
    if (Number.isFinite(saved.width) && Number.isFinite(saved.height)) return saved
  } catch {
    /* first run */
  }
  return null
}

function saveWindowState() {
  if (!mainWindow || mainWindow.isDestroyed() || mainWindow.isFullScreen()) return
  try {
    fs.writeFileSync(stateFile(), JSON.stringify(mainWindow.getNormalBounds()))
  } catch {
    /* window state is best effort */
  }
}

const DOC_FILTERS = [
  { name: 'All Documents', extensions: ['docx', 'md', 'markdown', 'txt'] },
  { name: 'Word Document', extensions: ['docx'] },
  { name: 'Markdown', extensions: ['md', 'markdown'] },
  { name: 'Plain Text', extensions: ['txt'] },
  { name: 'All Files', extensions: ['*'] }
]

// Opening the wrong kind of file should explain itself rather than do nothing.
const CANT_OPEN = {
  '.doc': 'Word 97–2003 (.doc) files use a different format. Open it in Word and save it as .docx first.',
  '.rtf': 'Rich Text (.rtf) is a different format. Save it as .docx or .md first.',
  '.pages': 'Pages documents are not Word files. Export it from Pages as Word (.docx) first.',
  '.odt': 'OpenDocument (.odt) is a different format. Save it as .docx or .md first.',
  '.pdf': 'PDF is a finished format, not an editable one. JarvisDoc can write PDFs but not read them.'
}

function refuse(filePath) {
  const ext = path.extname(filePath).toLowerCase()
  return CANT_OPEN[ext] || `JarvisDoc opens Word (.docx), Markdown (.md) and plain text (.txt) files.`
}

async function openPath(filePath) {
  if (!isSupportedFile(filePath)) {
    await dialog.showMessageBox(mainWindow, {
      type: 'warning',
      message: `Cannot open ${path.basename(filePath)}`,
      detail: refuse(filePath),
      buttons: ['OK']
    })
    return
  }
  const buffer = await readFile(filePath)
  const parsed = await fileToHtml(filePath, buffer)
  app.addRecentDocument(filePath)
  currentPath = filePath
  sourceDocx = parsed.format === 'docx' ? buffer : null
  currentPage = parsed.page || { ...DEFAULT_PAGE }
  currentHeader = parsed.headerHtml || ''
  currentFooter = parsed.footerHtml || ''
  readOnly = !(await canWrite(filePath))
  dirty = false
  refreshTitle()
  mainWindow?.webContents.send('doc:load', {
    html: parsed.html,
    title: fileTitle(filePath),
    format: parsed.format,
    path: filePath,
    page: parsed.page,
    headerHtml: parsed.headerHtml,
    footerHtml: parsed.footerHtml,
    readOnly
  })
}

async function restoreVersion(id, html) {
  if (!currentPath) return false
  const found = await versions.read(currentPath, id)
  if (!found) return false
  await snapshot(html, { kind: 'restore', label: 'Before restoring a version' })
  const parsed = await fileToHtml(currentPath, found.buffer)
  sourceDocx = parsed.format === 'docx' ? found.buffer : null
  currentPage = parsed.page || { ...DEFAULT_PAGE }
  currentHeader = parsed.headerHtml || ''
  currentFooter = parsed.footerHtml || ''
  dirty = true
  refreshTitle()
  mainWindow?.webContents.send('doc:load', {
    html: parsed.html,
    title: fileTitle(currentPath),
    format: parsed.format,
    path: currentPath,
    page: parsed.page,
    headerHtml: parsed.headerHtml,
    footerHtml: parsed.footerHtml,
    dirty: true
  })
  return true
}

async function exportVersion(id) {
  if (!currentPath) return false
  const found = await versions.read(currentPath, id)
  if (!found) return false
  const stamp = new Date(found.entry.time)
    .toLocaleString('en-CA', { year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit', hour12: false })
    .replace(/[,:]/g, '')
    .replace(/\s+/g, '-')
  const ext = path.extname(currentPath)
  const base = path.basename(currentPath, ext)
  const picked = await dialog.showSaveDialog(mainWindow, {
    title: 'Export Version',
    defaultPath: path.join(path.dirname(currentPath), `${base} ${stamp}${ext}`),
    filters: filtersFor(ext.toLowerCase())
  })
  if (picked.canceled || !picked.filePath) return false
  await writeFile(picked.filePath, found.buffer)
  mainWindow?.webContents.send('doc:toast', `Exported ${path.basename(picked.filePath)}`)
  return true
}

function filtersFor(forcedExt) {
  if (forcedExt === '.pdf') return [PDF_FILTER]
  if (forcedExt === '.md') return [MARKDOWN_FILTER]
  if (forcedExt === '.docx') return [WORD_FILTER]
  return [WORD_FILTER, PDF_FILTER, MARKDOWN_FILTER]
}

function pageInches(page) {
  const width = parseFloat(page?.width)
  const height = parseFloat(page?.height)
  return {
    width: Number.isFinite(width) && width > 0 ? width : 8.5,
    height: Number.isFinite(height) && height > 0 ? height : 11
  }
}

async function wordBuffer(html) {
  if (!dirty && sourceDocx) return sourceDocx
  const name = currentPath && isDocx(currentPath) ? currentPath : 'Document.docx'
  return htmlToFileBuffer(html, name, { sourceBuffer: sourceDocx })
}

async function currentBuffer(html) {
  if (!currentPath) return null
  if (isDocx(currentPath)) return wordBuffer(html)
  return htmlToFileBuffer(html, currentPath)
}

function notifyVersions() {
  mainWindow?.webContents.send('versions:changed')
}

async function snapshot(html, options) {
  if (!currentPath) return null
  try {
    const buffer = await currentBuffer(html)
    const entry = await versions.add(currentPath, buffer, { ...options, html })
    if (entry) notifyVersions()
    return entry
  } catch {
    return null
  }
}

function scheduleAutoVersion() {
  if (autoVersionTimer) clearInterval(autoVersionTimer)
  autoVersionTimer = setInterval(() => {
    if (!dirty || !currentPath || !mainWindow) return
    void askHtml().then((html) => snapshot(html, { kind: 'save', label: 'Autosaved draft' }))
  }, AUTO_VERSION_MS)
}

async function previewPdf() {
  const size = pageInches(currentPage)
  return mainWindow.webContents.printToPDF({
    printBackground: true,
    preferCSSPageSize: false,
    pageSize: size,
    margins: { marginType: 'none' }
  })
}

async function writePdf(html, dest) {
  const dir = await mkdtemp(path.join(os.tmpdir(), 'jarvisdoc-'))
  const docxPath = path.join(dir, 'document.docx')
  const pdfPath = path.join(dir, 'document.pdf')
  try {
    await writeFile(docxPath, await wordBuffer(html))
    try {
      await convertToPdf(docxPath, pdfPath)
      await writeFile(dest, await readFile(pdfPath))
      return
    } catch {
      // No native converter on this machine, so render the page ourselves.
      const pdf = await previewPdf()
      if (!pdf || pdf.length < 800 || pdf.subarray(0, 5).toString() !== '%PDF-') {
        throw new Error('Could not create a PDF from this document.')
      }
      await writeFile(dest, pdf)
    }
  } finally {
    await rm(dir, { recursive: true, force: true })
  }
}

function docReport(html) {
  return markdownReport(html, {
    page: currentPage,
    headerHtml: currentHeader,
    footerHtml: currentFooter
  })
}

// Converting to Markdown is lossy by nature, so say exactly what goes before writing it.
async function confirmMarkdown(html, dest) {
  const report = docReport(html)
  if (report.ready) return true
  const lines = report.findings.map((item) => `• ${item.label}${item.count > 1 ? ` (${item.count})` : ''} — ${item.detail}`)
  const keeping = currentPath && isDocx(currentPath) ? `\n\n${path.basename(currentPath)} is untouched and keeps all of it.` : ''
  const { response, checkboxChecked } = await dialog.showMessageBox(mainWindow, {
    type: 'info',
    buttons: ['Save as Markdown', 'Cancel'],
    defaultId: 0,
    cancelId: 1,
    message: `Markdown keeps the words and the structure, not the styling.`,
    detail: `These would not come across:\n\n${lines.join('\n')}${keeping}`,
    checkboxLabel: 'Also keep editing the Word file',
    checkboxChecked: Boolean(currentPath && isDocx(currentPath))
  })
  if (response === 1) return false
  keepWordOpen = checkboxChecked
  return true
}

// Node's errno text is useless to a reader, and a failed save needs a way out, not just a sorry.
function saveProblem(err, dest) {
  const name = path.basename(dest)
  switch (err?.code) {
    case 'EACCES':
    case 'EPERM':
      return {
        message: `You do not have permission to change ${name}.`,
        detail: 'The file belongs to another user account on this Mac, or it is locked. Save a copy somewhere in your own folders instead.'
      }
    case 'EROFS':
      return { message: `${name} is on a read-only disk.`, detail: 'Save a copy somewhere you can write to.' }
    case 'ENOSPC':
      return { message: 'The disk is full.', detail: 'Free up some space, or save a copy to another disk.' }
    case 'ENOENT':
      return { message: `The folder for ${name} is no longer there.`, detail: 'It may have been moved or deleted. Save a copy somewhere else.' }
    case 'EBUSY':
      return { message: `${name} is in use by another program.`, detail: 'Close it there, or save a copy under a different name.' }
    default:
      return { message: `Could not save ${name}.`, detail: err instanceof Error ? err.message : String(err) }
  }
}

async function reportSaveFailure(err, dest, html) {
  const problem = saveProblem(err, dest)
  const { response } = await dialog.showMessageBox(mainWindow, {
    type: 'warning',
    buttons: ['Save As…', 'Cancel'],
    defaultId: 0,
    cancelId: 1,
    message: problem.message,
    detail: `${problem.detail}\n\nYour changes are still open and unsaved.`
  })
  if (response !== 0) return false
  return save(html, true)
}

async function canWrite(filePath) {
  try {
    await access(filePath, constants.W_OK)
    return true
  } catch {
    return false
  }
}

function defaultSaveName(forcedExt) {
  if (currentPath) {
    if (forcedExt) return currentPath.replace(/\.[^.]+$/, forcedExt)
    return currentPath
  }
  return path.join(app.getPath('documents'), `Untitled${forcedExt || '.docx'}`)
}

async function save(html, saveAs, forcedExt) {
  let dest = currentPath
  if (saveAs || !dest || forcedExt) {
    const fallback = forcedExt || (isMarkdown(dest) ? '.md' : '.docx')
    const picked = await dialog.showSaveDialog(mainWindow, {
      title: fallback === '.pdf' ? 'Save as PDF' : 'Save Document',
      defaultPath: defaultSaveName(fallback),
      filters: filtersFor(forcedExt)
    })
    if (picked.canceled || !picked.filePath) return false
    dest = picked.filePath
    if (!path.extname(dest)) dest = ensureExt(dest, fallback)
  }
  if (path.extname(dest).toLowerCase() === '.pdf') {
    try {
      await writePdf(html, dest)
    } catch (err) {
      return reportSaveFailure(err, dest, html)
    }
    mainWindow?.webContents.send('doc:toast', `Saved ${path.basename(dest)}`)
    return true
  }
  keepWordOpen = false
  if (isMarkdown(dest) && !(await confirmMarkdown(html, dest))) return false
  const stayOnWord = keepWordOpen && isMarkdown(dest) && currentPath && isDocx(currentPath)
  let saved = null
  try {
    const buffer = isDocx(dest) ? await wordBuffer(html) : await htmlToFileBuffer(html, dest)
    await writeFile(dest, buffer)
    if (isDocx(dest)) sourceDocx = buffer
    saved = buffer
  } catch (err) {
    return reportSaveFailure(err, dest, html)
  }
  app.addRecentDocument(dest)
  readOnly = false
  if (stayOnWord) {
    // They wanted a Markdown copy, not a move; the Word file is still the document.
    mainWindow?.webContents.send('doc:toast', `Saved ${path.basename(dest)} — still editing ${fileTitle(currentPath)}`)
    return true
  }
  currentPath = dest
  dirty = false
  refreshTitle()
  mainWindow?.webContents.send('doc:saved', {
    title: fileTitle(dest),
    format: currentFormat(),
    path: dest
  })
  try {
    const entry = await versions.add(dest, saved, { kind: 'save', html })
    if (entry) notifyVersions()
  } catch {
    /* version history is best effort */
  }
  return true
}

async function newDocument() {
  if (!(await confirmDiscard())) return
  currentPath = null
  sourceDocx = null
  readOnly = false
  currentPage = { ...DEFAULT_PAGE }
  dirty = false
  refreshTitle()
  mainWindow?.webContents.send('doc:load', {
    html: '<p></p>',
    title: 'Untitled',
    format: 'docx',
    path: null,
    page: DEFAULT_PAGE,
    headerHtml: '',
    footerHtml: ''
  })
}

function overlayColors() {
  const dark = nativeTheme.shouldUseDarkColors
  return { color: dark ? '#1b1c1f' : '#e7e5e0', symbolColor: dark ? '#e9e7e3' : '#2c2a26', height: 38 }
}

// Only macOS and Windows can host the title bar inside the page. On Linux the window
// manager draws the frame, so asking for a hidden one would leave no way to close it.
function chromeOptions() {
  if (IS_MAC) return { titleBarStyle: 'hiddenInset', trafficLightPosition: { x: 15, y: 13 } }
  if (IS_WIN) return { titleBarStyle: 'hidden', titleBarOverlay: overlayColors() }
  return {}
}

function createWindow() {
  const saved = loadWindowState()
  mainWindow = new BrowserWindow({
    width: saved?.width ?? 1040,
    height: saved?.height ?? 1080,
    x: saved?.x,
    y: saved?.y,
    minWidth: 900,
    minHeight: 680,
    title: 'Untitled',
    icon: IS_LINUX ? path.join(__dirname, 'build', 'icon.png') : undefined,
    backgroundColor: nativeTheme.shouldUseDarkColors ? '#1b1c1f' : '#e7e5e0',
    ...chromeOptions(),
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: true
    }
  })
  refreshTitle()
  void mainWindow.loadFile(path.join(__dirname, 'renderer', 'index.html'))
  installContextMenu(mainWindow.webContents)

  mainWindow.on('resize', saveWindowState)
  mainWindow.on('move', saveWindowState)

  mainWindow.on('close', (event) => {
    saveWindowState()
    if (!dirty) return
    event.preventDefault()
    void confirmDiscard().then((ok) => {
      if (!ok) return
      dirty = false
      mainWindow?.destroy()
    })
  })

  mainWindow.on('closed', () => {
    mainWindow = null
  })
}

nativeTheme.on('updated', () => {
  if (!mainWindow || mainWindow.isDestroyed()) return
  mainWindow.setBackgroundColor(nativeTheme.shouldUseDarkColors ? '#1b1c1f' : '#e7e5e0')
  if (IS_WIN) mainWindow.setTitleBarOverlay(overlayColors())
})

async function revertToSaved() {
  if (!currentPath) return
  if (dirty) {
    const { response } = await dialog.showMessageBox(mainWindow, {
      type: 'warning',
      buttons: ['Revert', 'Cancel'],
      defaultId: 1,
      cancelId: 1,
      message: `Revert changes to ${fileTitle(currentPath)}?`,
      detail: 'The document will go back to the last saved copy. Recent versions stay in version history.'
    })
    if (response === 1) return
    await snapshot(await askHtml(), { kind: 'restore', label: 'Before reverting to saved' })
  }
  dirty = false
  await openPath(currentPath)
}

function installContextMenu(contents) {
  contents.on('context-menu', (_event, params) => {
    void (async () => {
      const send = (action, value) => () => contents.send('menu', { action, value })
      const at = await contents
        .executeJavaScript(`window.jarvisdocEditor.contextAt(${params.x}, ${params.y})`)
        .catch(() => null)
      const menu = new Menu()
      const add = (options) => menu.append(new MenuItem(options))

      for (const word of params.dictionarySuggestions.slice(0, 5)) {
        add({ label: word, click: () => contents.replaceMisspelling(word) })
      }
      if (params.dictionarySuggestions.length) add({ type: 'separator' })
      if (params.misspelledWord) {
        add({
          label: 'Add to Dictionary',
          click: () => contents.session.addWordToSpellCheckerDictionary(params.misspelledWord)
        })
        add({ type: 'separator' })
      }

      if (params.mediaType === 'image') {
        add({ label: 'Replace Image…', click: () => void insertImage(true) })
        add({ label: 'Delete Image', click: send('deleteImage') })
        add({ type: 'separator' })
      }

      add({ role: 'cut', enabled: params.editFlags.canCut })
      add({ role: 'copy', enabled: params.editFlags.canCopy })
      add({ role: 'paste', enabled: params.editFlags.canPaste })
      add({ type: 'separator' })

      if (at?.inTable) {
        const table = new Menu()
        for (const item of [
          { label: 'Insert Row Above', action: 'rowAbove' },
          { label: 'Insert Row Below', action: 'rowBelow' },
          { label: 'Insert Column Left', action: 'colBefore' },
          { label: 'Insert Column Right', action: 'colAfter' },
          { type: 'separator' },
          { label: 'Delete Row', action: 'delRow' },
          { label: 'Delete Column', action: 'delCol' },
          { label: 'Delete Table', action: 'delTable' }
        ]) {
          table.append(new MenuItem(item.type ? item : { label: item.label, click: send(item.action) }))
        }
        add({ label: 'Table', submenu: table })
        add({ type: 'separator' })
      }

      if (params.selectionText) {
        add({ label: 'Bold', click: send('bold') })
        add({ label: 'Italic', click: send('italic') })
        add({ label: 'Underline', click: send('underline') })
        add({ label: 'Add Link…', click: send('link') })
        add({ label: 'Clear Formatting', click: send('clear') })
        add({ type: 'separator' })
      }
      add({ label: 'Find…', accelerator: 'CmdOrCtrl+F', click: send('find') })

      menu.popup({ window: BrowserWindow.fromWebContents(contents) })
    })()
  })
}

function installMenu() {
  const send = (action, value) => () => mainWindow?.webContents.send('menu', { action, value })
  const appMenu = IS_MAC
    ? [
        {
          label: app.name,
          submenu: [
            { role: 'about' },
            { type: 'separator' },
            { role: 'hide' },
            { role: 'hideOthers' },
            { role: 'unhide' },
            { type: 'separator' },
            { role: 'quit' }
          ]
        }
      ]
    : []
  const template = [
    ...appMenu,
    {
      label: 'File',
      submenu: [
        { label: 'New', accelerator: 'CmdOrCtrl+N', click: () => void newDocument() },
        { label: 'Open…', accelerator: 'CmdOrCtrl+O', click: () => void openFromDialog() },
        {
          label: 'Open Recent',
          role: 'recentDocuments',
          submenu: [{ label: 'Clear Menu', role: 'clearRecentDocuments' }]
        },
        { type: 'separator' },
        {
          label: 'Save',
          accelerator: 'CmdOrCtrl+S',
          click: async () => save(await askHtml(), false)
        },
        {
          label: 'Save As…',
          accelerator: 'Shift+CmdOrCtrl+S',
          click: async () => save(await askHtml(), true)
        },
        {
          label: 'Save As Word (.docx)…',
          click: async () => save(await askHtml(), true, '.docx')
        },
        {
          label: 'Save As PDF…',
          click: async () => save(await askHtml(), true, '.pdf')
        },
        {
          label: 'Save As Markdown (.md)…',
          click: async () => save(await askHtml(), true, '.md')
        },
        { type: 'separator' },
        { label: 'Save a Version…', accelerator: 'Alt+CmdOrCtrl+S', click: send('versionSave') },
        { label: 'Revert to Saved', click: () => void revertToSaved() },
        { type: 'separator' },
        { label: 'Print…', accelerator: 'CmdOrCtrl+P', click: () => void printDocument() },
        { type: 'separator' },
        IS_MAC ? { role: 'close' } : { role: 'quit' }
      ]
    },
    {
      label: 'Edit',
      submenu: [
        { role: 'undo' },
        { role: 'redo' },
        { type: 'separator' },
        { role: 'cut' },
        { role: 'copy' },
        { role: 'paste' },
        { role: 'selectAll' },
        { type: 'separator' },
        { label: 'Find…', accelerator: 'CmdOrCtrl+F', click: send('find') },
        { label: 'Find and Replace…', accelerator: 'Alt+CmdOrCtrl+F', click: send('replace') }
      ]
    },
    {
      label: 'Format',
      submenu: [
        { label: 'Bold', accelerator: 'CmdOrCtrl+B', click: send('bold') },
        { label: 'Italic', accelerator: 'CmdOrCtrl+I', click: send('italic') },
        { label: 'Underline', accelerator: 'CmdOrCtrl+U', click: send('underline') },
        { label: 'Strikethrough', click: send('strike') },
        { label: 'Superscript', click: send('sup') },
        { label: 'Subscript', click: send('sub') },
        { type: 'separator' },
        { label: 'Title', click: send('h1') },
        { label: 'Heading', click: send('h2') },
        { label: 'Subheading', click: send('h3') },
        { label: 'Body', click: send('p') },
        { type: 'separator' },
        { label: 'Bulleted List', click: send('ul') },
        { label: 'Numbered List', click: send('ol') },
        { label: 'Increase Indent', accelerator: 'CmdOrCtrl+]', click: send('indent') },
        { label: 'Decrease Indent', accelerator: 'CmdOrCtrl+[', click: send('outdent') },
        { type: 'separator' },
        { label: 'Align Left', click: send('left') },
        { label: 'Align Center', click: send('center') },
        { label: 'Align Right', click: send('right') },
        { label: 'Justify', click: send('justify') },
        { type: 'separator' },
        { label: 'Clear Formatting', click: send('clear') }
      ]
    },
    {
      label: 'Insert',
      submenu: [
        { label: 'Link…', accelerator: 'CmdOrCtrl+K', click: send('link') },
        { label: 'Image…', click: () => void insertImage() },
        { label: 'Table', click: send('table') },
        {
          label: 'Table Row or Column',
          submenu: [
            { label: 'Insert Row Above', click: send('rowAbove') },
            { label: 'Insert Row Below', click: send('rowBelow') },
            { label: 'Insert Column Left', click: send('colBefore') },
            { label: 'Insert Column Right', click: send('colAfter') },
            { type: 'separator' },
            { label: 'Delete Row', click: send('delRow') },
            { label: 'Delete Column', click: send('delCol') },
            { label: 'Delete Table', click: send('delTable') }
          ]
        }
      ]
    },
    {
      label: 'View',
      submenu: [
        { label: 'Zoom In', accelerator: 'CmdOrCtrl+Plus', click: send('zoomIn') },
        { label: 'Zoom Out', accelerator: 'CmdOrCtrl+-', click: send('zoomOut') },
        { label: 'Actual Size', accelerator: 'CmdOrCtrl+0', click: send('zoomReset') },
        { type: 'separator' },
        { label: 'Version History', accelerator: 'Alt+CmdOrCtrl+V', click: send('versions') },
        { label: 'Markdown View', accelerator: 'Alt+CmdOrCtrl+M', click: send('markdown') },
        {
          label: 'Markdown Tips',
          type: 'checkbox',
          checked: tipsEnabled,
          click: (item) => {
            tipsEnabled = item.checked
            savePrefs()
            mainWindow?.webContents.send('menu', { action: 'tipsToggle', value: tipsEnabled })
          }
        },
        { label: 'Markdown Cheat Sheet', click: send('cheatsheet') },
        { type: 'separator' },
        { role: 'togglefullscreen' },
        { role: 'toggleDevTools' }
      ]
    },
    { role: 'windowMenu' },
    {
      role: 'help',
      submenu: [
        { label: 'Why Markdown?', click: send('guide') },
        { label: 'Markdown Cheat Sheet', click: send('cheatsheet') },
        {
          label: 'JarvisDoc on the Web',
          click: () => void shell.openExternal('https://github.com/sw7ft/jarvisdoc')
        },
        ...(IS_MAC ? [] : [{ type: 'separator' }, { role: 'about' }])
      ]
    }
  ]
  Menu.setApplicationMenu(Menu.buildFromTemplate(template))
}

// Printing goes through the same renderer the page is drawn in, so the print stylesheet
// decides what lands on paper. Panels and chrome are hidden there.
async function printDocument() {
  if (!mainWindow) return
  const inches = pageInches(currentPage)
  return new Promise((resolve) => {
    mainWindow.webContents.print(
      {
        silent: false,
        printBackground: true,
        margins: { marginType: 'none' },
        pageSize: { width: Math.round(inches.width * 25400), height: Math.round(inches.height * 25400) }
      },
      (ok, reason) => {
        if (!ok && reason && reason !== 'cancelled' && reason !== 'Print job canceled') {
          void dialog.showMessageBox(mainWindow, {
            type: 'warning',
            message: 'Could not print this document.',
            detail: `${reason}\n\nYou can use File → Save As PDF instead.`
          })
        }
        resolve(ok)
      }
    )
  })
}

// The guide compares what the document weighs against what its words weigh, so the
// number has to come from the file on disk rather than anything the editor is holding.
function fileSize() {
  if (sourceDocx) return sourceDocx.length
  if (!currentPath) return 0
  try {
    return fs.statSync(currentPath).size
  } catch {
    return 0
  }
}

async function openFromDialog() {
  if (!(await confirmDiscard())) return
  const picked = await dialog.showOpenDialog(mainWindow, {
    title: 'Open Document',
    properties: ['openFile'],
    filters: DOC_FILTERS
  })
  if (picked.canceled || !picked.filePaths[0]) return
  try {
    await openPath(picked.filePaths[0])
  } catch (err) {
    await dialog.showErrorBox('Could not open document', err instanceof Error ? err.message : String(err))
  }
}

function isSupportedFile(filePath) {
  return Boolean(formatOf(filePath))
}

async function insertImage(replace = false) {
  const picked = await dialog.showOpenDialog(mainWindow, {
    title: replace ? 'Replace Image' : 'Insert Image',
    properties: ['openFile'],
    filters: [{ name: 'Images', extensions: ['png', 'jpg', 'jpeg', 'gif', 'webp'] }]
  })
  if (picked.canceled || !picked.filePaths[0]) return
  const filePath = picked.filePaths[0]
  const buffer = await readFile(filePath)
  const ext = path.extname(filePath).slice(1).toLowerCase().replace('jpg', 'jpeg')
  const src = `data:image/${ext};base64,${buffer.toString('base64')}`
  mainWindow?.webContents.send('menu', { action: replace ? 'replaceImage' : 'image', value: src })
}

app.setName('JarvisDoc')

// macOS delivers open-file before the app is ready on a cold launch, so queue until a window exists.
let pendingOpen = null
let readyToOpen = false

async function openRequest(filePath) {
  if (!readyToOpen) {
    pendingOpen = filePath
    return
  }
  if (!(await confirmDiscard())) return
  await openPath(filePath)
}

app.on('open-file', (event, filePath) => {
  event.preventDefault()
  void openRequest(filePath)
})

const gotLock = app.requestSingleInstanceLock()
if (!gotLock) {
  app.quit()
} else {
  app.on('second-instance', (_e, argv) => {
    const file = argv.find((a) => isSupportedFile(a))
    if (mainWindow) {
      if (mainWindow.isMinimized()) mainWindow.restore()
      mainWindow.show()
      mainWindow.focus()
    }
    if (file) void openPath(file)
  })

  app.whenReady().then(() => {
    tipsEnabled = loadPrefs().tips !== false
    ipcMain.handle('doc:ready', () => ({
      title: fileTitle(currentPath),
      path: currentPath,
      format: currentFormat(),
      tips: tipsEnabled
    }))
    ipcMain.on('doc:changed', (_e, nextDirty) => {
      dirty = Boolean(nextDirty)
      refreshTitle()
    })
    ipcMain.handle('doc:new', () => newDocument())
    ipcMain.handle('doc:open', () => openFromDialog())
    ipcMain.handle('doc:save', (_e, html) => save(html, false))
    ipcMain.handle('doc:saveAs', (_e, payload) => save(payload?.html, true, payload?.forcedExt))
    ipcMain.handle('doc:insertImage', () => insertImage())
    ipcMain.handle('doc:openPath', async (_e, filePath) => {
      if (!isSupportedFile(filePath)) return false
      if (!(await confirmDiscard())) return false
      try {
        await openPath(filePath)
        return true
      } catch (err) {
        await dialog.showErrorBox('Could not open document', err instanceof Error ? err.message : String(err))
        return false
      }
    })

    ipcMain.handle('md:report', (_e, html) => {
      const report = docReport(html)
      return {
        ...report,
        markdown: toMarkdown(html || ''),
        format: currentFormat(),
        fileBytes: fileSize()
      }
    })

    ipcMain.handle('versions:list', async () => ({
      path: currentPath,
      entries: currentPath ? await versions.list(currentPath) : []
    }))
    ipcMain.handle('versions:preview', async (_e, id) => {
      if (!currentPath) return null
      const found = await versions.read(currentPath, id)
      if (!found) return null
      const parsed = await fileToHtml(currentPath, found.buffer)
      return { entry: found.entry, html: parsed.html }
    })
    ipcMain.handle('versions:restore', (_e, payload) => restoreVersion(payload?.id, payload?.html))
    ipcMain.handle('versions:delete', async (_e, id) => {
      if (!currentPath) return false
      const done = await versions.remove(currentPath, id)
      if (done) notifyVersions()
      return done
    })
    ipcMain.handle('versions:save', async (_e, payload) => {
      if (!currentPath) {
        await dialog.showMessageBox(mainWindow, {
          type: 'info',
          message: 'Save the document first',
          detail: 'Version history starts once the document has a file on disk.'
        })
        return null
      }
      const entry = await snapshot(payload?.html, { kind: 'manual', label: payload?.label })
      if (entry) mainWindow?.webContents.send('doc:toast', 'Version saved')
      return entry
    })
    ipcMain.handle('versions:export', (_e, id) => exportVersion(id))
    ipcMain.handle('doc:state', () => {
      sendState()
      return { title: fileTitle(currentPath), format: currentFormat() }
    })

    installMenu()
    createWindow()
    scheduleAutoVersion()

    const opened = pendingOpen || process.argv.find((a) => isSupportedFile(a))
    mainWindow?.webContents.once('did-finish-load', () => {
      readyToOpen = true
      const queued = pendingOpen || opened
      pendingOpen = null
      if (queued) void openPath(queued)
    })

    app.on('activate', () => {
      if (BrowserWindow.getAllWindows().length === 0) createWindow()
    })
  })
}

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit()
})

app.on('web-contents-created', (_e, contents) => {
  contents.setWindowOpenHandler(({ url }) => {
    void shell.openExternal(url)
    return { action: 'deny' }
  })
  contents.on('will-navigate', (event, url) => {
    if (url.startsWith('file://') && url.endsWith('renderer/index.html')) return
    event.preventDefault()
    if (/^https?:/i.test(url)) void shell.openExternal(url)
  })
})
