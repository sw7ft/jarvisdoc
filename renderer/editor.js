const editor = document.getElementById('editor')
const page = document.querySelector('.page')
const headerPreview = document.getElementById('header-preview')
const footerPreview = document.getElementById('footer-preview')
const nameEl = document.getElementById('doc-name')
const dirtyDot = document.getElementById('doc-dirty')
const formatEl = document.getElementById('doc-format')
const readOnlyBadge = document.getElementById('doc-readonly')
const countEl = document.getElementById('word-count')
const charsEl = document.getElementById('char-count')
const buttons = [...document.querySelectorAll('.ribbon [data-cmd]')]
const styleSelect = document.getElementById('style-select')
const fontSelect = document.getElementById('font-select')
const sizeSelect = document.getElementById('size-select')
const colorInput = document.getElementById('color-input')
const colorBar = document.getElementById('color-bar')
const highlightInput = document.getElementById('highlight-input')
const highlightBar = document.getElementById('highlight-bar')
const findBar = document.getElementById('find-bar')
const findInput = document.getElementById('find-input')
const findCount = document.getElementById('find-count')
const zoomValue = document.getElementById('zoom-reset')
const toast = document.getElementById('toast')
const panel = document.getElementById('versions')
const panelList = document.getElementById('versions-list')
const panelToggle = document.getElementById('versions-toggle')
const panelCount = document.getElementById('versions-count')
const previewBanner = document.getElementById('preview-banner')
const previewLabel = document.getElementById('preview-label')
const modal = document.getElementById('modal')
const modalForm = document.getElementById('modal-form')
const modalTitle = document.getElementById('modal-title')
const modalInput = document.getElementById('modal-input')

const replaceRow = document.getElementById('replace-row')
const replaceInput = document.getElementById('replace-input')
const mdPanel = document.getElementById('mdpanel')
const mdToggle = document.getElementById('md-toggle')
const mdDot = document.getElementById('md-dot')
const mdState = document.getElementById('md-state')
const mdSource = document.getElementById('md-source')
const mdFit = document.getElementById('md-fit')
const mdLearn = document.getElementById('md-learn')
const mdTabs = [...document.querySelectorAll('.md-tab')]
const hint = document.getElementById('hint')
const hintText = document.getElementById('hint-text')
const hintAction = document.getElementById('hint-action')

const PASSIVE = new Set([
  'find',
  'replace',
  'zoomIn',
  'zoomOut',
  'zoomReset',
  'versions',
  'versionSave',
  'markdown',
  'tipsToggle',
  'cheatsheet',
  'guide'
])

let lastHtml = editor.innerHTML
let zoom = 1
let toastTimer = 0
let docPath = null
let selectedVersion = null
let preview = null
let modalResolve = null
let contextTarget = null
let settleUntil = 0
let mdTimer = 0
let lastReport = null
let hintTimer = 0
let hintId = null
let hintDo = null
let lastHintAt = 0
let pendingTail = null
let pendingSpace = null
const startedAt = Date.now()
let tipsAllowed = true
const migrationOffered = new Set()

function formatLabel(format) {
  return format === 'md' ? 'Markdown · .md' : 'Word · .docx'
}

// The badge names the file in front of you, not the family it is handled as.
function badgeFor(payload) {
  const ext = (String(payload?.title || '').match(/\.[a-z0-9]+$/i) || [''])[0].toLowerCase()
  if (ext === '.txt') return 'Text · .txt'
  if (ext === '.markdown') return 'Markdown · .markdown'
  return formatLabel(payload?.format)
}

function setMeta(payload) {
  if (payload?.title) nameEl.textContent = payload.title
  if (payload?.format) formatEl.textContent = badgeFor(payload)
}

function setDirty(value) {
  dirtyDot.hidden = !value
}

function showToast(message) {
  toast.textContent = message
  toast.hidden = false
  window.clearTimeout(toastTimer)
  toastTimer = window.setTimeout(() => {
    toast.hidden = true
  }, 1600)
}

function textOf() {
  return (editor.innerText || '').replace(/\s+/g, ' ').trim()
}

function updateCounts() {
  const text = textOf()
  const words = text ? text.split(' ').length : 0
  const chars = (editor.innerText || '').replace(/\n+$/, '').length
  countEl.textContent = words === 1 ? '1 word' : `${words} words`
  charsEl.textContent = chars === 1 ? '1 character' : `${chars} characters`
  window.JarvisGuide.note('words', words)
}

function getHtml() {
  const html = editor.innerHTML.replace(/\u200b/g, '').trim()
  return html && html !== '<br>' ? html : '<p></p>'
}

function setBand(el, html) {
  const value = (html || '').trim()
  el.innerHTML = value
  el.hidden = !value
}

function applyPage(pageSetup) {
  if (!pageSetup) return
  const root = document.documentElement
  if (pageSetup.width) root.style.setProperty('--page-width', pageSetup.width)
  if (pageSetup.height) root.style.setProperty('--page-min-height', pageSetup.height)
  if (pageSetup.marginTop) root.style.setProperty('--page-pad-top', pageSetup.marginTop)
  if (pageSetup.marginRight) root.style.setProperty('--page-pad-right', pageSetup.marginRight)
  if (pageSetup.marginBottom) root.style.setProperty('--page-pad-bottom', pageSetup.marginBottom)
  if (pageSetup.marginLeft) root.style.setProperty('--page-pad-left', pageSetup.marginLeft)
}

function setHtml(html) {
  editor.innerHTML = html && html.trim() ? html : '<p></p>'
  lastHtml = editor.innerHTML
  updateCounts()
  editor.focus()
  // Placing the caret makes Chromium tidy the markup, which is not an edit by the reader.
  settleUntil = performance.now() + 500
}

function markDirty() {
  const html = editor.innerHTML
  if (html === lastHtml) return
  lastHtml = html
  updateCounts()
  if (performance.now() < settleUntil) return
  window.JarvisGuide.note('typing')
  setDirty(true)
  window.jarvisdoc.changed(true)
}

function applyZoom(next) {
  zoom = Math.min(2, Math.max(0.5, Math.round(next * 20) / 20))
  document.documentElement.style.setProperty('--zoom', String(zoom))
  zoomValue.textContent = `${Math.round(zoom * 100)}%`
}

function applyFontSize(pt) {
  document.execCommand('styleWithCSS', false, true)
  document.execCommand('fontSize', false, '7')
  for (const font of [...editor.querySelectorAll('font[size="7"]')]) {
    const span = document.createElement('span')
    span.style.fontSize = `${pt}pt`
    span.innerHTML = font.innerHTML
    font.replaceWith(span)
  }
  for (const span of [...editor.querySelectorAll('span[style*="xxx-large"], span[style*="2em"]')]) {
    span.style.fontSize = `${pt}pt`
  }
}

function ask(title, value = '', okLabel = 'OK') {
  modalTitle.textContent = title
  modalInput.value = value
  document.getElementById('modal-ok').textContent = okLabel
  modal.hidden = false
  modalInput.focus()
  modalInput.select()
  return new Promise((resolve) => {
    modalResolve = resolve
  })
}

function closeModal(value) {
  modal.hidden = true
  const resolve = modalResolve
  modalResolve = null
  resolve?.(value)
}

async function insertLink() {
  const selection = window.getSelection()
  const saved = selection && selection.rangeCount ? selection.getRangeAt(0).cloneRange() : null
  const current = document.queryCommandValue('createLink')
  const url = await ask('Link URL', current || 'https://', 'Add Link')
  if (!url) return
  editor.focus()
  if (saved) {
    const live = window.getSelection()
    live.removeAllRanges()
    live.addRange(saved)
  }
  document.execCommand('createLink', false, url)
  markDirty()
}

function insertImage(src) {
  if (!src) {
    void window.jarvisdoc.insertImage()
    return
  }
  const probe = new Image()
  probe.onload = () => {
    let width = probe.naturalWidth || 320
    let height = probe.naturalHeight || 240
    const max = 624
    if (width > max && width > 0) {
      height = Math.max(1, Math.round((height * max) / width))
      width = max
    }
    const img = document.createElement('img')
    img.src = src
    img.alt = ''
    img.width = width
    img.height = height
    document.execCommand('insertHTML', false, img.outerHTML)
    markDirty()
  }
  probe.src = src
}

function insertTable() {
  const cell = '<td><p><br></p></td>'
  const row = `<tr>${cell}${cell}${cell}</tr>`
  document.execCommand('insertHTML', false, `<table><tbody>${row}${row}${row}</tbody></table><p><br></p>`)
}

function countMatches(term) {
  if (!term) return 0
  const hay = (editor.innerText || '').toLowerCase()
  const needle = term.toLowerCase()
  let from = 0
  let total = 0
  for (;;) {
    const at = hay.indexOf(needle, from)
    if (at === -1) return total
    total += 1
    from = at + needle.length
  }
}

function updateFindCount() {
  const term = findInput.value
  if (!term) {
    findCount.textContent = ''
    findCount.classList.remove('none')
    return
  }
  const total = countMatches(term)
  findCount.textContent = total ? `${total}` : 'none'
  findCount.classList.toggle('none', total === 0)
}

function findNext(backwards) {
  const term = findInput.value
  if (!term) return
  window.getSelection()?.removeAllRanges()
  window.find(term, false, Boolean(backwards), true, false, false, false)
}

function showFind() {
  findBar.hidden = false
  updateFindCount()
  findInput.focus()
  findInput.select()
}

function hideFind() {
  findBar.hidden = true
  replaceRow.hidden = true
  editor.focus()
}

function formatTime(ms) {
  return new Date(ms).toLocaleTimeString([], { hour: 'numeric', minute: '2-digit' })
}

function formatDay(ms) {
  const date = new Date(ms)
  const today = new Date()
  const yesterday = new Date(today.getTime() - 86400000)
  const sameDay = (a, b) => a.toDateString() === b.toDateString()
  if (sameDay(date, today)) return 'Today'
  if (sameDay(date, yesterday)) return 'Yesterday'
  return date.toLocaleDateString([], { weekday: 'short', month: 'short', day: 'numeric' })
}

function formatSize(bytes) {
  if (bytes < 1024) return `${bytes} B`
  if (bytes < 1024 * 1024) return `${Math.round(bytes / 1024)} KB`
  return `${(bytes / 1024 / 1024).toFixed(1)} MB`
}

const TAGS = { manual: 'Named', restore: 'Pre-restore', save: '' }

function renderVersions(entries) {
  panelList.textContent = ''
  if (!docPath) {
    panelList.innerHTML =
      '<p class="versions-empty">Save this document to a file and JarvisDoc will keep a version each time you save.</p>'
    return
  }
  if (!entries.length) {
    panelList.innerHTML =
      '<p class="versions-empty">No versions yet. A version is kept every time you save, and you can name one at any point.</p>'
    return
  }

  let day = ''
  for (const entry of entries) {
    const entryDay = formatDay(entry.time)
    if (entryDay !== day) {
      day = entryDay
      const heading = document.createElement('div')
      heading.className = 'versions-day'
      heading.textContent = day
      panelList.append(heading)
    }

    const item = document.createElement('div')
    item.className = 'version'
    item.classList.toggle('selected', entry.id === selectedVersion)
    item.tabIndex = 0

    const top = document.createElement('div')
    top.className = 'version-top'
    const time = document.createElement('span')
    time.className = 'version-time'
    time.textContent = formatTime(entry.time)
    top.append(time)
    const tag = TAGS[entry.kind]
    if (tag) {
      const badge = document.createElement('span')
      badge.className = 'version-tag'
      badge.textContent = tag
      top.append(badge)
    }
    item.append(top)

    if (entry.label) {
      const label = document.createElement('div')
      label.className = 'version-label'
      label.textContent = entry.label
      item.append(label)
    }

    const sub = document.createElement('div')
    sub.className = 'version-sub'
    sub.textContent = [entry.words != null ? `${entry.words} words` : '', formatSize(entry.bytes)]
      .filter(Boolean)
      .join(' · ')
    item.append(sub)

    const actions = document.createElement('div')
    actions.className = 'version-actions'
    actions.hidden = entry.id !== selectedVersion
    actions.append(
      button('View', 'pill', () => viewVersion(entry)),
      button('Restore', 'pill pill-strong', () => restoreVersion(entry)),
      button('Export', 'pill', () => window.jarvisdoc.versions.exportTo(entry.id)),
      button('Delete', 'pill pill-danger', () => deleteVersion(entry))
    )
    item.append(actions)

    item.addEventListener('click', (event) => {
      if (event.target.closest('button')) return
      selectedVersion = selectedVersion === entry.id ? null : entry.id
      renderVersions(entries)
    })
    panelList.append(item)
  }
}

function button(text, className, onClick) {
  const el = document.createElement('button')
  el.type = 'button'
  el.className = className
  el.textContent = text
  el.addEventListener('click', onClick)
  return el
}

async function loadVersions() {
  const data = await window.jarvisdoc.versions.list()
  docPath = data.path
  const entries = data.entries || []
  panelCount.textContent = entries.length
    ? `${entries.length} version${entries.length === 1 ? '' : 's'}`
    : 'No versions'
  renderVersions(entries)
}

function togglePanel(force) {
  const show = force ?? panel.hidden
  panel.hidden = !show
  panelToggle.classList.toggle('on', show)
  if (show) {
    toggleMdPanel(false)
    void loadVersions()
  }
  syncPanelClass()
}

/* ---------- markdown ---------- */

function escapeHtml(text) {
  return String(text).replace(/[&<>]/g, (ch) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;' })[ch])
}

function renderFit(report) {
  if (!report) return
  const head = report.ready
    ? '<div class="md-fit-head">Nothing would be lost</div>'
    : `<div class="md-fit-head">${report.findings.length} ${report.findings.length === 1 ? 'thing' : 'things'} would be lost</div>`
  const rows = report.findings
    .map(
      (item) => `<div class="md-loss">
        <div class="md-loss-top"><span>${escapeHtml(item.label)}</span><span class="md-loss-count">${item.count > 1 ? item.count : ''}</span></div>
        <div class="md-loss-detail">${escapeHtml(item.detail)}</div>
      </div>`
    )
    .join('')
  const keeps = Object.entries(report.kept || {})
    .filter(([, count]) => count > 0)
    .map(([name, count]) => `${count} ${name}`)
  const kept = keeps.length ? `<div class="md-kept">Carried over intact: ${escapeHtml(keeps.join(', '))}.</div>` : ''
  mdFit.innerHTML = head + rows + kept
}

function paintBadge(report) {
  if (!report) return
  mdDot.className = `md-dot ${report.ready ? 'ready' : 'lossy'}`
  mdState.textContent = report.ready ? 'Markdown-ready' : `${report.findings.length} won’t fit`
  mdToggle.title = `${report.summary} · Markdown view  ⌥⌘M`
}

async function refreshMarkdown() {
  if (preview) return
  const report = await window.jarvisdoc.mdReport(getHtml())
  lastReport = report
  mdSource.textContent = report.markdown || ''
  renderFit(report)
  paintBadge(report)
  offerMigration(report)
  window.JarvisGuide.note('report', report)
}

function scheduleMarkdown() {
  window.clearTimeout(mdTimer)
  mdTimer = window.setTimeout(() => void refreshMarkdown(), 900)
}

function toggleMdPanel(force) {
  const show = force ?? mdPanel.hidden
  mdPanel.hidden = !show
  mdToggle.classList.toggle('on', show)
  if (show) {
    panel.hidden = true
    panelToggle.classList.remove('on')
    void refreshMarkdown()
  }
  syncPanelClass()
}

function syncPanelClass() {
  document.body.classList.toggle('panel-open', !mdPanel.hidden || !panel.hidden)
}

function showMdTab(name) {
  for (const tab of mdTabs) tab.classList.toggle('active', tab.dataset.mdTab === name)
  mdSource.hidden = name !== 'source'
  mdFit.hidden = name !== 'fit'
  mdLearn.hidden = name !== 'learn'
}

/* ---------- tips ---------- */

const TIPS = {
  bold: 'Faster next time: type <code>**bold**</code> and it formats itself.',
  italic: 'Faster next time: <code>*italic*</code> turns italic as you type.',
  strike: 'Faster next time: <code>~~struck~~</code> does this as you type.',
  h1: 'Faster next time: start the line with <code># </code>.',
  h2: 'Faster next time: start the line with <code>## </code>.',
  h3: 'Faster next time: start the line with <code>### </code>.',
  h4: 'Faster next time: start the line with <code>#### </code>.',
  ul: 'Faster next time: start the line with <code>- </code>.',
  ol: 'Faster next time: start the line with <code>1. </code>.',
  link: 'Faster next time: type <code>[text](https://…)</code>.'
}

function tipsOn() {
  return tipsAllowed
}

function tipCounts() {
  try {
    return JSON.parse(localStorage.getItem('jd.tipSeen') || '{}')
  } catch {
    return {}
  }
}

function noteTip(id, times) {
  const seen = tipCounts()
  seen[id] = times ?? (seen[id] || 0) + 1
  localStorage.setItem('jd.tipSeen', JSON.stringify(seen))
}

function hideHint() {
  hint.hidden = true
  hintAction.hidden = true
  window.clearTimeout(hintTimer)
}

function showHint(id, html, options = {}) {
  if (!tipsOn() && !options.always) return
  hintId = id
  hintText.innerHTML = html
  hintAction.hidden = !options.actionLabel
  if (options.actionLabel) hintAction.textContent = options.actionLabel
  hintDo = options.action || null
  hint.hidden = false
  lastHintAt = Date.now()
  window.clearTimeout(hintTimer)
  hintTimer = window.setTimeout(hideHint, options.linger || 9000)
}

// A tip should feel like a nudge, not a nag: three sightings each, well spaced.
function maybeTip(cmd) {
  if (!tipsOn() || !TIPS[cmd]) return
  if (Date.now() - lastHintAt < 45000 || Date.now() - startedAt < 4000) return
  const seen = tipCounts()
  if ((seen[cmd] || 0) >= 3) return
  noteTip(cmd)
  showHint(cmd, TIPS[cmd])
}

// Confirm the syntax worked the first couple of times, then get out of the way.
function cheer(id, html) {
  if (!tipsOn()) return
  const seen = tipCounts()
  const key = `did.${id}`
  if ((seen[key] || 0) >= 2) return
  noteTip(key)
  showHint(key, html, { linger: 5000 })
}

function offerMigration(report) {
  if (!tipsOn() || !report?.ready || report.format !== 'docx') return
  if (!docPath || migrationOffered.has(docPath)) return
  const words = Number((countEl.textContent || '').split(' ')[0]) || 0
  if (words < 40) return
  if (Date.now() - lastHintAt < 45000 || Date.now() - startedAt < 8000) return
  const seen = tipCounts()
  if ((seen['migrate'] || 0) >= 4) return
  migrationOffered.add(docPath)
  noteTip('migrate')
  showHint('migrate', 'Nothing in this document needs Word. It would survive as Markdown exactly as it is.', {
    actionLabel: 'Save as Markdown',
    action: () => window.jarvisdoc.saveAs(getHtml(), '.md'),
    linger: 14000
  })
}

/* ---------- markdown input rules ---------- */

const BLOCK_TAGS_MD = 'p, h1, h2, h3, h4, h5, h6, li, blockquote, div, td, th'

function blockAt(node) {
  const el = node?.nodeType === 1 ? node : node?.parentElement
  const block = el?.closest(BLOCK_TAGS_MD)
  return block && editor.contains(block) ? block : null
}

function rangeToCaret(block) {
  const selection = window.getSelection()
  if (!selection?.isCollapsed || !selection.anchorNode) return null
  const range = document.createRange()
  range.selectNodeContents(block)
  try {
    range.setEnd(selection.anchorNode, selection.anchorOffset)
  } catch {
    return null
  }
  return range
}

const BLOCK_RULES = [
  { re: /^#$/, run: () => document.execCommand('formatBlock', false, 'H1'), said: '<code>#</code> makes a title. That is Markdown.' },
  { re: /^##$/, run: () => document.execCommand('formatBlock', false, 'H2'), said: '<code>##</code> makes a heading. That is Markdown.' },
  { re: /^###$/, run: () => document.execCommand('formatBlock', false, 'H3'), said: '<code>###</code> makes a subheading.' },
  { re: /^####$/, run: () => document.execCommand('formatBlock', false, 'H4'), said: '<code>####</code> makes a minor heading.' },
  { re: /^[-*+]$/, run: () => document.execCommand('insertUnorderedList'), said: '<code>-</code> starts a bullet list.' },
  { re: /^\d{1,3}[.)]$/, run: () => document.execCommand('insertOrderedList'), said: '<code>1.</code> starts a numbered list.' },
  { re: /^>$/, run: () => document.execCommand('formatBlock', false, 'BLOCKQUOTE'), said: '<code>&gt;</code> makes a quote.' }
]

function runBlockRule() {
  const selection = window.getSelection()
  const block = blockAt(selection?.anchorNode)
  if (!block || /^(TD|TH)$/.test(block.tagName)) return false
  const inList = block.tagName === 'LI'
  const range = rangeToCaret(block)
  if (!range) return false
  const prefix = range.toString()
  for (const rule of BLOCK_RULES) {
    if (!rule.re.test(prefix)) continue
    if (inList && rule.re.source.includes('#')) return false
    if (inList) return false
    range.deleteContents()
    rule.run()
    markDirty()
    cheer(rule.re.source, rule.said)
    return true
  }
  return false
}

function runEnterRule() {
  const selection = window.getSelection()
  const block = blockAt(selection?.anchorNode)
  if (!block || block.tagName !== 'P') return false
  const text = (block.textContent || '').trim()
  if (/^(-{3,}|\*{3,}|_{3,})$/.test(text)) {
    const rule = document.createElement('hr')
    const next = document.createElement('p')
    next.innerHTML = '<br>'
    block.replaceWith(rule)
    rule.after(next)
    caretTo(next)
    markDirty()
    cheer('hr', '<code>---</code> draws a divider.')
    return true
  }
  if (/^```[a-z]*$/i.test(text)) {
    block.innerHTML = '<br>'
    document.execCommand('formatBlock', false, 'PRE')
    markDirty()
    cheer('pre', '<code>```</code> starts a code block.')
    return true
  }
  return false
}

const INLINE_RULES = [
  { re: /\*\*([^*\s][^*]*)\*\*$/, tag: 'strong', cmd: 'bold', said: '<code>**text**</code> is bold. That is Markdown.' },
  { re: /__([^_\s][^_]*)__$/, tag: 'strong', cmd: 'bold', said: '<code>__text__</code> is bold.' },
  { re: /~~([^~\s][^~]*)~~$/, tag: 's', cmd: 'strikeThrough', said: '<code>~~text~~</code> strikes through.' },
  { re: /(^|[^*\w])\*([^*\s][^*]*)\*$/, tag: 'em', cmd: 'italic', lead: 1, inner: 2, said: '<code>*text*</code> is italic. That is Markdown.' },
  { re: /(^|[^_\w])_([^_\s][^_]*)_$/, tag: 'em', cmd: 'italic', lead: 1, inner: 2, said: '<code>_text_</code> is italic.' },
  { re: /`([^`\s][^`]*)`$/, tag: 'code', said: '<code>`code`</code> formats as code.' },
  { re: /\[([^\]\n]+)\]\((\S+?)\)$/, tag: 'a', said: '<code>[text](url)</code> makes a link.' }
]

// A space typed at the end of a line becomes a non-breaking space. Once real text follows
// it, it is an ordinary word gap again and should be written as one.
function relaxPendingSpace() {
  const node = pendingSpace
  pendingSpace = null
  if (!node?.isConnected || !/^\u00a0\S/.test(node.data)) return
  const selection = window.getSelection()
  const here = selection?.anchorNode === node
  const offset = here ? selection.anchorOffset : 0
  node.data = ` ${node.data.slice(1)}`
  if (!here) return
  const range = document.createRange()
  range.setStart(node, Math.min(offset, node.data.length))
  range.collapse(true)
  selection.removeAllRanges()
  selection.addRange(range)
}

function dropPendingTail() {
  const node = pendingTail
  pendingTail = null
  if (!node?.isConnected || !node.data.includes('\u200b')) return
  const selection = window.getSelection()
  const here = selection?.anchorNode === node
  const offset = here ? selection.anchorOffset : 0
  node.data = node.data.replace(/\u200b/g, '')
  if (!node.data) {
    node.remove()
    return
  }
  pendingSpace = node
  if (!here) return
  const range = document.createRange()
  range.setStart(node, Math.max(0, Math.min(node.data.length, offset - 1)))
  range.collapse(true)
  selection.removeAllRanges()
  selection.addRange(range)
}

function runInlineRule() {
  const selection = window.getSelection()
  const node = selection?.anchorNode
  if (!node || node.nodeType !== 3 || !selection.isCollapsed) return false
  if (node.parentElement?.closest('pre, code')) return false
  const caret = selection.anchorOffset
  const before = node.data.slice(0, caret)
  for (const rule of INLINE_RULES) {
    const found = rule.re.exec(before)
    if (!found) continue
    const consumed = rule.lead ? found[0].length - found[rule.lead].length : found[0].length
    const inner = found[rule.inner || 1]
    const range = document.createRange()
    range.setStart(node, caret - consumed)
    range.setEnd(node, caret)
    range.deleteContents()
    const el = document.createElement(rule.tag)
    if (rule.tag === 'a') {
      el.href = found[2]
      el.textContent = found[1]
    } else {
      el.textContent = inner
    }
    range.insertNode(el)
    // An empty text node is not enough to escape an inline element; Chromium normalises the
    // caret back inside it. A throwaway zero-width space holds the position until the next key.
    const tail = document.createTextNode('\u200b')
    el.parentNode.insertBefore(tail, el.nextSibling)
    pendingTail = tail
    const after = document.createRange()
    after.setStart(tail, 1)
    after.collapse(true)
    selection.removeAllRanges()
    selection.addRange(after)
    // The caret sits just outside the new element; drop any style it inherited.
    if (rule.cmd && document.queryCommandState(rule.cmd)) document.execCommand(rule.cmd)
    markDirty()
    cheer(rule.tag + (rule.cmd || ''), rule.said)
    return true
  }
  return false
}


async function viewVersion(entry) {
  const data = await window.jarvisdoc.versions.preview(entry.id)
  if (!data) return
  if (!preview) preview = { html: editor.innerHTML }
  preview.id = entry.id
  editor.innerHTML = data.html
  editor.setAttribute('contenteditable', 'false')
  page.classList.add('previewing')
  previewLabel.textContent = `Viewing the version from ${formatDay(entry.time)} at ${formatTime(entry.time)}`
  previewBanner.hidden = false
  updateCounts()
}

function endPreview() {
  if (!preview) return
  editor.innerHTML = preview.html
  editor.setAttribute('contenteditable', 'true')
  page.classList.remove('previewing')
  previewBanner.hidden = true
  preview = null
  updateCounts()
  editor.focus()
}

async function restoreVersion(entry) {
  const id = entry?.id || preview?.id
  if (!id) return
  const live = preview ? preview.html : editor.innerHTML
  preview = null
  previewBanner.hidden = true
  page.classList.remove('previewing')
  editor.setAttribute('contenteditable', 'true')
  await window.jarvisdoc.versions.restore(id, live)
  showToast('Version restored — save to keep it')
}

async function deleteVersion(entry) {
  if (preview?.id === entry.id) endPreview()
  await window.jarvisdoc.versions.remove(entry.id)
  if (selectedVersion === entry.id) selectedVersion = null
}

async function saveVersion() {
  const label = await ask('Name this version', '', 'Save Version')
  if (label === null) return
  await window.jarvisdoc.versions.save(label.trim(), getHtml())
}

function cellAt(node) {
  const el = node?.nodeType === 1 ? node : node?.parentElement
  const cell = el?.closest('td, th')
  return cell && editor.contains(cell) ? cell : null
}

function currentCell() {
  return cellAt(contextTarget) || cellAt(window.getSelection()?.anchorNode)
}

function cellIndex(cell) {
  return [...cell.parentElement.children].indexOf(cell)
}

function tableRows(table) {
  return [...table.querySelectorAll('tr')]
}

function blankCell(like) {
  const cell = document.createElement(like?.tagName === 'TH' ? 'th' : 'td')
  cell.innerHTML = '<br>'
  return cell
}

function caretTo(cell) {
  if (!cell) return
  const range = document.createRange()
  range.selectNodeContents(cell)
  range.collapse(true)
  const selection = window.getSelection()
  selection?.removeAllRanges()
  selection?.addRange(range)
}

function insertRow(after) {
  const cell = currentCell()
  if (!cell) return
  const row = cell.parentElement
  const clone = document.createElement('tr')
  for (const source of row.children) clone.append(blankCell(source))
  row.parentElement.insertBefore(clone, after ? row.nextSibling : row)
  caretTo(clone.children[cellIndex(cell)] || clone.children[0])
}

function insertColumn(after) {
  const cell = currentCell()
  if (!cell) return
  const index = cellIndex(cell)
  for (const row of tableRows(cell.closest('table'))) {
    const at = row.children[index]
    const fresh = blankCell(at)
    if (at) row.insertBefore(fresh, after ? at.nextSibling : at)
    else row.append(fresh)
  }
  caretTo(cell.parentElement.children[after ? index + 1 : index])
}

function deleteRow() {
  const cell = currentCell()
  if (!cell) return
  const table = cell.closest('table')
  const row = cell.parentElement
  if (tableRows(table).length <= 1) return deleteTable()
  const next = row.nextElementSibling || row.previousElementSibling
  row.remove()
  caretTo(next?.children[0])
}

function deleteColumn() {
  const cell = currentCell()
  if (!cell) return
  const table = cell.closest('table')
  const index = cellIndex(cell)
  if (cell.parentElement.children.length <= 1) return deleteTable()
  for (const row of tableRows(table)) row.children[index]?.remove()
  caretTo(tableRows(table)[0]?.children[Math.max(0, index - 1)])
}

function deleteTable() {
  const cell = currentCell()
  const table = cell?.closest('table')
  if (!table) return
  const after = table.nextElementSibling
  table.remove()
  caretTo(after)
}

function deleteImage() {
  if (contextTarget?.tagName === 'IMG') contextTarget.remove()
}

function replaceImage(src) {
  const img = contextTarget?.tagName === 'IMG' ? contextTarget : null
  if (!img || !src) return
  img.removeAttribute('data-rid')
  img.removeAttribute('data-drawing')
  img.src = src
}

function nextCell(cell, backwards) {
  const table = cell.closest('table')
  const cells = [...table.querySelectorAll('td, th')]
  const at = cells.indexOf(cell)
  if (backwards) return cells[at - 1] || null
  if (at < cells.length - 1) return cells[at + 1]
  insertRow(true)
  return cell.parentElement.nextElementSibling?.children[0] || null
}

function handleTab(event) {
  const cell = cellAt(window.getSelection()?.anchorNode)
  if (cell) {
    event.preventDefault()
    caretTo(nextCell(cell, event.shiftKey))
    markDirty()
    return
  }
  const item = selectionElement()?.closest('li')
  if (item && editor.contains(item)) {
    event.preventDefault()
    document.execCommand(event.shiftKey ? 'outdent' : 'indent')
    markDirty()
    return
  }
  event.preventDefault()
  if (event.shiftKey) return
  document.execCommand('insertHTML', false, '<span class="jd-tab">\t</span>')
  // execCommand copies the caret's computed style onto the new span, CSS variables and all.
  for (const span of editor.querySelectorAll('span.jd-tab[style]')) span.removeAttribute('style')
  markDirty()
}

function showReplace() {
  replaceRow.hidden = false
  showFind()
}

function replaceOne() {
  const term = findInput.value
  if (!term) return
  const selection = window.getSelection()
  const matched = selection?.toString().toLowerCase() === term.toLowerCase()
  if (!matched) {
    findNext(false)
    return
  }
  editor.focus()
  document.execCommand('insertText', false, replaceInput.value)
  markDirty()
  updateFindCount()
  findNext(false)
}

function replaceAll() {
  const term = findInput.value
  if (!term) return
  editor.focus()
  window.getSelection()?.removeAllRanges()
  let count = 0
  while (window.find(term, false, false, false, false, false, false)) {
    document.execCommand('insertText', false, replaceInput.value)
    count += 1
    if (count > 5000) break
  }
  if (count) markDirty()
  updateFindCount()
  showToast(count ? `Replaced ${count} ${count === 1 ? 'match' : 'matches'}` : 'No matches')
}

function format(cmd, value) {
  if (!PASSIVE.has(cmd)) editor.focus()
  document.execCommand('styleWithCSS', false, true)
  const map = {
    bold: () => document.execCommand('bold'),
    italic: () => document.execCommand('italic'),
    underline: () => document.execCommand('underline'),
    strike: () => document.execCommand('strikeThrough'),
    sup: () => document.execCommand('superscript'),
    sub: () => document.execCommand('subscript'),
    h1: () => document.execCommand('formatBlock', false, 'H1'),
    h2: () => document.execCommand('formatBlock', false, 'H2'),
    h3: () => document.execCommand('formatBlock', false, 'H3'),
    h4: () => document.execCommand('formatBlock', false, 'H4'),
    p: () => document.execCommand('formatBlock', false, 'P'),
    ul: () => document.execCommand('insertUnorderedList'),
    ol: () => document.execCommand('insertOrderedList'),
    indent: () => document.execCommand('indent'),
    outdent: () => document.execCommand('outdent'),
    left: () => document.execCommand('justifyLeft'),
    center: () => document.execCommand('justifyCenter'),
    right: () => document.execCommand('justifyRight'),
    justify: () => document.execCommand('justifyFull'),
    clear: () => document.execCommand('removeFormat'),
    font: () => document.execCommand('fontName', false, value),
    size: () => applyFontSize(value),
    color: () => document.execCommand('foreColor', false, value),
    highlight: () => document.execCommand('hiliteColor', false, value),
    link: () => insertLink(),
    image: () => insertImage(value),
    table: () => insertTable(),
    find: () => showFind(),
    replace: () => showReplace(),
    rowAbove: () => insertRow(false),
    rowBelow: () => insertRow(true),
    colBefore: () => insertColumn(false),
    colAfter: () => insertColumn(true),
    delRow: () => deleteRow(),
    delCol: () => deleteColumn(),
    delTable: () => deleteTable(),
    deleteImage: () => deleteImage(),
    replaceImage: () => replaceImage(value),
    zoomIn: () => applyZoom(zoom + 0.1),
    zoomOut: () => applyZoom(zoom - 0.1),
    zoomReset: () => applyZoom(1),
    versions: () => togglePanel(),
    versionSave: () => void saveVersion(),
    markdown: () => toggleMdPanel(),
    tipsToggle: () => {
      tipsAllowed = Boolean(value)
      hideHint()
      window.JarvisGuide.setEnabled(tipsAllowed)
      showToast(tipsAllowed ? 'Markdown tips on' : 'Markdown tips off')
    },
    cheatsheet: () => {
      toggleMdPanel(true)
      showMdTab('learn')
    },
    guide: () => window.JarvisGuide.summon()
  }
  map[cmd]?.()
  if (PASSIVE.has(cmd)) return
  if (cmd === 'image' && !value) return
  markDirty()
  refreshToolbar()
}

// The same command, run from the toolbar or a menu, is the moment to teach its syntax.
function formatFromUi(cmd, value) {
  format(cmd, value)
  maybeTip(cmd)
}

function selectionElement() {
  const selection = window.getSelection()
  const node = selection?.anchorNode
  if (!node) return null
  const el = node.nodeType === 1 ? node : node.parentElement
  return el && editor.contains(el) ? el : null
}

function matchOption(select, value) {
  const wanted = String(value).toLowerCase()
  const option = [...select.options].find((item) => item.value.toLowerCase() === wanted)
  if (option) select.value = option.value
}

function syncFields() {
  const el = selectionElement()
  if (!el) return
  const styles = window.getComputedStyle(el)
  const family = styles.fontFamily.split(',')[0].replace(/["']/g, '').trim()
  matchOption(fontSelect, family)
  const px = parseFloat(styles.fontSize)
  if (Number.isFinite(px)) matchOption(sizeSelect, String(Math.round((px * 72) / 96)))
  const block = String(document.queryCommandValue('formatBlock') || '').toLowerCase()
  matchOption(styleSelect, ['h1', 'h2', 'h3', 'h4'].includes(block) ? block : 'p')
}

function refreshToolbar() {
  const states = {
    bold: document.queryCommandState('bold'),
    italic: document.queryCommandState('italic'),
    underline: document.queryCommandState('underline'),
    strike: document.queryCommandState('strikeThrough'),
    sup: document.queryCommandState('superscript'),
    sub: document.queryCommandState('subscript'),
    ul: document.queryCommandState('insertUnorderedList'),
    ol: document.queryCommandState('insertOrderedList'),
    left: document.queryCommandState('justifyLeft'),
    center: document.queryCommandState('justifyCenter'),
    right: document.queryCommandState('justifyRight'),
    justify: document.queryCommandState('justifyFull')
  }
  for (const btn of buttons) {
    btn.classList.toggle('active', Boolean(states[btn.dataset.cmd]))
  }
  syncFields()
}

editor.addEventListener('input', markDirty)
editor.addEventListener('input', () => {
  relaxPendingSpace()
  dropPendingTail()
  runInlineRule()
  scheduleMarkdown()
})
editor.addEventListener('keydown', (event) => {
  if (event.metaKey || event.ctrlKey || event.altKey) return
  if (event.key === 'Tab') {
    handleTab(event)
    return
  }
  if (event.key === ' ' && runBlockRule()) event.preventDefault()
  if (event.key === 'Enter' && runEnterRule()) event.preventDefault()
})
editor.addEventListener('keyup', refreshToolbar)
editor.addEventListener('mouseup', refreshToolbar)
editor.addEventListener('focus', refreshToolbar)

page?.addEventListener('mousedown', (event) => {
  if (event.target.closest('#header-preview, #footer-preview, #editor')) return
  editor.focus()
})

editor.addEventListener('paste', (event) => {
  const items = [...(event.clipboardData?.items || [])]
  const imageItem = items.find((item) => item.type.startsWith('image/'))
  if (imageItem) {
    event.preventDefault()
    const file = imageItem.getAsFile()
    if (!file) return
    const reader = new FileReader()
    reader.onload = () => insertImage(String(reader.result || ''))
    reader.readAsDataURL(file)
    return
  }
  event.preventDefault()
  const html = event.clipboardData?.getData('text/html')
  if (html) {
    const clean = html
      .replace(/<script[\s\S]*?>[\s\S]*?<\/script>/gi, '')
      .replace(/on\w+="[^"]*"/gi, '')
    document.execCommand('insertHTML', false, clean)
    return
  }
  document.execCommand('insertText', false, event.clipboardData?.getData('text/plain') || '')
})

for (const btn of buttons) {
  btn.addEventListener('click', () => formatFromUi(btn.dataset.cmd))
}

styleSelect.addEventListener('change', () => formatFromUi(styleSelect.value))
fontSelect.addEventListener('change', () => format('font', fontSelect.value))
sizeSelect.addEventListener('change', () => format('size', sizeSelect.value))

colorInput.addEventListener('input', () => {
  colorBar.style.background = colorInput.value
  format('color', colorInput.value)
})
highlightInput.addEventListener('input', () => {
  highlightBar.style.background = highlightInput.value
  format('highlight', highlightInput.value)
})

document.getElementById('zoom-in').addEventListener('click', () => applyZoom(zoom + 0.1))
document.getElementById('zoom-out').addEventListener('click', () => applyZoom(zoom - 0.1))
zoomValue.addEventListener('click', () => applyZoom(1))

document.getElementById('find-next').addEventListener('click', () => findNext(false))
document.getElementById('find-prev').addEventListener('click', () => findNext(true))
document.getElementById('find-close').addEventListener('click', hideFind)
document.getElementById('replace-one').addEventListener('click', replaceOne)
document.getElementById('replace-all').addEventListener('click', replaceAll)
replaceInput.addEventListener('keydown', (event) => {
  if (event.key === 'Enter') {
    event.preventDefault()
    replaceOne()
  }
  if (event.key === 'Escape') {
    event.preventDefault()
    hideFind()
  }
})
findInput.addEventListener('input', updateFindCount)
findInput.addEventListener('keydown', (event) => {
  if (event.key === 'Enter') {
    event.preventDefault()
    findNext(event.shiftKey)
  }
  if (event.key === 'Escape') {
    event.preventDefault()
    hideFind()
  }
})

mdToggle.addEventListener('click', () => toggleMdPanel())
document.getElementById('md-close').addEventListener('click', () => toggleMdPanel(false))
document.getElementById('md-copy').addEventListener('click', () => {
  void navigator.clipboard.writeText(mdSource.textContent || '')
  showToast('Markdown copied')
})
for (const tab of mdTabs) tab.addEventListener('click', () => showMdTab(tab.dataset.mdTab))

hintAction.addEventListener('click', () => {
  const run = hintDo
  hideHint()
  run?.()
})
document.getElementById('hint-got').addEventListener('click', () => {
  if (hintId) noteTip(hintId, 99)
  hideHint()
})
document.getElementById('hint-close').addEventListener('click', hideHint)

panelToggle.addEventListener('click', () => togglePanel())
document.getElementById('versions-close').addEventListener('click', () => togglePanel(false))
document.getElementById('version-save').addEventListener('click', () => void saveVersion())
document.getElementById('preview-done').addEventListener('click', endPreview)
document.getElementById('preview-restore').addEventListener('click', () => void restoreVersion())

modalForm.addEventListener('submit', (event) => {
  event.preventDefault()
  closeModal(modalInput.value)
})
document.getElementById('modal-cancel').addEventListener('click', () => closeModal(null))
modal.addEventListener('mousedown', (event) => {
  if (event.target === modal) closeModal(null)
})

document.addEventListener('keydown', (event) => {
  if (event.key !== 'Escape') return
  if (!hint.hidden) {
    hideHint()
    return
  }
  if (window.JarvisGuide.isOpen()) {
    window.JarvisGuide.close()
    return
  }
  if (!modal.hidden) {
    event.preventDefault()
    closeModal(null)
    return
  }
  if (preview) {
    event.preventDefault()
    endPreview()
    return
  }
  if (!findBar.hidden) hideFind()
})

document.body.dataset.platform = window.jarvisdoc.platform || 'darwin'
colorBar.style.background = colorInput.value
highlightBar.style.background = highlightInput.value
applyZoom(1)

const DOC_FILE = /\.(docx|md|markdown|txt)$/i

function dropFiles(event) {
  return [...(event.dataTransfer?.files || [])]
}

document.addEventListener('dragover', (event) => {
  if (!event.dataTransfer?.types.includes('Files')) return
  event.preventDefault()
  event.dataTransfer.dropEffect = 'copy'
  document.body.classList.add('dropping')
})
document.addEventListener('dragleave', (event) => {
  if (event.relatedTarget) return
  document.body.classList.remove('dropping')
})
document.addEventListener('drop', (event) => {
  const files = dropFiles(event)
  if (!files.length) return
  event.preventDefault()
  document.body.classList.remove('dropping')
  const image = files.find((file) => file.type.startsWith('image/'))
  if (image) {
    const reader = new FileReader()
    reader.onload = () => insertImage(String(reader.result || ''))
    reader.readAsDataURL(image)
    return
  }
  const doc = files.find((file) => DOC_FILE.test(file.name))
  if (!doc) {
    showToast('Only .docx, .md, .txt and images can be dropped here')
    return
  }
  const filePath = window.jarvisdoc.pathForFile(doc)
  if (filePath) void window.jarvisdoc.openPath(filePath)
})

window.jarvisdocEditor = {
  getHtml,
  setHtml,
  contextAt(x, y) {
    contextTarget = document.elementFromPoint(x, y)
    return { inTable: Boolean(contextTarget?.closest('#editor table')) }
  }
}

window.jarvisdoc.onMenu((payload) => format(payload.action, payload.value))
window.jarvisdoc.onLoad((payload) => {
  preview = null
  previewBanner.hidden = true
  page.classList.remove('previewing')
  editor.setAttribute('contenteditable', 'true')
  selectedVersion = null
  docPath = payload.path || null
  applyPage(payload.page)
  setBand(headerPreview, payload.headerHtml)
  setBand(footerPreview, payload.footerHtml)
  setHtml(payload.html)
  setMeta(payload)
  setDirty(Boolean(payload.dirty))
  readOnlyBadge.hidden = !payload.readOnly
  if (payload.readOnly) showToast('Read-only file — saving will ask for a new name')
  window.jarvisdoc.changed(Boolean(payload.dirty))
  hideHint()
  window.JarvisGuide.note('opened', payload)
  void loadVersions()
  void refreshMarkdown()
})
window.jarvisdoc.onToast((message) => showToast(message))
window.jarvisdoc.onSaved((payload) => {
  lastHtml = editor.innerHTML
  docPath = payload?.path || docPath
  setMeta(payload)
  setDirty(false)
  readOnlyBadge.hidden = true
  showToast(payload?.title ? `Saved ${payload.title}` : 'Saved')
  void refreshMarkdown()
})
window.jarvisdoc.onVersionsChanged(() => {
  void loadVersions()
})

window.JarvisGuide.mount({
  saveAsMarkdown: () => window.jarvisdoc.saveAs(getHtml(), '.md'),
  showMarkdown: () => {
    toggleMdPanel(true)
    showMdTab('source')
  },
  toast: showToast,
  // One voice at a time: a syntax hint and the guide never share the screen.
  busy: () => !hint.hidden || !modal.hidden || !findBar.hidden || Boolean(preview)
})

window.jarvisdoc.ready().then((info) => {
  docPath = info?.path || null
  tipsAllowed = info?.tips !== false
  window.JarvisGuide.setEnabled(tipsAllowed)
  setMeta(info)
  updateCounts()
  void loadVersions()
  void refreshMarkdown()
  editor.focus()
})
