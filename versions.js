const { app } = require('electron')
const crypto = require('crypto')
const path = require('path')
const { mkdir, readFile, writeFile, rm } = require('fs/promises')

const KEEP_AUTOMATIC = 40
const KEEP_TOTAL = 150

function storeRoot() {
  return path.join(app.getPath('userData'), 'versions')
}

function folderFor(docPath) {
  const key = crypto.createHash('sha1').update(path.resolve(docPath)).digest('hex').slice(0, 16)
  return path.join(storeRoot(), key)
}

function indexPath(dir) {
  return path.join(dir, 'index.json')
}

async function readIndex(dir) {
  try {
    const raw = await readFile(indexPath(dir), 'utf8')
    const parsed = JSON.parse(raw)
    return { docPath: parsed.docPath || '', entries: Array.isArray(parsed.entries) ? parsed.entries : [] }
  } catch {
    return { docPath: '', entries: [] }
  }
}

async function writeIndex(dir, index) {
  await mkdir(dir, { recursive: true })
  await writeFile(indexPath(dir), JSON.stringify(index, null, 2), 'utf8')
}

function wordsFromHtml(html) {
  const text = String(html || '')
    .replace(/<(script|style)[\s\S]*?<\/\1>/gi, ' ')
    .replace(/<[^>]*>/g, ' ')
    .replace(/&nbsp;/gi, ' ')
    .replace(/&[a-z]+;|&#\d+;/gi, '')
    .replace(/\s+/g, ' ')
    .trim()
  return text ? text.split(' ').length : 0
}

function newId() {
  return `${Date.now().toString(36)}${crypto.randomBytes(3).toString('hex')}`
}

async function list(docPath) {
  if (!docPath) return []
  const { entries } = await readIndex(folderFor(docPath))
  return [...entries].sort((a, b) => b.time - a.time)
}

async function prune(dir, index) {
  const keep = []
  const drop = []
  const sorted = [...index.entries].sort((a, b) => b.time - a.time)
  let automatic = 0
  for (const entry of sorted) {
    if (entry.kind === 'manual') {
      keep.push(entry)
      continue
    }
    automatic += 1
    if (automatic <= KEEP_AUTOMATIC) keep.push(entry)
    else drop.push(entry)
  }
  const ordered = keep.sort((a, b) => b.time - a.time)
  while (ordered.length > KEEP_TOTAL) drop.push(ordered.pop())
  for (const entry of drop) {
    await rm(path.join(dir, entry.file), { force: true })
  }
  index.entries = ordered
}

async function add(docPath, buffer, options = {}) {
  if (!docPath || !buffer?.length) return null
  const dir = folderFor(docPath)
  const index = await readIndex(dir)
  index.docPath = docPath

  const digest = crypto.createHash('sha1').update(buffer).digest('hex')
  const newest = [...index.entries].sort((a, b) => b.time - a.time)[0]
  if (newest && newest.digest === digest && options.kind !== 'manual') return null

  const id = newId()
  const ext = path.extname(docPath) || '.docx'
  const file = `${id}${ext}`
  await mkdir(dir, { recursive: true })
  await writeFile(path.join(dir, file), buffer)

  const entry = {
    id,
    file,
    digest,
    time: Date.now(),
    bytes: buffer.length,
    kind: options.kind === 'manual' ? 'manual' : options.kind === 'restore' ? 'restore' : 'save',
    label: String(options.label || '').slice(0, 120),
    words: options.html != null ? wordsFromHtml(options.html) : null,
    name: path.basename(docPath)
  }
  index.entries.push(entry)
  await prune(dir, index)
  await writeIndex(dir, index)
  return entry
}

async function read(docPath, id) {
  if (!docPath || !id) return null
  const dir = folderFor(docPath)
  const { entries } = await readIndex(dir)
  const entry = entries.find((item) => item.id === id)
  if (!entry) return null
  try {
    return { entry, buffer: await readFile(path.join(dir, entry.file)) }
  } catch {
    return null
  }
}

async function remove(docPath, id) {
  if (!docPath || !id) return false
  const dir = folderFor(docPath)
  const index = await readIndex(dir)
  const entry = index.entries.find((item) => item.id === id)
  if (!entry) return false
  await rm(path.join(dir, entry.file), { force: true })
  index.entries = index.entries.filter((item) => item.id !== id)
  await writeIndex(dir, index)
  return true
}

module.exports = { add, list, read, remove, wordsFromHtml }
