const crypto = require('crypto')
const { mkdir, readFile, writeFile } = require('fs/promises')
const path = require('path')
const cheerio = require('cheerio')
const MarkdownIt = require('markdown-it')
const TurndownService = require('turndown')
const { htmlToDocxBuffer, docxToHtml, DEFAULT_PAGE } = require('./docx-io')

const markdown = new MarkdownIt({
  html: false,
  linkify: true,
  typographer: true,
  breaks: false
})

const turndown = new TurndownService({
  headingStyle: 'atx',
  codeBlockStyle: 'fenced',
  bulletListMarker: '-'
})

function firstRow(node) {
  let table = node.parentNode
  while (table && table.nodeName !== 'TABLE') table = table.parentNode
  return table ? table.querySelector('tr') : null
}

// Turndown has no table support of its own, so tables would otherwise be flattened.
turndown.addRule('tableCell', {
  filter: ['th', 'td'],
  replacement: (content) => ` ${content.replace(/\r?\n+/g, ' ').replace(/\|/g, '\\|').trim()} |`
})

turndown.addRule('tableRow', {
  filter: 'tr',
  replacement: (content, node) => {
    const row = `|${content}`
    if (firstRow(node) !== node) return `\n${row}`
    const columns = node.children.length || 1
    return `\n${row}\n|${' --- |'.repeat(columns)}`
  }
})

turndown.addRule('table', {
  filter: 'table',
  replacement: (content) => `\n\n${content.trim()}\n\n`
})

// Turndown pads list markers to four columns; the tighter form is what people actually type.
turndown.addRule('listItem', {
  filter: 'li',
  replacement: (content, node) => {
    const body = content.replace(/^\n+/, '').replace(/\n+$/, '\n').replace(/\n/gm, '\n  ')
    const parent = node.parentNode
    let marker = '- '
    if (parent.nodeName === 'OL') {
      const start = Number(parent.getAttribute('start')) || 1
      marker = `${start + Array.prototype.indexOf.call(parent.children, node)}. `
    }
    return marker + body + (node.nextSibling && !/\n$/.test(body) ? '\n' : '')
  }
})

turndown.addRule('tabStop', {
  filter: (node) => node.nodeName === 'SPAN' && (node.getAttribute('class') || '').includes('jd-tab'),
  replacement: () => ' '
})

const IMAGE_EXT = { 'image/png': 'png', 'image/jpeg': 'jpg', 'image/gif': 'gif', 'image/webp': 'webp' }

function assetDirFor(filePath) {
  const base = path.basename(filePath, path.extname(filePath))
  return { name: `${base}.assets`, dir: path.join(path.dirname(filePath), `${base}.assets`) }
}

// Images live beside the .md in a sidecar folder so the file stays readable and portable.
async function externaliseImages(html, filePath) {
  const $ = cheerio.load(html || '', { decodeEntities: false })
  const images = $('img[src^="data:"]').toArray()
  if (!images.length) return html
  const { name, dir } = assetDirFor(filePath)
  await mkdir(dir, { recursive: true })
  for (const img of images) {
    const src = $(img).attr('src') || ''
    const match = /^data:([^;,]+);base64,(.+)$/s.exec(src)
    if (!match) continue
    const buffer = Buffer.from(match[2], 'base64')
    const ext = IMAGE_EXT[match[1].toLowerCase()] || 'png'
    const file = `image-${crypto.createHash('sha1').update(buffer).digest('hex').slice(0, 10)}.${ext}`
    await writeFile(path.join(dir, file), buffer)
    $(img).attr('src', `${name}/${file}`)
    $(img).removeAttr('data-rid').removeAttr('data-cx').removeAttr('data-cy')
  }
  return $.html()
}

async function inlineImages(html, filePath) {
  const $ = cheerio.load(html || '', { decodeEntities: false })
  const images = $('img[src]').toArray()
  if (!images.length) return html
  const root = path.dirname(filePath)
  for (const img of images) {
    const src = $(img).attr('src') || ''
    if (!src || /^(data:|https?:)/i.test(src)) continue
    const resolved = path.resolve(root, decodeURIComponent(src))
    if (!resolved.startsWith(path.resolve(root))) continue
    try {
      const buffer = await readFile(resolved)
      const ext = path.extname(resolved).slice(1).toLowerCase().replace('jpg', 'jpeg')
      $(img).attr('src', `data:image/${ext || 'png'};base64,${buffer.toString('base64')}`)
      $(img).attr('data-md-src', src)
    } catch {
      /* a missing image stays a broken reference rather than failing the open */
    }
  }
  return $.html()
}

function toMarkdown(html) {
  return turndown.turndown(html || '')
}

function extOf(filePath) {
  return path.extname(filePath || '').toLowerCase()
}

// A .txt file is Markdown that has not been told so yet, and it round-trips the same way.
function isMarkdown(filePath) {
  const ext = extOf(filePath)
  return ext === '.md' || ext === '.markdown' || ext === '.txt'
}

function isDocx(filePath) {
  return extOf(filePath) === '.docx'
}

function formatOf(filePath) {
  if (isMarkdown(filePath)) return 'md'
  if (isDocx(filePath)) return 'docx'
  return null
}

function ensureExt(filePath, ext) {
  if (path.extname(filePath)) return filePath
  return `${filePath}${ext}`
}

async function fileToHtml(filePath, buffer) {
  if (isMarkdown(filePath)) {
    const rendered = markdown.render(buffer.toString('utf8')).trim() || '<p></p>'
    const html = await inlineImages(rendered, filePath)
    return { html, format: 'md', page: DEFAULT_PAGE, headerHtml: '', footerHtml: '' }
  }
  const parsed = await docxToHtml(buffer)
  return {
    html: parsed.html,
    format: 'docx',
    page: parsed.page || DEFAULT_PAGE,
    headerHtml: parsed.headerHtml || '',
    footerHtml: parsed.footerHtml || ''
  }
}

async function htmlToFileBuffer(html, filePath, options = {}) {
  if (isMarkdown(filePath)) {
    const withAssets = await externaliseImages(html, filePath)
    return Buffer.from(`${toMarkdown(withAssets)}\n`, 'utf8')
  }
  return htmlToDocxBuffer(html, path.basename(filePath, path.extname(filePath)), {
    sourceBuffer: options.sourceBuffer || null
  })
}

module.exports = {
  extOf,
  isMarkdown,
  isDocx,
  formatOf,
  ensureExt,
  fileToHtml,
  htmlToFileBuffer,
  toMarkdown,
  DEFAULT_PAGE
}
