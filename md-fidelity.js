const cheerio = require('cheerio')
const { DEFAULT_PAGE } = require('./ooxml')

// Markdown carries structure, not presentation. Everything listed here survives a .docx
// save but would be flattened the moment the document is written as .md.
function styleHas(style, property) {
  return new RegExp(`(^|;)\\s*${property}\\s*:`, 'i').test(style || '')
}

function countStyled($, property) {
  let total = 0
  $('[style]').each((_, el) => {
    if (styleHas($(el).attr('style'), property)) total += 1
  })
  return total
}

function countAligned($) {
  let total = 0
  $('[style]').each((_, el) => {
    const match = /(^|;)\s*text-align\s*:\s*([a-z]+)/i.exec($(el).attr('style') || '')
    if (match && !['left', 'start'].includes(match[2].toLowerCase())) total += 1
  })
  return total
}

function countIndented($) {
  let total = 0
  $('[style]').each((_, el) => {
    const match = /(^|;)\s*(margin-left|padding-left|text-indent)\s*:\s*([\d.]+)/i.exec($(el).attr('style') || '')
    if (match && parseFloat(match[3]) > 0) total += 1
  })
  return total
}

function richCells($) {
  let total = 0
  $('td, th').each((_, cell) => {
    const blocks = $(cell).children('p, ul, ol, table, blockquote, pre').length
    const paragraphs = $(cell).children('p').length
    if (blocks > 1 || paragraphs > 1 || $(cell).children('ul, ol, table, blockquote, pre').length) total += 1
  })
  return total
}

function samePage(page) {
  if (!page) return true
  return ['width', 'height', 'marginTop', 'marginRight', 'marginBottom', 'marginLeft'].every(
    (key) => String(page[key] ?? '') === String(DEFAULT_PAGE[key] ?? '')
  )
}

function plural(count, one, many) {
  return `${count} ${count === 1 ? one : many}`
}

// Returns the things that would not survive a save to Markdown, loudest first.
function markdownReport(html, context = {}) {
  const $ = cheerio.load(html || '', { decodeEntities: false })
  const findings = []
  const add = (id, label, count, detail) => {
    if (count > 0) findings.push({ id, label, count, detail })
  }

  if (context.headerHtml) {
    add('header', 'Page header', 1, 'Letterhead and running headers have no place to live in Markdown.')
  }
  if (context.footerHtml) {
    add('footer', 'Page footer', 1, 'Footers and page numbers are dropped.')
  }
  if (!samePage(context.page)) {
    add('page', 'Page size and margins', 1, 'Markdown has no page setup; the reader decides the layout.')
  }

  add('font', 'Fonts', countStyled($, 'font-family'), 'Typeface choices are dropped; Markdown uses the reader\u2019s font.')
  add('size', 'Font sizes', countStyled($, 'font-size'), 'Point sizes are dropped. Use headings for emphasis instead.')
  add('color', 'Text colour', countStyled($, 'color'), 'Coloured text becomes plain text.')
  add('highlight', 'Highlighting', countStyled($, 'background-color'), 'Highlights and shading are dropped.')
  add('align', 'Alignment', countAligned($), 'Centred, right-aligned and justified text all become left-aligned.')
  add('indent', 'Indents', countIndented($), 'Custom indents are dropped. Lists and quotes indent on their own.')
  add('tab', 'Tab stops', $('.jd-tab').length, 'Tabs collapse to a single space.')
  add('underline', 'Underline', $('u').length, 'Markdown has no underline. Bold or italic carry the same emphasis.')
  add('supsub', 'Superscript and subscript', $('sup, sub').length, 'These are written out as ordinary text.')
  add('merged', 'Merged table cells', $('[colspan], [rowspan]').length, 'Markdown tables are a plain grid and cannot merge cells.')
  add('richcell', 'Table cells with layout', richCells($), 'Lists and multiple paragraphs inside a cell are flattened to one line.')

  const kept = {
    headings: $('h1, h2, h3, h4, h5, h6').length,
    lists: $('ul, ol').length,
    tables: $('table').length,
    images: $('img').length,
    links: $('a[href]').length,
    emphasis: $('strong, b, em, i, s, code').length
  }

  return {
    ready: findings.length === 0,
    findings,
    kept,
    summary: findings.length
      ? `${plural(findings.length, 'thing', 'things')} would not survive`
      : 'Everything in this document fits Markdown'
  }
}

module.exports = { markdownReport }
