const { load } = require('cheerio')
const JSZip = require('jszip')
const { Document, Packer, Paragraph } = require('docx')

const EMU_PER_PX = 9525
const PAGE_CONTENT_PX = 624

const DEFAULT_PAGE = {
  width: '8.5in',
  height: '11in',
  marginTop: '1in',
  marginRight: '1in',
  marginBottom: '1in',
  marginLeft: '1in'
}

const HIGHLIGHTS = {
  yellow: 'FFFF00',
  green: '00FF00',
  cyan: '00FFFF',
  magenta: 'FF00FF',
  blue: '0000FF',
  red: 'FF0000',
  darkBlue: '00008B',
  darkCyan: '008B8B',
  darkGreen: '006400',
  darkMagenta: '8B008B',
  darkRed: '8B0000',
  darkYellow: '808000',
  darkGray: 'A9A9A9',
  lightGray: 'D3D3D3',
  black: '000000',
  white: 'FFFFFF'
}

const BLOCK_TAGS = new Set([
  'p', 'div', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6',
  'blockquote', 'ul', 'ol', 'table', 'section', 'article', 'pre'
])

function escXml(value) {
  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
}

function escAttr(value) {
  return escXml(value).replace(/"/g, '&quot;')
}

function escHtml(value) {
  return escXml(value)
}

function onOff($, el) {
  if (!el || !$(el).length) return false
  const val = String($(el).attr('w:val') ?? '').toLowerCase()
  return val !== '0' && val !== 'false' && val !== 'off'
}

function inches(twips) {
  const n = Number(twips)
  if (!Number.isFinite(n) || n <= 0) return '1in'
  return `${(n / 1440).toFixed(3).replace(/0+$/, '').replace(/\.$/, '')}in`
}

function cssToHex(raw) {
  const value = String(raw || '').trim()
  const hex = /^#([0-9a-f]{3}|[0-9a-f]{6})$/i.exec(value)
  if (hex) {
    const h = hex[1]
    if (h.length === 3) return (h[0] + h[0] + h[1] + h[1] + h[2] + h[2]).toUpperCase()
    return h.toUpperCase()
  }
  const rgb = /^rgba?\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)/i.exec(value)
  if (rgb) {
    return [rgb[1], rgb[2], rgb[3]]
      .map((n) => Number(n).toString(16).padStart(2, '0'))
      .join('')
      .toUpperCase()
  }
  const named = {
    black: '000000', white: 'FFFFFF', red: 'FF0000', blue: '0000FF',
    green: '008000', yellow: 'FFFF00', gray: '808080', grey: '808080'
  }
  return named[value.toLowerCase()]
}

function highlightName(hex) {
  if (!hex) return null
  const up = hex.toUpperCase()
  return Object.keys(HIGHLIGHTS).find((name) => HIGHLIGHTS[name] === up) || null
}

function toTwips(raw) {
  const match = /^(-?[\d.]+)\s*(px|pt|in|cm|mm|em)?/i.exec(String(raw || '').trim())
  if (!match) return null
  const n = parseFloat(match[1])
  if (!Number.isFinite(n)) return null
  const unit = (match[2] || 'px').toLowerCase()
  if (unit === 'pt') return Math.round(n * 20)
  if (unit === 'in') return Math.round(n * 1440)
  if (unit === 'cm') return Math.round((n * 1440) / 2.54)
  if (unit === 'mm') return Math.round((n * 1440) / 25.4)
  if (unit === 'em') return Math.round(n * 240)
  return Math.round(n * 15)
}

function twipsToPt(twips) {
  return Math.round((Number(twips) / 20) * 100) / 100
}

function mimeOf(name) {
  const ext = String(name || '').split('.').pop().toLowerCase()
  if (ext === 'png') return 'image/png'
  if (ext === 'jpg' || ext === 'jpeg') return 'image/jpeg'
  if (ext === 'gif') return 'image/gif'
  if (ext === 'webp') return 'image/webp'
  if (ext === 'bmp') return 'image/bmp'
  if (ext === 'tif' || ext === 'tiff') return 'image/tiff'
  return null
}

function imageSize(buf) {
  if (!buf || buf.length < 10) return null
  if (buf[0] === 0x89 && buf[1] === 0x50 && buf.length >= 24) {
    return { width: buf.readUInt32BE(16), height: buf.readUInt32BE(20) }
  }
  if (buf.toString('ascii', 0, 3) === 'GIF') {
    return { width: buf.readUInt16LE(6), height: buf.readUInt16LE(8) }
  }
  if (buf[0] === 0xff && buf[1] === 0xd8) {
    let i = 2
    while (i < buf.length - 8) {
      if (buf[i] !== 0xff) {
        i += 1
        continue
      }
      const marker = buf[i + 1]
      if (marker === 0xc0 || marker === 0xc1 || marker === 0xc2) {
        return { height: buf.readUInt16BE(i + 5), width: buf.readUInt16BE(i + 7) }
      }
      const len = buf.readUInt16BE(i + 2)
      if (!len) break
      i += 2 + len
    }
  }
  return null
}

function resolveTarget(partPath, target) {
  if (!target) return ''
  if (/^[a-z]+:/i.test(target)) return target
  const cleaned = target.replace(/^\/+/, '')
  if (target.startsWith('/')) return cleaned
  const dir = partPath.split('/').slice(0, -1)
  for (const part of cleaned.split('/')) {
    if (part === '..') dir.pop()
    else if (part && part !== '.') dir.push(part)
  }
  return dir.join('/')
}

function relsPathFor(partPath) {
  const bits = partPath.split('/')
  const name = bits.pop()
  return [...bits, '_rels', `${name}.rels`].join('/')
}

async function zipText(zip, path) {
  const file = zip.file(path)
  if (!file) return ''
  return file.async('string')
}

function parseRels(xml) {
  const map = new Map()
  if (!xml) return map
  const $ = load(xml, { xmlMode: true })
  $('Relationship').each((_, el) => {
    map.set($(el).attr('Id'), {
      type: $(el).attr('Type') || '',
      target: $(el).attr('Target') || '',
      mode: $(el).attr('TargetMode') || ''
    })
  })
  return map
}

async function loadMedia(zip, rels, partPath) {
  const media = new Map()
  for (const [id, rel] of rels) {
    if (!String(rel.type).includes('/image')) continue
    const path = resolveTarget(partPath, rel.target)
    const file = zip.file(path)
    if (!file) continue
    const buf = await file.async('nodebuffer')
    const mime = mimeOf(path)
    const size = imageSize(buf) || { width: 0, height: 0 }
    media.set(id, {
      mime,
      base64: buf.toString('base64'),
      width: size.width,
      height: size.height,
      opaque: !mime
    })
  }
  return media
}

async function readStyles(zip) {
  const styles = new Map()
  const xml = await zipText(zip, 'word/styles.xml')
  if (!xml) return styles
  const $ = load(xml, { xmlMode: true })
  $('w\\:style').each((_, el) => {
    const id = $(el).attr('w:styleId')
    if (!id) return
    styles.set(id, { name: $(el).children('w\\:name').attr('w:val') || id })
  })
  return styles
}

async function readNumbering(zip) {
  const formats = new Map()
  const xml = await zipText(zip, 'word/numbering.xml')
  if (!xml) return formats
  const $ = load(xml, { xmlMode: true })
  const abstracts = new Map()
  $('w\\:abstractNum').each((_, el) => {
    const id = $(el).attr('w:abstractNumId')
    const levels = new Map()
    $(el).children('w\\:lvl').each((__, lvl) => {
      const ilvl = $(lvl).attr('w:ilvl') || '0'
      const fmt = $(lvl).children('w\\:numFmt').attr('w:val') || 'decimal'
      levels.set(ilvl, fmt === 'bullet' ? 'bullet' : 'decimal')
    })
    abstracts.set(id, levels)
  })
  $('w\\:num').each((_, el) => {
    const numId = $(el).attr('w:numId')
    const abs = $(el).children('w\\:abstractNumId').attr('w:val')
    if (numId) formats.set(numId, abstracts.get(abs) || new Map())
  })
  return formats
}

function headingTag(styleId, styleName) {
  const blob = `${styleId || ''} ${styleName || ''}`
  if (/\btitle\b/i.test(blob) && !/subtitle/i.test(blob)) return 'h1'
  const match = /heading\s*([1-6])/i.exec(blob)
  if (match) return `h${match[1]}`
  return null
}

function headingId(tag) {
  const match = /^h([1-6])$/i.exec(tag || '')
  return match ? `Heading${match[1]}` : null
}

function preserveSpaces(text) {
  return escHtml(text)
    .replace(/ {2,}/g, (spaces) => '&nbsp;'.repeat(spaces.length))
    .replace(/^ | $/g, '&nbsp;')
}

function runFormatting($, rPr) {
  const css = []
  const tags = []
  if (!rPr) return { css, tags }
  const node = $(rPr)
  const fonts = node.children('w\\:rFonts')
  const font = fonts.attr('w:ascii') || fonts.attr('w:hAnsi') || fonts.attr('w:cs')
  if (font) css.push(`font-family: '${String(font).replace(/'/g, '')}'`)
  const sz = node.children('w\\:sz').attr('w:val')
  if (sz) css.push(`font-size: ${Number(sz) / 2}pt`)
  const color = node.children('w\\:color').attr('w:val')
  if (color && color.toLowerCase() !== 'auto') css.push(`color: #${color}`)
  const highlight = node.children('w\\:highlight').attr('w:val')
  if (highlight && HIGHLIGHTS[highlight]) css.push(`background-color: #${HIGHLIGHTS[highlight]}`)
  else {
    const fill = node.children('w\\:shd').attr('w:fill')
    if (fill && fill.toLowerCase() !== 'auto') css.push(`background-color: #${fill}`)
  }
  if (onOff($, node.children('w\\:b'))) tags.push('strong')
  if (onOff($, node.children('w\\:i'))) tags.push('em')
  if (onOff($, node.children('w\\:u'))) tags.push('u')
  if (onOff($, node.children('w\\:strike')) || onOff($, node.children('w\\:dstrike'))) tags.push('s')
  const vert = node.children('w\\:vertAlign').attr('w:val')
  if (vert === 'superscript') tags.push('sup')
  if (vert === 'subscript') tags.push('sub')
  return { css, tags }
}

function wrapFormatted(html, formatting) {
  let out = html
  for (const tag of formatting.tags) out = `<${tag}>${out}</${tag}>`
  if (formatting.css.length) out = `<span style="${formatting.css.join('; ')}">${out}</span>`
  return out
}

function imageHtml($, node, media) {
  const blip = $(node).find('a\\:blip').first()
  let rid = blip.attr('r:embed') || blip.attr('r:link') || ''
  if (!rid) rid = $(node).find('v\\:imagedata').first().attr('r:id') || ''
  if (!rid || !media.has(rid)) return ''
  const item = media.get(rid)
  const extent = $(node).find('wp\\:extent').first()
  const cx = Number(extent.attr('cx') || 0)
  const cy = Number(extent.attr('cy') || 0)
  if (item.opaque) return `<span data-rid="${escAttr(rid)}" data-drawing="1"></span>`
  let width = cx ? Math.max(1, Math.round(cx / EMU_PER_PX)) : item.width || 320
  let height = cy ? Math.max(1, Math.round(cy / EMU_PER_PX)) : item.height || 240
  if (!cx && item.width && item.height && width > PAGE_CONTENT_PX) {
    height = Math.max(1, Math.round((item.height * PAGE_CONTENT_PX) / item.width))
    width = PAGE_CONTENT_PX
  }
  return `<img src="data:${item.mime};base64,${item.base64}" data-rid="${escAttr(rid)}" data-cx="${cx || ''}" data-cy="${cy || ''}" data-w="${width}" data-h="${height}" width="${width}" height="${height}" alt="">`
}

function renderRun($, run, media) {
  const formatting = runFormatting($, $(run).children('w\\:rPr').get(0))
  let html = ''
  $(run).children().each((_, child) => {
    const name = child.name
    if (name === 'w:rPr') return
    if (name === 'w:t') {
      const text = $(child).text()
      if (text) html += wrapFormatted(preserveSpaces(text), formatting)
      return
    }
    if (name === 'w:br' || name === 'w:cr') {
      html += '<br>'
      return
    }
    if (name === 'w:tab') {
      html += wrapFormatted('<span class="jd-tab">\t</span>', formatting)
      return
    }
    if (name === 'w:footnoteReference') {
      const id = $(child).attr('w:id') || ''
      html += `<sup data-footnote="${escAttr(id)}">${escHtml(id)}</sup>`
      return
    }
    if (name === 'w:endnoteReference') {
      const id = $(child).attr('w:id') || ''
      html += `<sup data-endnote="${escAttr(id)}">${escHtml(id)}</sup>`
      return
    }
    if (name === 'w:commentReference') {
      const id = $(child).attr('w:id') || ''
      html += `<span data-comment="${escAttr(id)}">&#8203;</span>`
      return
    }
    if (name === 'w:drawing' || name === 'w:pict' || name === 'mc:AlternateContent' || name === 'w:object') {
      html += imageHtml($, child, media)
    }
  })
  return html
}

function renderInlines($, node, ctx) {
  let html = ''
  $(node).children().each((_, child) => {
    const name = child.name
    if (name === 'w:r') html += renderRun($, child, ctx.media)
    else if (name === 'w:hyperlink') html += renderHyperlink($, child, ctx)
    else if (name === 'w:ins' || name === 'w:smartTag') html += renderInlines($, child, ctx)
    else if (name === 'w:del' || name === 'w:bookmarkStart' || name === 'w:bookmarkEnd') return
    else if (name === 'w:sdt') {
      const content = $(child).children('w\\:sdtContent').get(0)
      if (content) html += renderInlines($, content, ctx)
    } else if (name === 'w:fldSimple' || name === 'mc:AlternateContent') {
      const image = imageHtml($, child, ctx.media)
      html += image || renderInlines($, child, ctx)
    }
  })
  return html
}

function renderHyperlink($, node, ctx) {
  const rid = $(node).attr('r:id')
  const anchor = $(node).attr('w:anchor')
  const rel = rid ? ctx.rels.get(rid) : null
  const href = rel?.target || (anchor ? `#${anchor}` : '')
  const inner = renderInlines($, node, ctx)
  if (!href) return inner
  return `<a href="${escAttr(href)}">${inner}</a>`
}

function paragraphProps($, node, ctx) {
  const pPr = $(node).children('w\\:pPr')
  const styleId = pPr.children('w\\:pStyle').attr('w:val') || ''
  const styleName = ctx.styles.get(styleId)?.name || ''
  const tag = headingTag(styleId, styleName) || 'p'
  const css = []
  const align = pPr.children('w\\:jc').attr('w:val')
  if (align === 'center') css.push('text-align: center')
  else if (align === 'right' || align === 'end') css.push('text-align: right')
  else if (align === 'both' || align === 'distribute') css.push('text-align: justify')
  const ind = pPr.children('w\\:ind')
  if (ind.attr('w:left')) css.push(`margin-left: ${twipsToPt(ind.attr('w:left'))}pt`)
  if (ind.attr('w:right')) css.push(`margin-right: ${twipsToPt(ind.attr('w:right'))}pt`)
  if (ind.attr('w:hanging')) css.push(`text-indent: -${twipsToPt(ind.attr('w:hanging'))}pt`)
  else if (ind.attr('w:firstLine')) css.push(`text-indent: ${twipsToPt(ind.attr('w:firstLine'))}pt`)
  const spacing = pPr.children('w\\:spacing')
  if (spacing.attr('w:before')) css.push(`margin-top: ${twipsToPt(spacing.attr('w:before'))}pt`)
  if (spacing.attr('w:after')) css.push(`margin-bottom: ${twipsToPt(spacing.attr('w:after'))}pt`)
  const fill = pPr.children('w\\:shd').attr('w:fill')
  if (fill && fill.toLowerCase() !== 'auto') css.push(`background-color: #${fill}`)
  const num = pPr.children('w\\:numPr')
  const numId = num.children('w\\:numId').attr('w:val') || ''
  const ilvl = num.children('w\\:ilvl').attr('w:val') || '0'
  const listed = Boolean(numId && numId !== '0')
  const levels = listed ? ctx.numbering.get(numId) : null
  const fmt = levels?.get(String(ilvl)) || levels?.get('0') || 'decimal'
  return {
    tag: listed ? 'p' : tag,
    pStyle: styleId,
    css,
    listed,
    numId,
    ilvl: Number(ilvl) || 0,
    fmt
  }
}

function paragraphInner($, node, ctx) {
  return renderInlines($, node, ctx)
}

function itemsToList(items) {
  function build(index, level) {
    const tag = items[index].fmt === 'bullet' ? 'ul' : 'ol'
    let html = `<${tag} data-num-id="${escAttr(items[index].numId)}" data-fmt="${items[index].fmt}">`
    while (index < items.length) {
      if (items[index].ilvl < level) break
      if (items[index].ilvl > level) {
        const nested = build(index, items[index].ilvl)
        html = html.replace(/<\/li>$/, `${nested.html}</li>`)
        index = nested.index
        continue
      }
      const style = items[index].css.length ? ` style="${items[index].css.join('; ')}"` : ''
      const dataStyle = items[index].pStyle ? ` data-style="${escAttr(items[index].pStyle)}"` : ''
      html += `<li${style}${dataStyle}>${items[index].inner || '<br>'}`
      index += 1
      if (index < items.length && items[index].ilvl > level) {
        const nested = build(index, items[index].ilvl)
        html += nested.html
        index = nested.index
      }
      html += '</li>'
    }
    html += `</${tag}>`
    return { html, index }
  }
  return build(0, items[0].ilvl).html
}

function tableHtml($, table, ctx) {
  const rows = $(table).children('w\\:tr').toArray()
  if (!rows.length) {
    return $(table).children('w\\:tblGrid, w\\:tblPr').length
      ? ''
      : ''
  }
  let body = ''
  for (const tr of rows) {
    let cells = ''
    for (const tc of $(tr).children('w\\:tc').toArray()) {
      const tcPr = $(tc).children('w\\:tcPr')
      const span = Number(tcPr.children('w\\:gridSpan').attr('w:val') || 1)
      const colspan = span > 1 ? ` colspan="${span}"` : ''
      const fill = tcPr.children('w\\:shd').attr('w:fill')
      const style = fill && fill.toLowerCase() !== 'auto' ? ` style="background-color:#${fill}"` : ''
      const inner = renderChildren($, tc, ctx).trim() || '<p></p>'
      cells += `<td${colspan}${style}>${inner}</td>`
    }
    body += `<tr>${cells}</tr>`
  }
  return `<table><tbody>${body}</tbody></table>`
}

function renderChildren($, parent, ctx) {
  const nodes = $(parent).children().toArray().filter((node) => node.type === 'tag')
  let html = ''
  for (let i = 0; i < nodes.length; i += 1) {
    const name = nodes[i].name
    if (name === 'w:sectPr' || name === 'w:tblPr' || name === 'w:tblGrid' || name === 'w:tcPr' || name === 'w:trPr') continue
    if (name === 'w:sdt') {
      const content = $(nodes[i]).children('w\\:sdtContent').get(0)
      if (content) html += renderChildren($, content, ctx)
      continue
    }
    if (name === 'w:p') {
      const props = paragraphProps($, nodes[i], ctx)
      if (props.listed) {
        const items = []
        while (i < nodes.length && nodes[i].name === 'w:p') {
          const next = paragraphProps($, nodes[i], ctx)
          if (!next.listed) break
          items.push({ ...next, inner: paragraphInner($, nodes[i], ctx) })
          i += 1
        }
        i -= 1
        if (items.length) html += itemsToList(items)
        continue
      }
      const inner = paragraphInner($, nodes[i], ctx)
      const style = props.css.length ? ` style="${props.css.join('; ')}"` : ''
      const dataStyle = props.pStyle ? ` data-style="${escAttr(props.pStyle)}"` : ''
      html += `<${props.tag}${style}${dataStyle}>${inner || '<br>'}</${props.tag}>`
      continue
    }
    if (name === 'w:tbl') {
      html += tableHtml($, nodes[i], ctx)
      continue
    }
    if ($(nodes[i]).find('w\\:p, w\\:tbl').length) html += renderChildren($, nodes[i], ctx)
  }
  return html
}

function readPage($) {
  const sect = $('w\\:body > w\\:sectPr').last()
  if (!sect.length) return { ...DEFAULT_PAGE }
  const size = sect.children('w\\:pgSz')
  const margin = sect.children('w\\:pgMar')
  const width = Number(size.attr('w:w')) || 12240
  const height = Number(size.attr('w:h')) || 15840
  const orient = size.attr('w:orient')
  const landscape = orient === 'landscape' && width < height
  return {
    width: inches(landscape ? height : width),
    height: inches(landscape ? width : height),
    marginTop: inches(margin.attr('w:top') || 1440),
    marginRight: inches(margin.attr('w:right') || 1440),
    marginBottom: inches(margin.attr('w:bottom') || 1440),
    marginLeft: inches(margin.attr('w:left') || 1440)
  }
}

async function readBand(zip, rels, rid) {
  const rel = rels.get(rid)
  if (!rel?.target) return ''
  const partPath = resolveTarget('word/document.xml', rel.target)
  const xml = await zipText(zip, partPath)
  if (!xml) return ''
  const partRels = parseRels(await zipText(zip, relsPathFor(partPath)))
  const media = await loadMedia(zip, partRels, partPath)
  const $ = load(xml, { xmlMode: true })
  const root = $('w\\:hdr, w\\:ftr').get(0)
  if (!root) return ''
  const ctx = { styles: new Map(), numbering: new Map(), rels: partRels, media }
  return renderChildren($, root, ctx).trim()
}

async function readBands(zip, $, rels) {
  const sect = $('w\\:body > w\\:sectPr').last()
  const pick = (tag) => {
    const all = sect.children(tag).toArray()
    const preferred = all.find((el) => $(el).attr('w:type') === 'default') || all[0]
    return preferred ? $(preferred).attr('r:id') : ''
  }
  const headerId = pick('w\\:headerReference')
  const footerId = pick('w\\:footerReference')
  return {
    header: headerId ? await readBand(zip, rels, headerId) : '',
    footer: footerId ? await readBand(zip, rels, footerId) : ''
  }
}

async function readDocx(buffer) {
  let zip
  try {
    zip = await JSZip.loadAsync(buffer)
  } catch {
    throw new Error('This file is not a Word .docx document. It may be password-protected or an older .doc file.')
  }
  const documentXml = await zipText(zip, 'word/document.xml')
  if (!documentXml) {
    throw new Error('This file is not a Word .docx document.')
  }
  const $ = load(documentXml, { xmlMode: true })
  const rels = parseRels(await zipText(zip, 'word/_rels/document.xml.rels'))
  const ctx = {
    styles: await readStyles(zip),
    numbering: await readNumbering(zip),
    rels,
    media: await loadMedia(zip, rels, 'word/document.xml')
  }
  const body = $('w\\:body').get(0)
  const html = (body ? renderChildren($, body, ctx) : '').trim() || '<p></p>'
  const bands = await readBands(zip, $, rels)
  return {
    html,
    page: readPage($),
    headerHtml: bands.header,
    footerHtml: bands.footer
  }
}

function parseStyle(style, next, block = false) {
  const font = /font-family:\s*([^;]+)/i.exec(style)
  if (font) next.font = font[1].replace(/['"]/g, '').split(',')[0].trim()
  const size = /font-size:\s*([\d.]+)\s*(px|pt)/i.exec(style)
  if (size) {
    const n = parseFloat(size[1])
    next.size = size[2].toLowerCase() === 'pt' ? Math.round(n * 2) : Math.round(((n * 72) / 96) * 2)
  }
  const color = /(?:^|;)\s*color:\s*([^;]+)/i.exec(style)
  if (color) next.color = cssToHex(color[1])
  if (!block) {
    const bg = /background(?:-color)?:\s*([^;]+)/i.exec(style)
    if (bg) next.highlight = cssToHex(bg[1])
  }
  const weight = /font-weight:\s*([^;]+)/i.exec(style)
  if (weight) next.bold = /bold|[6-9]00/.test(weight[1])
  const italic = /font-style:\s*([^;]+)/i.exec(style)
  if (italic) next.italics = /italic|oblique/.test(italic[1])
  const deco = /text-decoration(?:-line)?:\s*([^;]+)/i.exec(style)
  if (deco) {
    if (/underline/.test(deco[1])) next.underline = true
    if (/line-through/.test(deco[1])) next.strike = true
    if (/none/.test(deco[1])) {
      next.underline = false
      next.strike = false
    }
  }
}

function runOpts($, el, inherited = {}) {
  const tag = (el.name || '').toLowerCase()
  const next = { ...inherited }
  if (tag === 'strong' || tag === 'b') next.bold = true
  if (tag === 'em' || tag === 'i') next.italics = true
  if (tag === 'u') next.underline = true
  if (tag === 's' || tag === 'strike' || tag === 'del') next.strike = true
  if (tag === 'sup') next.superScript = true
  if (tag === 'sub') next.subScript = true
  if (tag === 'font') {
    const face = $(el).attr('face')
    const color = $(el).attr('color')
    if (face) next.font = face.split(',')[0].trim()
    if (color) next.color = cssToHex(color)
  }
  const style = $(el).attr('style') || ''
  if (style) parseStyle(style, next, false)
  return next
}

function parseDataImage(src) {
  const match = /^data:(image\/[a-z0-9.+-]+);base64,([a-z0-9+/=\s]+)$/i.exec(src || '')
  if (!match) return null
  let mime = match[1].toLowerCase()
  const base64 = match[2].replace(/\s+/g, '')
  const buffer = Buffer.from(base64, 'base64')
  if (!buffer.length) return null
  let ext = 'png'
  if (mime.includes('jpeg') || mime.includes('jpg')) {
    mime = 'image/jpeg'
    ext = 'jpg'
  } else if (mime.includes('gif')) ext = 'gif'
  else if (mime.includes('webp')) ext = 'webp'
  else if (mime.includes('bmp')) ext = 'bmp'
  else if (mime.includes('tiff')) ext = 'tiff'
  else if (!mime.includes('png')) return null
  const size = imageSize(buffer) || { width: 0, height: 0 }
  return { mime, ext, base64, width: size.width, height: size.height }
}

function collectRuns($, node, inherited = {}, href = null, writer = null) {
  const runs = []
  for (const child of $(node).contents().toArray()) {
    if (child.type === 'text') {
      // The browser collapses whitespace runs to one space, so Word should see the same.
      const text = child.data.replace(/\u00a0/g, ' ').replace(/[ \t\r\n]+/g, ' ')
      if (text) runs.push({ text, opts: inherited, href })
      continue
    }
    if (child.type !== 'tag') continue
    const tag = child.name.toLowerCase()
    if (tag === 'br') {
      runs.push({ break: true, opts: inherited, href })
      continue
    }
    if ((child.attribs?.class || '').split(/\s+/).includes('jd-tab')) {
      runs.push({ text: '\t', opts: inherited, href })
      continue
    }
    if (tag === 'img') {
      runs.push({ image: child, href: null })
      continue
    }
    if (child.attribs?.['data-footnote']) {
      runs.push({ footnote: child.attribs['data-footnote'] })
      continue
    }
    if (child.attribs?.['data-endnote']) {
      runs.push({ endnote: child.attribs['data-endnote'] })
      continue
    }
    if (child.attribs?.['data-comment']) {
      runs.push({ comment: child.attribs['data-comment'] })
      continue
    }
    if (child.attribs?.['data-drawing'] && child.attribs?.['data-rid']) {
      runs.push({ keepRid: child.attribs['data-rid'] })
      continue
    }
    if (tag === 'a') {
      const link = ($(child).attr('href') || '').trim() || href
      runs.push(...collectRuns($, child, inherited, link, writer))
      continue
    }
    runs.push(...collectRuns($, child, runOpts($, child, inherited), href, writer))
  }
  return runs
}

function rPrXml(opts = {}, link = false) {
  const parts = []
  if (link && !opts.color) parts.push('<w:rStyle w:val="Hyperlink"/>')
  if (opts.font) {
    const font = escAttr(opts.font)
    parts.push(`<w:rFonts w:ascii="${font}" w:hAnsi="${font}" w:cs="${font}"/>`)
  }
  if (opts.bold) parts.push('<w:b/>')
  if (opts.italics) parts.push('<w:i/>')
  if (opts.underline) parts.push('<w:u w:val="single"/>')
  if (opts.strike) parts.push('<w:strike/>')
  if (opts.superScript) parts.push('<w:vertAlign w:val="superscript"/>')
  if (opts.subScript) parts.push('<w:vertAlign w:val="subscript"/>')
  if (opts.size) {
    parts.push(`<w:sz w:val="${opts.size}"/>`)
    parts.push(`<w:szCs w:val="${opts.size}"/>`)
  }
  if (opts.color) parts.push(`<w:color w:val="${opts.color}"/>`)
  const named = highlightName(opts.highlight)
  if (named) parts.push(`<w:highlight w:val="${named}"/>`)
  if (opts.highlight) parts.push(`<w:shd w:val="clear" w:color="auto" w:fill="${opts.highlight}"/>`)
  return parts.length ? `<w:rPr>${parts.join('')}</w:rPr>` : ''
}

function textRunXml(run, link) {
  const props = rPrXml(run.opts, link)
  if (run.break) return `<w:r>${props}<w:br/></w:r>`
  const text = run.text || ''
  let inner = ''
  text.split('\t').forEach((part, index) => {
    if (index) inner += '<w:tab/>'
    if (!part) return
    const space = /^\s|\s$/.test(part) ? ' xml:space="preserve"' : ''
    inner += `<w:t${space}>${escXml(part)}</w:t>`
  })
  if (!inner) inner = '<w:t/>'
  return `<w:r>${props}${inner}</w:r>`
}

function drawingXml(rid, cx, cy, docPr) {
  return `<w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${cx}" cy="${cy}"/><wp:effectExtent t="0" r="0" b="0" l="0"/><wp:docPr id="${docPr}" name="Picture ${docPr}"/><wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/></wp:cNvGraphicFramePr><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="${docPr}" name="Picture ${docPr}" descr=""/><pic:cNvPicPr><a:picLocks noChangeAspect="1" noChangeArrowheads="1"/></pic:cNvPicPr></pic:nvPicPr><pic:blipFill><a:blip r:embed="${rid}" cstate="none"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr bwMode="auto"><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>`
}

function imageXml($, el, writer) {
  const keep = $(el).attr('data-drawing')
  const ridAttr = $(el).attr('data-rid') || ''
  if (keep && ridAttr) {
    const queued = writer.drawings.get(ridAttr)?.shift()
    if (queued) return `<w:r>${queued}</w:r>`
  }
  const parsed = parseDataImage($(el).attr('src') || '')
  if (!parsed) return ''
  let width = Number($(el).attr('width')) || parsed.width || 320
  let height = Number($(el).attr('height')) || parsed.height || 240
  if (!$(el).attr('width') && parsed.width && parsed.height && width > PAGE_CONTENT_PX) {
    height = Math.max(1, Math.round((parsed.height * PAGE_CONTENT_PX) / parsed.width))
    width = PAGE_CONTENT_PX
  }
  const cx = Math.max(1, Math.round(width * EMU_PER_PX))
  const cy = Math.max(1, Math.round(height * EMU_PER_PX))
  const sameBytes = ridAttr && writer.mediaB64.get(ridAttr) === parsed.base64
  const sameSize = String($(el).attr('data-w') || '') === String(width) && String($(el).attr('data-h') || '') === String(height)
  if (sameBytes) {
    const queued = sameSize ? writer.drawings.get(ridAttr)?.shift() : null
    if (queued) return `<w:r>${queued}</w:r>`
    return `<w:r>${drawingXml(ridAttr, cx, cy, writer.nextDocPr())}</w:r>`
  }
  const rid = writer.addImage(parsed)
  return `<w:r>${drawingXml(rid, cx, cy, writer.nextDocPr())}</w:r>`
}

function runsToXml(runs, $, writer) {
  let xml = ''
  for (let i = 0; i < runs.length; i += 1) {
    const run = runs[i]
    if (run.image) {
      xml += imageXml($, run.image, writer)
      continue
    }
    if (run.keepRid) {
      const queued = writer.drawings.get(run.keepRid)?.shift()
      if (queued) xml += `<w:r>${queued}</w:r>`
      continue
    }
    if (run.footnote) {
      xml += `<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="${escAttr(run.footnote)}"/></w:r>`
      continue
    }
    if (run.endnote) {
      xml += `<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:endnoteReference w:id="${escAttr(run.endnote)}"/></w:r>`
      continue
    }
    if (run.comment) {
      xml += `<w:r><w:commentReference w:id="${escAttr(run.comment)}"/></w:r>`
      continue
    }
    const href = run.href || ''
    const external = /^(https?:|mailto:)/i.test(href)
    const anchor = href.startsWith('#') ? href.slice(1) : ''
    if (external || anchor) {
      let inner = ''
      while (i < runs.length && (runs[i].href || '') === href && !runs[i].image) {
        inner += textRunXml(runs[i], external)
        i += 1
      }
      i -= 1
      if (external) xml += `<w:hyperlink r:id="${writer.hyperlink(href)}" w:history="1">${inner}</w:hyperlink>`
      else xml += `<w:hyperlink w:anchor="${escAttr(anchor)}" w:history="1">${inner}</w:hyperlink>`
      continue
    }
    xml += textRunXml(run, false)
  }
  return xml
}

function alignmentOf($, el) {
  const style = ($(el).attr('style') || '').toLowerCase()
  if (/text-align:\s*center/.test(style)) return 'center'
  if (/text-align:\s*right/.test(style)) return 'right'
  if (/text-align:\s*justify/.test(style)) return 'both'
  const align = ($(el).attr('align') || '').toLowerCase()
  if (align === 'center' || align === 'right') return align
  if (align === 'justify') return 'both'
  return ''
}

function pPrXml($, el, extra = {}) {
  const tag = (el.name || '').toLowerCase()
  const style = $(el).attr('style') || ''
  const pStyle = $(el).attr('data-style') || headingId(tag) || ''
  const parts = []
  if (pStyle) parts.push(`<w:pStyle w:val="${escAttr(pStyle)}"/>`)
  if (extra.numId) {
    parts.push(`<w:numPr><w:ilvl w:val="${extra.ilvl || 0}"/><w:numId w:val="${escAttr(extra.numId)}"/></w:numPr>`)
  }
  const align = alignmentOf($, el)
  if (align) parts.push(`<w:jc w:val="${align}"/>`)
  const spacing = []
  const before = /margin-top:\s*([^;]+)/i.exec(style)
  const after = /margin-bottom:\s*([^;]+)/i.exec(style)
  if (before) {
    const twips = toTwips(before[1])
    if (twips != null) spacing.push(`w:before="${Math.max(0, twips)}"`)
  }
  if (after) {
    const twips = toTwips(after[1])
    if (twips != null) spacing.push(`w:after="${Math.max(0, twips)}"`)
  }
  if (spacing.length) parts.push(`<w:spacing ${spacing.join(' ')}/>`)
  const ind = []
  let left = 0
  const marginLeft = /margin-left:\s*([^;]+)/i.exec(style)
  const marginRight = /margin-right:\s*([^;]+)/i.exec(style)
  const textIndent = /text-indent:\s*([^;]+)/i.exec(style)
  if (marginLeft) left = toTwips(marginLeft[1]) || 0
  if (tag === 'blockquote') left = Math.max(left, 720)
  if (textIndent) {
    const indent = toTwips(textIndent[1]) || 0
    if (indent < 0) ind.push(`w:hanging="${Math.abs(indent)}"`)
    else if (indent > 0) ind.push(`w:firstLine="${indent}"`)
  }
  if (left) ind.push(`w:left="${left}"`)
  if (marginRight) {
    const right = toTwips(marginRight[1])
    if (right) ind.push(`w:right="${right}"`)
  }
  if (ind.length && !extra.numId) parts.push(`<w:ind ${ind.join(' ')}/>`)
  const bg = /background(?:-color)?:\s*([^;]+)/i.exec(style)
  const fill = bg ? cssToHex(bg[1]) : ''
  if (fill) parts.push(`<w:shd w:val="clear" w:color="auto" w:fill="${fill}"/>`)
  return parts.length ? `<w:pPr>${parts.join('')}</w:pPr>` : ''
}

function paragraphXml($, el, extra, writer) {
  const tag = (el.name || '').toLowerCase()
  const inherited = {}
  parseStyle($(el).attr('style') || '', inherited, true)
  if (tag === 'blockquote') inherited.italics = inherited.italics || true
  if (tag === 'pre') {
    inherited.font = inherited.font || 'Courier New'
    inherited.size = inherited.size || 20
  }
  const runs = collectRuns($, el, inherited, null, writer)
  return `<w:p>${pPrXml($, el, extra)}${runsToXml(runs, $, writer)}</w:p>`
}

function hasBlockChild($, node) {
  return $(node).children().toArray().some((child) => BLOCK_TAGS.has((child.name || '').toLowerCase()))
}

function listXml($, node, ilvl, writer) {
  const tag = (node.name || '').toLowerCase()
  const wanted = tag === 'ol' ? 'decimal' : 'bullet'
  const marked = $(node).attr('data-fmt')
  const reuse = $(node).attr('data-num-id')
  const numId = reuse && (!marked || marked === wanted) ? reuse : writer.listNum(wanted)
  let xml = ''
  $(node).children('li').each((_, li) => {
    const nested = $(li).children('ul, ol').toArray()
    const clone = $(li).clone()
    clone.children('ul, ol').remove()
    const blocks = hasBlockChild($, clone.get(0))
      ? $(clone).children().toArray().filter((child) => BLOCK_TAGS.has((child.name || '').toLowerCase()))
      : []
    if (!blocks.length) {
      xml += paragraphXml($, clone.get(0), { numId, ilvl }, writer)
    } else {
      blocks.forEach((block, index) => {
        xml += paragraphXml($, block, index === 0 ? { numId, ilvl } : {}, writer)
      })
    }
    for (const child of nested) xml += listXml($, child, ilvl + 1, writer)
  })
  return xml
}

function columnCount($, table) {
  let max = 1
  $(table).children('tr, tbody, thead, tfoot').each((_, section) => {
    const rows = section.name === 'tr' ? [section] : $(section).children('tr').toArray()
    for (const tr of rows) {
      let count = 0
      $(tr).children('td, th').each((__, cell) => {
        count += Number($(cell).attr('colspan') || 1)
      })
      max = Math.max(max, count)
    }
  })
  return max
}

function tableXml($, table, writer) {
  const cols = columnCount($, table)
  const grid = Array.from({ length: cols }, () => `<w:gridCol w:w="${Math.round(9360 / cols)}"/>`).join('')
  const borders = ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']
    .map((edge) => `<w:${edge} w:val="single" w:sz="4" w:space="0" w:color="auto"/>`)
    .join('')
  let rows = ''
  const eachRow = (tr) => {
    let cells = ''
    let seen = 0
    $(tr).children('td, th').each((_, cell) => {
      const span = Number($(cell).attr('colspan') || 1)
      seen += span
      const width = Math.max(1, Math.round((5000 * span) / cols))
      const bg = cssToHex((($(cell).attr('style') || '').match(/background(?:-color)?:\s*([^;]+)/i) || [])[1])
      const tcPr = [
        `<w:tcW w:w="${width}" w:type="pct"/>`,
        span > 1 ? `<w:gridSpan w:val="${span}"/>` : '',
        bg ? `<w:shd w:val="clear" w:color="auto" w:fill="${bg}"/>` : ''
      ].join('')
      let inner = blocksFrom($, cell, writer)
      if (!inner) inner = '<w:p/>'
      cells += `<w:tc><w:tcPr>${tcPr}</w:tcPr>${inner}</w:tc>`
    })
    if (seen < cols) {
      const width = Math.max(1, Math.round((5000 * (cols - seen)) / cols))
      cells += `<w:tc><w:tcPr><w:tcW w:w="${width}" w:type="pct"/></w:tcPr><w:p/></w:tc>`
    }
    rows += `<w:tr>${cells}</w:tr>`
  }
  $(table).children().each((_, child) => {
    if (child.name === 'tr') eachRow(child)
    if (child.name === 'tbody' || child.name === 'thead' || child.name === 'tfoot') {
      $(child).children('tr').each((__, tr) => eachRow(tr))
    }
  })
  return `<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="pct"/><w:tblBorders>${borders}</w:tblBorders></w:tblPr><w:tblGrid>${grid}</w:tblGrid>${rows}</w:tbl>`
}

function emitBlock($, child, writer) {
  const tag = (child.name || '').toLowerCase()
  if (tag === 'ul' || tag === 'ol') return listXml($, child, 0, writer)
  if (tag === 'table') return tableXml($, child, writer)
  if (tag === 'br') return '<w:p/>'
  if (['p', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'blockquote', 'pre', 'li'].includes(tag)) {
    return paragraphXml($, child, {}, writer)
  }
  if (tag === 'img') return `<w:p>${imageXml($, child, writer)}</w:p>`
  if (hasBlockChild($, child)) return blocksFrom($, child, writer)
  if (tag === 'div' || tag === 'section' || tag === 'article') return paragraphXml($, child, {}, writer)
  return ''
}

function blocksFrom($, node, writer) {
  let xml = ''
  let loose = []
  const flush = () => {
    if (!loose.length) return
    const wrapper = $('<p></p>')
    for (const piece of loose) wrapper.append(piece)
    xml += paragraphXml($, wrapper.get(0), {}, writer)
    loose = []
  }
  for (const child of $(node).contents().toArray()) {
    const block = child.type === 'tag' && BLOCK_TAGS.has((child.name || '').toLowerCase())
    if (!block) {
      if (child.type === 'text' && !String(child.data || '').replace(/\u00a0/g, ' ').trim() && !loose.length) continue
      if (child.type === 'tag' || child.type === 'text') loose.push(child)
      continue
    }
    flush()
    xml += emitBlock($, child, writer)
  }
  flush()
  return xml
}

function indexDrawings(xml) {
  const drawings = new Map()
  if (!xml) return drawings
  const $ = load(xml, { xmlMode: true })
  $('a\\:blip, v\\:imagedata').each((_, el) => {
    const rid = $(el).attr('r:embed') || $(el).attr('r:id') || ''
    if (!rid) return
    let host = $(el).closest('mc\\:AlternateContent')
    if (!host.length) host = $(el).closest('w\\:drawing')
    if (!host.length) host = $(el).closest('w\\:pict')
    if (!host.length) return
    const snippet = $.html(host)
    const list = drawings.get(rid) || []
    if (!list.includes(snippet)) list.push(snippet)
    drawings.set(rid, list)
  })
  return drawings
}

class PackageWriter {
  constructor(zip, relsPath, relsXml, drawings, mediaB64) {
    this.zip = zip
    this.relsPath = relsPath
    this.relsXml = relsXml || ''
    this.drawings = drawings
    this.mediaB64 = mediaB64
    this.docPrId = 3000
    this.imageCount = Object.keys(zip.files).filter((name) => name.startsWith('word/media/')).length
    this.byHash = new Map()
    this.linkIds = new Map()
    this.createdLists = {}
    this.numberingDirty = false
  }

  nextDocPr() {
    this.docPrId += 1
    return this.docPrId
  }

  nextRid() {
    let max = 0
    const re = /Id="rId(\d+)"/g
    let match = re.exec(this.relsXml)
    while (match) {
      max = Math.max(max, Number(match[1]))
      match = re.exec(this.relsXml)
    }
    const id = `rId${max + 1}`
    return id
  }

  addRel(id, type, target, external = false) {
    const mode = external ? ' TargetMode="External"' : ''
    const rel = `<Relationship Id="${id}" Type="${type}" Target="${escAttr(target)}"${mode}/>`
    if (this.relsXml.includes('</Relationships>')) {
      this.relsXml = this.relsXml.replace('</Relationships>', `${rel}</Relationships>`)
    } else {
      this.relsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rel}</Relationships>`
    }
  }

  ensureContentType(ext, mime) {
    const path = '[Content_Types].xml'
    const xml = this.contentTypes || ''
    if (!xml || new RegExp(`Extension="${ext}"`, 'i').test(xml)) return
    this.contentTypes = xml.replace('</Types>', `<Default Extension="${ext}" ContentType="${mime}"/></Types>`)
    this.zip.file(path, this.contentTypes)
  }

  addImage(parsed) {
    const cached = this.byHash.get(parsed.base64)
    if (cached) return cached
    this.imageCount += 1
    const rid = this.nextRid()
    const target = `media/jarvisdoc-${this.imageCount}.${parsed.ext}`
    this.zip.file(`word/${target}`, Buffer.from(parsed.base64, 'base64'))
    this.addRel(rid, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image', target)
    this.ensureContentType(parsed.ext, parsed.mime)
    this.byHash.set(parsed.base64, rid)
    return rid
  }

  hyperlink(url) {
    const existing = this.linkIds.get(url)
    if (existing) return existing
    const match = this.relsXml.match(new RegExp(`Target="${escAttr(url).replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}"[^>]*Id="([^"]+)"|Id="([^"]+)"[^>]*Target="${escAttr(url).replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}"`))
    if (match) {
      const id = match[1] || match[2]
      this.linkIds.set(url, id)
      return id
    }
    const rid = this.nextRid()
    this.addRel(rid, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink', url, true)
    this.linkIds.set(url, rid)
    return rid
  }

  listNum(kind) {
    if (this.createdLists[kind]) return this.createdLists[kind]
    const xml = this.numberingXml || ''
    const ids = [...xml.matchAll(/w:numId="(\d+)"/g)].map((match) => Number(match[1]))
    const abs = [...xml.matchAll(/w:abstractNumId="(\d+)"/g)].map((match) => Number(match[1]))
    const numId = String(Math.max(10, ...ids, 0) + Object.keys(this.createdLists).length + 1)
    const absId = String(Math.max(0, ...abs) + Object.keys(this.createdLists).length + 1)
    const levels = []
    for (let i = 0; i < 9; i += 1) {
      const bullet = kind === 'bullet'
      const text = bullet ? (['•', '◦', '▪'][i % 3]) : `%${i + 1}.`
      const left = 720 * (i + 1)
      levels.push(`<w:lvl w:ilvl="${i}"><w:start w:val="1"/><w:numFmt w:val="${bullet ? 'bullet' : 'decimal'}"/><w:lvlText w:val="${text}"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="${left}" w:hanging="360"/></w:pPr></w:lvl>`)
    }
    const abstract = `<w:abstractNum w:abstractNumId="${absId}"><w:multiLevelType w:val="hybridMultilevel"/>${levels.join('')}</w:abstractNum>`
    const num = `<w:num w:numId="${numId}"><w:abstractNumId w:val="${absId}"/></w:num>`
    if (!this.numberingXml) {
      this.numberingXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${abstract}${num}</w:numbering>`
    } else if (this.numberingXml.includes('<w:num ')) {
      this.numberingXml = this.numberingXml
        .replace('<w:num ', `${abstract}<w:num `)
        .replace('</w:numbering>', `${num}</w:numbering>`)
    } else {
      this.numberingXml = this.numberingXml.replace('</w:numbering>', `${abstract}${num}</w:numbering>`)
    }
    this.numberingDirty = true
    this.createdLists[kind] = numId
    this.ensureNumberingPart()
    return numId
  }

  ensureNumberingPart() {
    const type = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering'
    if (!this.relsXml.includes('/numbering')) {
      const rid = this.nextRid()
      this.addRel(rid, type, 'numbering.xml')
    }
    const types = this.contentTypes || ''
    if (types && !types.includes('/word/numbering.xml')) {
      this.contentTypes = types.replace(
        '</Types>',
        '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/></Types>'
      )
      this.zip.file('[Content_Types].xml', this.contentTypes)
    }
  }

  commit() {
    if (this.relsPath) this.zip.file(this.relsPath, this.relsXml)
    if (this.numberingDirty) this.zip.file('word/numbering.xml', this.numberingXml)
  }
}

function htmlToBodyXml(html, writer) {
  const $ = load(`<div id="jarvisdoc-root">${html || ''}</div>`)
  $('#jarvisdoc-root [data-role="header"], #jarvisdoc-root [data-role="footer"]').remove()
  let xml = blocksFrom($, $('#jarvisdoc-root').get(0), writer)
  if (!xml) xml = '<w:p/>'
  return xml
}

let blankPromise = null

function blankDocx() {
  if (!blankPromise) {
    const doc = new Document({
      creator: 'JarvisDoc',
      title: 'Untitled',
      styles: {
        default: {
          document: {
            run: { font: 'Calibri', size: 22 }
          }
        }
      },
      sections: [
        {
          properties: {
            page: {
              size: { width: 12240, height: 15840 },
              margin: { top: 1440, right: 1440, bottom: 1440, left: 1440, header: 708, footer: 708 }
            }
          },
          children: [new Paragraph({})]
        }
      ]
    })
    blankPromise = Packer.toBuffer(doc)
  }
  return blankPromise
}

async function writeDocx(html, options = {}) {
  const source = options.sourceBuffer?.length ? options.sourceBuffer : await blankDocx()
  let zip
  try {
    zip = await JSZip.loadAsync(source)
  } catch {
    throw new Error('Could not save this Word document.')
  }
  const part = 'word/document.xml'
  const xml = await zipText(zip, part)
  if (!xml) throw new Error('Could not save this Word document.')
  const relPath = 'word/_rels/document.xml.rels'
  const relsXml = await zipText(zip, relPath)
  const rels = parseRels(relsXml)
  const media = await loadMedia(zip, rels, part)
  const mediaB64 = new Map()
  for (const [id, item] of media) {
    if (item.base64) mediaB64.set(id, item.base64)
  }
  const writer = new PackageWriter(zip, relPath, relsXml, indexDrawings(xml), mediaB64)
  writer.contentTypes = await zipText(zip, '[Content_Types].xml')
  writer.numberingXml = await zipText(zip, 'word/numbering.xml')
  const inner = htmlToBodyXml(html, writer)
  const $ = load(xml, { xmlMode: true })
  const body = $('w\\:body')
  const sect = body.children('w\\:sectPr')
  const sectXml = sect.toArray().map((node) => $.html(node)).join('')
  body.empty()
  body.append(inner)
  if (sectXml) body.append(sectXml)
  zip.file(part, $.xml())
  writer.commit()
  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' })
}

module.exports = {
  DEFAULT_PAGE,
  readDocx,
  writeDocx
}
