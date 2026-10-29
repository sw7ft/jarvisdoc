const { DEFAULT_PAGE, readDocx, writeDocx } = require('./ooxml')

async function htmlToDocxBuffer(html, _title = 'Untitled', options = {}) {
  return writeDocx(html, { sourceBuffer: options.sourceBuffer || null })
}

async function docxToHtml(buffer) {
  return readDocx(buffer)
}

module.exports = { htmlToDocxBuffer, docxToHtml, DEFAULT_PAGE }
