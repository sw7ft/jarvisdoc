const { execFile } = require('child_process')
const { copyFile, readFile } = require('fs/promises')
const path = require('path')

const IS_MAC = process.platform === 'darwin'
const IS_WIN = process.platform === 'win32'

function appleString(value) {
  return JSON.stringify(String(value))
}

function run(command, args, options = {}) {
  return new Promise((resolve, reject) => {
    execFile(command, args, { timeout: 60000, windowsHide: true, ...options }, (error, stdout, stderr) => {
      if (error) {
        reject(new Error(String(stderr || error.message || `${command} failed`).trim()))
        return
      }
      resolve(String(stdout || ''))
    })
  })
}

// Pages renders Word layout far more faithfully than a browser can, and every Mac has it.
function pagesExportPdf(docxPath, pdfPath) {
  const script = `
set pagesWasRunning to application "Pages" is running
with timeout of 40 seconds
  tell application "Pages"
    set theDoc to open POSIX file ${appleString(docxPath)}
    export theDoc to POSIX file ${appleString(pdfPath)} as PDF
    close theDoc saving no
    if not pagesWasRunning then quit
  end tell
end timeout
`
  return run('osascript', ['-e', script], { timeout: 50000 })
}

const SOFFICE_CANDIDATES = {
  darwin: ['/Applications/LibreOffice.app/Contents/MacOS/soffice', 'soffice'],
  win32: [
    `${process.env['ProgramFiles'] || 'C:\\Program Files'}\\LibreOffice\\program\\soffice.exe`,
    `${process.env['ProgramFiles(x86)'] || 'C:\\Program Files (x86)'}\\LibreOffice\\program\\soffice.exe`,
    'soffice.exe'
  ],
  linux: ['/usr/bin/soffice', '/usr/bin/libreoffice', '/snap/bin/libreoffice', 'soffice', 'libreoffice']
}

let sofficePath
async function findSoffice() {
  if (sofficePath !== undefined) return sofficePath
  sofficePath = null
  for (const candidate of SOFFICE_CANDIDATES[process.platform] || ['soffice']) {
    try {
      await run(candidate, ['--version'], { timeout: 20000 })
      sofficePath = candidate
      break
    } catch {
      /* try the next one */
    }
  }
  return sofficePath
}

// LibreOffice writes <name>.pdf into an output directory rather than to a path we choose.
async function sofficeExportPdf(docxPath, pdfPath) {
  const soffice = await findSoffice()
  if (!soffice) throw new Error('LibreOffice is not installed.')
  const outDir = path.dirname(pdfPath)
  await run(soffice, ['--headless', '--norestore', '--convert-to', 'pdf', '--outdir', outDir, docxPath], {
    timeout: 120000
  })
  const produced = path.join(outDir, `${path.basename(docxPath, path.extname(docxPath))}.pdf`)
  if (produced !== pdfPath) await copyFile(produced, pdfPath)
}

// Windows can drive a real Word install when one is present.
function wordExportPdf(docxPath, pdfPath) {
  const script = `
$ErrorActionPreference = 'Stop'
$word = New-Object -ComObject Word.Application
$word.Visible = $false
try {
  $doc = $word.Documents.Open([string]${JSON.stringify(docxPath)}, $false, $true)
  $doc.ExportAsFixedFormat([string]${JSON.stringify(pdfPath)}, 17)
  $doc.Close($false)
} finally {
  $word.Quit()
}
`
  return run('powershell.exe', ['-NoProfile', '-NonInteractive', '-Command', script], { timeout: 120000 })
}

async function isPdf(filePath) {
  try {
    const buf = await readFile(filePath)
    return buf.length > 800 && buf.subarray(0, 5).toString() === '%PDF-'
  } catch {
    return false
  }
}

function converters() {
  if (IS_MAC) return [pagesExportPdf, sofficeExportPdf]
  if (IS_WIN) return [wordExportPdf, sofficeExportPdf]
  return [sofficeExportPdf]
}

// Tries the best native converter available, in order. The caller falls back to the
// browser's own printing when every one of these is unavailable.
async function convertToPdf(docxPath, pdfPath) {
  const failures = []
  for (const convert of converters()) {
    try {
      await convert(docxPath, pdfPath)
      if (await isPdf(pdfPath)) return convert.name
      failures.push(`${convert.name}: produced no PDF`)
    } catch (err) {
      failures.push(`${convert.name}: ${err instanceof Error ? err.message : String(err)}`)
    }
  }
  throw new Error(failures.join('; ') || 'No PDF converter available.')
}

module.exports = { convertToPdf, pagesExportPdf, sofficeExportPdf, isPdf }
