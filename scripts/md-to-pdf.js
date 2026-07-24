const { app, BrowserWindow } = require('electron')
const fs = require('fs')
const MarkdownIt = require('markdown-it')

const mdPath = process.argv[2]
const pdfPath = process.argv[3]

const md = new MarkdownIt({ html: true, linkify: true, typographer: true })
const body = md.render(fs.readFileSync(mdPath, 'utf8'))

const html = `<!doctype html><html><head><meta charset="utf-8">
<style>
  @page { margin: 18mm 16mm; }
  * { box-sizing: border-box; }
  body {
    font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', system-ui, sans-serif;
    color: #1a1f2e; font-size: 10.5pt; line-height: 1.5; margin: 0;
  }
  h1 { font-size: 21pt; color: #4338ca; margin: 0 0 6pt; border-bottom: 2px solid #4f46e5; padding-bottom: 6pt; }
  h2 { font-size: 15pt; color: #4338ca; margin: 20pt 0 6pt; border-bottom: 1px solid #dfe3e8; padding-bottom: 3pt; page-break-after: avoid; }
  h3 { font-size: 12pt; color: #3730a3; margin: 14pt 0 4pt; page-break-after: avoid; }
  h4 { font-size: 10.5pt; color: #4b5563; margin: 10pt 0 4pt; page-break-after: avoid; }
  p, li { orphans: 2; widows: 2; }
  code { font-family: 'SF Mono', Menlo, Consolas, monospace; font-size: 9pt; background: #f3f4f6; padding: 1px 4px; border-radius: 4px; color: #b91c1c; }
  pre { background: #0f172a; color: #e2e8f0; padding: 12px 14px; border-radius: 8px; overflow-x: auto; font-size: 8.6pt; line-height: 1.45; page-break-inside: avoid; }
  pre code { background: transparent; color: inherit; padding: 0; }
  blockquote { border-left: 3px solid #a5b4fc; background: #eef2ff; margin: 8pt 0; padding: 6pt 12pt; color: #3730a3; border-radius: 0 6px 6px 0; }
  table { border-collapse: collapse; width: 100%; margin: 8pt 0; font-size: 9pt; page-break-inside: avoid; }
  th, td { border: 1px solid #d0d5dd; padding: 5px 8px; text-align: left; vertical-align: top; }
  th { background: #eef2ff; color: #3730a3; font-weight: 700; }
  tr:nth-child(even) td { background: #f9fafb; }
  a { color: #4f46e5; text-decoration: none; }
  hr { border: none; border-top: 1px solid #e5e7eb; margin: 16pt 0; }
  ul, ol { padding-left: 20pt; }
  strong { color: #111827; }
</style></head><body>${body}</body></html>`

app.disableHardwareAcceleration()
app.whenReady().then(async () => {
  const win = new BrowserWindow({ show: false, webPreferences: { offscreen: true } })
  await win.loadURL('data:text/html;charset=utf-8,' + encodeURIComponent(html))
  await new Promise((r) => setTimeout(r, 400))
  const pdf = await win.webContents.printToPDF({
    printBackground: true,
    pageSize: 'A4',
    margins: { marginType: 'default' }
  })
  fs.writeFileSync(pdfPath, pdf)
  console.log('WROTE ' + pdfPath + ' (' + pdf.length + ' bytes)')
  app.quit()
})
