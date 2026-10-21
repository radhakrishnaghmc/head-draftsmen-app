import HTMLtoDOCX from 'html-to-docx'
import PizZip from 'pizzip'
import { sanitizeHtmlDocxForWord2007 } from './word2007Compat'

/** html-to-docx's own portrait defaults (twips), so a caller overriding one
 * margin doesn't blank out the rest. */
const DEFAULT_MARGINS = { top: 1440, right: 1800, bottom: 1440, left: 1800, header: 720, footer: 720, gutter: 0 }

/** Convert filled document HTML into a real .docx file, formatting intact.
 * The output is passed through sanitizeHtmlDocxForWord2007 so it opens in Word
 * 2007 (html-to-docx emits invalid .rels content-types and out-of-order
 * property children that only Word 2010+ tolerates). */
export async function convertHtmlToDocx(
  html: string,
  margins?: Partial<typeof DEFAULT_MARGINS>
): Promise<Buffer> {
  const body = /<html/i.test(html) ? html : `<!DOCTYPE html><html><body>${html}</body></html>`
  const result = await HTMLtoDOCX(body, undefined, {
    table: { row: { cantSplit: true } },
    footer: false,
    header: false,
    margins: margins ? { ...DEFAULT_MARGINS, ...margins } : undefined
  })
  const buffer = Buffer.isBuffer(result) ? result : Buffer.from(result as ArrayBuffer)
  return sanitizeHtmlDocxForWord2007(compactRefTables(syncTableGridToCellWidths(buffer)))
}

/**
 * html-to-docx writes every table's <w:tblGrid> as equal columns no matter what
 * widths the cells carry (a Ref block's 600/520/7520 cells sit over a 2880/2880/2880
 * grid). Word sizes columns from the cells, but the in-app preview renderer follows the
 * grid — so the Ref rows of the SE Bid Evaluation / Agency Approval notes (and any
 * other hand-sized table) came out as three equal columns. Rewrites each grid from the
 * widths of the table's first row, when that row has one plain width per column.
 */
export function syncTableGridToCellWidths(buffer: Buffer): Buffer {
  const zip = new PizZip(buffer)
  const part = 'word/document.xml'
  const xml = zip.file(part)?.asText()
  if (!xml) return buffer
  const fixed = xml.replace(/<w:tbl>[\s\S]*?<\/w:tbl>/g, (tbl) => {
    const grid = /<w:tblGrid>([\s\S]*?)<\/w:tblGrid>/.exec(tbl)
    const firstRow = /<w:tr[ >][\s\S]*?<\/w:tr>/.exec(tbl)
    if (!grid || !firstRow || firstRow[0].includes('w:gridSpan') || firstRow[0].includes('<w:tbl>')) return tbl
    const widths = [...firstRow[0].matchAll(/<w:tcW w:w="(\d+)" w:type="dxa"\/>/g)].map((m) => m[1])
    const cols = (grid[1].match(/<w:gridCol /g) ?? []).length
    if (widths.length === 0 || widths.length !== cols) return tbl
    return tbl.replace(grid[0], `<w:tblGrid>${widths.map((w) => `<w:gridCol w:w="${w}"/>`).join('')}</w:tblGrid>`)
  })
  if (fixed === xml) return buffer
  zip.file(part, fixed)
  return zip.generate({ type: 'nodebuffer', compression: 'DEFLATE' })
}

/**
 * html-to-docx drops paragraph margins inside table cells, so every cell paragraph picks up the
 * document's 6pt after-spacing plus 4pt cell padding top and bottom — which spread the numbered
 * "Ref: 1) 2) 3) …" rows of the SE Bid Evaluation / Agency Approval notes far apart. A table whose
 * first cell reads "Ref:" is the Ref block: its cell padding goes to 0 and each row keeps a 2pt gap.
 */
export function compactRefTables(buffer: Buffer): Buffer {
  const zip = new PizZip(buffer)
  const part = 'word/document.xml'
  const xml = zip.file(part)?.asText()
  if (!xml) return buffer
  const fixed = xml.replace(/<w:tbl>[\s\S]*?<\/w:tbl>/g, (tbl) => {
    const firstCell = /<w:tc>[\s\S]*?<\/w:tc>/.exec(tbl)?.[0] ?? ''
    if (firstCell.replace(/<[^>]+>/g, '').trim() !== 'Ref:') return tbl
    return tbl
      .replace(/(<w:tblCellMar>)[\s\S]*?(<\/w:tblCellMar>)/, '$1<w:top w:type="dxa" w:w="0"/><w:left w:type="dxa" w:w="160"/><w:bottom w:type="dxa" w:w="0"/><w:right w:type="dxa" w:w="160"/>$2')
      .replace(/<w:spacing w:lineRule="auto"\/>|<w:spacing w:line="\d+" w:lineRule="auto"\/>/g, '<w:spacing w:before="0" w:after="40" w:line="240" w:lineRule="auto"/>')
  })
  if (fixed === xml) return buffer
  zip.file(part, fixed)
  return zip.generate({ type: 'nodebuffer', compression: 'DEFLATE' })
}
