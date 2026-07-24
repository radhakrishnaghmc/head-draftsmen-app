import { describe, it, expect } from 'vitest'
import * as ExcelJS from 'exceljs'
import PizZip from 'pizzip'
import { applyTechnicalSanctionEdits, stripWorkbookBloat } from '../core/technicalSanctionOutput'

async function buildSampleWorkbook(): Promise<Buffer> {
  const wb = new ExcelJS.Workbook()
  const ws = wb.addWorksheet('TS Copy')
  ws.addRow(['Sl. No.', 'Description', 'Rate'])
  ws.addRow([1, 'Earthwork excavation', 1])
  return Buffer.from(await wb.xlsx.writeBuffer())
}

// Injects the same kind of "phantom" bloat seen in a real sanctioned
// estimate — hundreds of external-link parts and thousands of defined names
// accumulated from years of copy-pasting cells from other workbooks — so the
// stripping logic can be exercised without needing that actual multi-MB file.
function injectBloat(buffer: Buffer, externalLinkCount: number, definedNameCount: number): Buffer {
  const zip = new PizZip(buffer)
  const workbookXml = zip.file('xl/workbook.xml')!.asText()

  let definedNames = ''
  for (let i = 0; i < definedNameCount; i++) {
    definedNames += `<definedName name="junk${i}">'TS Copy'!$A$1</definedName>`
  }
  const withDefinedNames = workbookXml.replace('</workbook>', `<definedNames>${definedNames}</definedNames></workbook>`)

  let externalReferences = ''
  const relsAdditions: string[] = []
  const contentTypeAdditions: string[] = []
  for (let i = 1; i <= externalLinkCount; i++) {
    externalReferences += `<externalReference r:id="rIdExt${i}"/>`
    zip.file(`xl/externalLinks/externalLink${i}.xml`, `<externalLink>${'x'.repeat(1000)}</externalLink>`)
    relsAdditions.push(
      `<Relationship Id="rIdExt${i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink" Target="externalLinks/externalLink${i}.xml"/>`
    )
    contentTypeAdditions.push(
      `<Override PartName="/xl/externalLinks/externalLink${i}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml"/>`
    )
  }
  const withExternalRefs = withDefinedNames.replace(
    '</workbook>',
    `<externalReferences>${externalReferences}</externalReferences></workbook>`
  )
  zip.file('xl/workbook.xml', withExternalRefs)

  const relsPath = 'xl/_rels/workbook.xml.rels'
  const relsXml = zip.file(relsPath)!.asText()
  zip.file(relsPath, relsXml.replace('</Relationships>', `${relsAdditions.join('')}</Relationships>`))

  const contentTypesXml = zip.file('[Content_Types].xml')!.asText()
  zip.file('[Content_Types].xml', contentTypesXml.replace('</Types>', `${contentTypeAdditions.join('')}</Types>`))

  return zip.generate({ type: 'nodebuffer', compression: 'DEFLATE' })
}

describe('stripWorkbookBloat', () => {
  it('leaves a clean workbook untouched', async () => {
    const clean = await buildSampleWorkbook()
    const result = stripWorkbookBloat(clean)
    expect(result).toBe(clean)
  })

  it('strips external links and defined names, and the result still loads correctly', async () => {
    const clean = await buildSampleWorkbook()
    const bloated = injectBloat(clean, 50, 200)
    expect(bloated.length).toBeGreaterThan(clean.length)

    const stripped = stripWorkbookBloat(bloated)
    expect(stripped.length).toBeLessThan(bloated.length)

    const zip = new PizZip(stripped)
    expect(zip.file(/^xl\/externalLinks\//).length).toBe(0)
    const workbookXml = zip.file('xl/workbook.xml')!.asText()
    expect(workbookXml).not.toContain('definedNames')
    expect(workbookXml).not.toContain('externalReferences')

    // The actual content survives untouched.
    const wb = new ExcelJS.Workbook()
    await wb.xlsx.load(stripped as unknown as ArrayBuffer)
    const ws = wb.getWorksheet('TS Copy')!
    expect(ws.getCell(2, 2).value).toBe('Earthwork excavation')
  })
})

describe('applyTechnicalSanctionEdits', () => {
  it('applies cell edits and creates Sheet 2 with rate-analysis rows even on a bloated workbook', async () => {
    const clean = await buildSampleWorkbook()
    const bloated = injectBloat(clean, 30, 100)

    const out = await applyTechnicalSanctionEdits(
      bloated,
      'TS Copy',
      [
        { row: 1, col: 2, value: 4445.91, color: 'green' as const }
      ],
      [['Rate Analysis: Earthwork excavation'], ['Cement', 'kg', '5.10']]
    )

    const wb = new ExcelJS.Workbook()
    await wb.xlsx.load(out as unknown as ArrayBuffer)
    const ws = wb.getWorksheet('TS Copy')!
    expect(ws.getCell(2, 3).value).toBe(4445.91)

    const sheet2 = wb.worksheets[1]
    expect(sheet2.name).toBe('Sheet 2')
    expect(sheet2.getCell(1, 1).value).toBe('Rate Analysis: Earthwork excavation')
    expect(sheet2.getCell(2, 1).value).toBe('Cement')
  })

  it('appends to an existing second sheet rather than replacing it', async () => {
    const wb = new ExcelJS.Workbook()
    wb.addWorksheet('TS Copy').addRow(['x'])
    const existingSecond = wb.addWorksheet('Existing Analysis')
    existingSecond.addRow(['already here'])
    const buffer = Buffer.from(await wb.xlsx.writeBuffer())

    const out = await applyTechnicalSanctionEdits(buffer, 'TS Copy', [], [['new row']])
    const outWb = new ExcelJS.Workbook()
    await outWb.xlsx.load(out as unknown as ArrayBuffer)
    expect(outWb.worksheets[1].name).toBe('Existing Analysis')
    expect(outWb.worksheets[1].getCell(1, 1).value).toBe('already here')
    expect(outWb.worksheets[1].getCell(2, 1).value).toBe('new row')
  })
})
