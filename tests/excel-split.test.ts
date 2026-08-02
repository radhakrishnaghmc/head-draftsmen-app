import { describe, it, expect, afterAll } from 'vitest'
import * as ExcelJS from 'exceljs'
import * as fs from 'fs'
import * as os from 'os'
import * as path from 'path'
import { splitWorkbookSheets } from '../electron/excelSplit'

const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'split-test-'))

afterAll(() => fs.rmSync(tmp, { recursive: true, force: true }))

async function makeWorkbook(file: string, sheetNames: string[]): Promise<void> {
  const wb = new ExcelJS.Workbook()
  for (const name of sheetNames) {
    const ws = wb.addWorksheet(name)
    ws.getCell('A1').value = `Estimate for ${name}`
    ws.getCell('A2').value = 10
    ws.getCell('B2').value = 5
    ws.getCell('C2').value = { formula: 'A2*B2', result: 50 }
    ws.mergeCells('A1:C1')
  }
  await wb.xlsx.writeFile(file)
}

describe('splitWorkbookSheets', () => {
  it('writes one file per sheet (named after the tab), flattening formulas to values and keeping merges', async () => {
    const src = path.join(tmp, 'multi.xlsx')
    await makeWorkbook(src, ['peddamma', 'kamalamma'])
    const out = path.join(tmp, 'out1')
    fs.mkdirSync(out)

    const files = await splitWorkbookSheets(src, out)
    expect(files.map((f) => path.basename(f))).toEqual(['peddamma.xlsx', 'kamalamma.xlsx'])

    const wb = new ExcelJS.Workbook()
    await wb.xlsx.readFile(path.join(out, 'peddamma.xlsx'))
    expect(wb.worksheets).toHaveLength(1)
    const ws = wb.worksheets[0]
    expect(ws.name).toBe('peddamma')
    // Formulas are pasted as values so a separated sheet never loses data to a
    // now-broken cross-sheet/external reference: C2 keeps its cached 50, no formula.
    expect(ws.getCell('C2').formula).toBeUndefined()
    expect(ws.getCell('C2').value).toBe(50)
    expect((ws.model.merges || []).length).toBe(1)
  })

  it('collapses whitespace and de-duplicates colliding output names', async () => {
    const src = path.join(tmp, 'collide.xlsx')
    // "a  b" (two spaces) and "a b" both sanitize to "a b"
    await makeWorkbook(src, ['a  b', 'a b'])
    const out = path.join(tmp, 'out2')
    fs.mkdirSync(out)

    const files = await splitWorkbookSheets(src, out).then((fs2) => fs2.map((f) => path.basename(f)))
    expect(files).toEqual(['a b.xlsx', 'a b (2).xlsx'])
  })
})
