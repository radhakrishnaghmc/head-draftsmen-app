import { describe, it, expect } from 'vitest'
import { detectCollisions, mergeTables } from '../core/merge'
import type { ExcelTable } from '../core/types'

function table(name: string, headers: string[], rows: Record<string, string>[]): ExcelTable {
  return { id: name, name, path: `/tmp/${name}`, headers, rows }
}

describe('detectCollisions', () => {
  it('flags columns shared across files', () => {
    const a = table('a.xlsx', ['Name', 'Email'], [])
    const b = table('b.xlsx', ['Name', 'Phone'], [])
    const collisions = detectCollisions([a, b])
    expect(collisions).toHaveLength(1)
    expect(collisions[0].column).toBe('Name')
    expect(collisions[0].sources.sort()).toEqual(['a.xlsx', 'b.xlsx'])
  })

  it('returns nothing when columns are unique', () => {
    const a = table('a.xlsx', ['Name'], [])
    const b = table('b.xlsx', ['Phone'], [])
    expect(detectCollisions([a, b])).toHaveLength(0)
  })
})

describe('mergeTables', () => {
  it('merges by row index and unions columns', () => {
    const a = table('a.xlsx', ['Name'], [{ Name: 'Ann' }, { Name: 'Bob' }])
    const b = table('b.xlsx', ['City'], [{ City: 'NY' }, { City: 'LA' }])
    const merged = mergeTables([a, b])
    expect(merged.columns.map((c) => c.name)).toEqual(['Name', 'City'])
    expect(merged.rows).toEqual([
      { Name: 'Ann', City: 'NY' },
      { Name: 'Bob', City: 'LA' }
    ])
  })

  it('honors collision resolution to pick the winning source', () => {
    const a = table('a.xlsx', ['Name'], [{ Name: 'FromA' }])
    const b = table('b.xlsx', ['Name'], [{ Name: 'FromB' }])
    const merged = mergeTables([a, b], { Name: 'b.xlsx' })
    expect(merged.rows[0].Name).toBe('FromB')
    expect(merged.columns.find((c) => c.name === 'Name')?.source).toBe('b.xlsx')
  })

  it('pads shorter files with empty values', () => {
    const a = table('a.xlsx', ['Name'], [{ Name: 'Ann' }, { Name: 'Bob' }])
    const b = table('b.xlsx', ['City'], [{ City: 'NY' }])
    const merged = mergeTables([a, b])
    expect(merged.rows[1]).toEqual({ Name: 'Bob', City: '' })
  })
})

describe('parseExcelBuffer header detection', () => {
  it('skips a merged title row and detects the real header row', async () => {
    const XLSX = await import('xlsx')
    const { parseExcelBuffer } = await import('../core/excel')
    const ws = XLSX.utils.aoa_to_sheet([
      ['List of works sanctioned'],
      ['Sl.NO.', 'Name', 'Amount', '', 'Remarks'],
      ['1', 'Road work', '20.00', '', 'ok'],
      ['2', 'Drain work', '13.00', '', 'ok']
    ])
    const wb = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(wb, ws, 'Sheet1')
    const buf = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' }) as Buffer
    const table = parseExcelBuffer(buf, 'works.xlsx')
    expect(table.headers).toEqual(['Sl.NO.', 'Name', 'Amount', 'Remarks'])
    expect(table.rows).toHaveLength(2)
    expect(table.rows[0]).toMatchObject({ 'Sl.NO.': '1', Name: 'Road work', Remarks: 'ok' })
  })
})

describe('buildTableFromGrid (Header Row Picker)', () => {
  it('builds a table using an explicitly chosen header row', async () => {
    const { buildTableFromGrid, guessHeaderRow } = await import('../core/sheet')
    const grid = [
      ['Monthly report — confidential'],
      ['generated 2026'],
      ['Name', 'Email', '', 'City'],
      ['Ada', 'ada@x.io', '', 'London'],
      ['Grace', 'grace@x.io', '', 'NYC']
    ]
    expect(guessHeaderRow(grid)).toBe(2)
    const table = buildTableFromGrid(grid, 2, { id: 'x', name: 'r.xlsx', path: '/r.xlsx' })
    expect(table.headers).toEqual(['Name', 'Email', 'City'])
    expect(table.rows).toHaveLength(2)
    expect(table.rows[0]).toMatchObject({ Name: 'Ada', Email: 'ada@x.io', City: 'London' })
  })

  it('names blank header cells that still carry data', async () => {
    const { buildTableFromGrid } = await import('../core/sheet')
    const grid = [
      ['A', '', 'C'],
      ['1', '2', '3']
    ]
    const table = buildTableFromGrid(grid, 0, { id: 'x', name: 'r.xlsx', path: '/r.xlsx' })
    expect(table.headers).toEqual(['A', 'Column 2', 'C'])
    expect(table.rows[0]).toMatchObject({ A: '1', 'Column 2': '2', C: '3' })
  })

  it('does not let a fully-populated numeric data row outscore the real (text) header row', async () => {
    const { guessHeaderRow } = await import('../core/sheet')
    // Header row has a genuinely blank trailing column (an unlabeled serial
    // number column) while every data row below has that column filled in —
    // pure non-empty-cell counting would pick the data row as the header.
    const grid = [
      ['Quantity', 'Description', 'Rate', ''],
      ['10', 'Earth work', '100', '1'],
      ['20', 'Concreting', '200', '2']
    ]
    expect(guessHeaderRow(grid)).toBe(0)
  })
})

describe('readExcelGrid sheet selection', () => {
  it('picks the first visible sheet, skipping a hidden lead sheet', async () => {
    const XLSX = await import('xlsx')
    const { parseExcelBuffer } = await import('../core/excel')
    const wb = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([['Summary'], ['x', 'y']]), 'Sheet1')
    XLSX.utils.book_append_sheet(
      wb,
      XLSX.utils.aoa_to_sheet([['Sl.NO.', 'Name'], ['1', 'Road']]),
      'Sheet2'
    )
    wb.Workbook = { Sheets: [{ Hidden: 1 }, { Hidden: 0 }] }
    const buf = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' }) as Buffer
    const table = parseExcelBuffer(buf, 'multi.xlsx')
    expect(table.headers).toEqual(['Sl.NO.', 'Name'])
    expect(table.rows).toHaveLength(1)
    expect(table.rows[0]).toMatchObject({ 'Sl.NO.': '1', Name: 'Road' })
  })
})

describe('calendar helpers', () => {
  it('holidayDay extracts the day number', async () => {
    const { holidayDay } = await import('../core/calendar')
    expect(holidayDay('Jul 12')).toBe(12)
    expect(holidayDay('Jan 26 – 28')).toBe(26)
    expect(holidayDay('no date')).toBeNull()
  })

  it('holidaysByDay maps public holidays for a month', async () => {
    const { holidaysByDay } = await import('../core/calendar')
    const data = {
      year: '2026',
      fetchedAt: '',
      months: [
        {
          month: 'January',
          holidays: [
            { date: 'Jan 26', name: 'Republic Day', type: 'public' as const },
            { date: 'Jan 14', name: 'Bhogi', type: 'optional' as const }
          ]
        }
      ]
    }
    const map = holidaysByDay(data, 0)
    expect(map.get(26)?.type).toBe('public')
    expect(map.get(14)?.type).toBe('optional')
    expect(map.get(1)).toBeUndefined()
  })
})

describe('extractEstimateItems', () => {
  it('keeps a normal single-row item as-is', async () => {
    const { extractEstimateItems } = await import('../core/estimateExtract')
    const header = ['S.No', 'Description', 'Qty', 'Rate', 'Unit']
    const grid = [header, ['1', 'Earth work excavation', '237.06', '308.31', 'Cum']]
    const items = extractEstimateItems(grid, 0)
    expect(items).toMatchObject([{ description: 'Earth work excavation', quantity: '237.06', rate: '308.31', unit: 'Cum' }])
  })

  it('stops at the grand Total row, ignoring the abstract and a second estimate stacked in the same sheet', async () => {
    const { extractEstimateItems } = await import('../core/estimateExtract')
    const header = ['S. No.', 'Description of work', 'Qty', 'Rate', 'Per', 'Amount']
    const grid = [
      header,
      ['1', 'Engaging of Tractor', '1088', '512', 'Hr', '557056'],
      ['2', 'Engaging of labour', '544', '847', '1 day', '460768'],
      // Estimate 1's abstract section — must NOT be read as items.
      ['', 'Total', 'Total', 'Total', 'Total', '1017824'],
      ['', 'Add Labour charges @ 1%', '', '', '', '10178'],
      ['', 'Add GST @ 18%', '', '', '', '183240'],
      ['', 'or say Rs.', '', '', '30', 'Lakhs'],
      // A second, unrelated estimate pasted below — must NOT bleed in.
      ['GREATER HYDERABAD MUNICIPAL CORPORATION', '', '', '', '', ''],
      ['S. No.', 'Description of work', 'Qty', 'Rate', 'Per', 'Amount'],
      ['1', 'Job work for engaging of JCB', '1440', '800', 'Hour', '1152000']
    ]
    const items = extractEstimateItems(grid, 0)
    expect(items.map((i) => i.description)).toEqual(['Engaging of Tractor', 'Engaging of labour'])
    expect(items.some((i) => /JCB/.test(i.description))).toBe(false)
  })

  it('keeps an item that has a quantity and rate but a blank unit cell', async () => {
    const { extractEstimateItems } = await import('../core/estimateExtract')
    const header = ['Sl. No', 'Description', 'No.s', 'L', 'B', 'D', 'Qty', 'Rate', 'Per', 'Unit', 'Amount']
    const grid = [
      header,
      ['1', 'Earth work excavation', '', '', '', '', '372.15', '308.31', '1', 'Cum', '114738'],
      // Manhole covers: quantity + rate present, unit cell left blank in the source.
      ['2', 'Manufacture & supply of manhole covers', '', '', '', '', '4', '2716', '1', '', '10864']
    ]
    const items = extractEstimateItems(grid, 0)
    expect(items.map((i) => ({ description: i.description, quantity: i.quantity, rate: i.rate, unit: i.unit }))).toEqual([
      { description: 'Earth work excavation', quantity: '372.15', rate: '308.31', unit: 'Cum' },
      { description: 'Manufacture & supply of manhole covers', quantity: '4', rate: '2716', unit: '' }
    ])
  })

  it('correctHeaderTypos fixes distinctive column-word typos but leaves Rate/Date/short headers alone', async () => {
    const { correctHeaderTypos } = await import('../core/estimateExtract')
    // Quantity family (substitution + transposition).
    expect(correctHeaderTypos('Qantity')).toBe('quantity')
    expect(correctHeaderTypos('Quntity')).toBe('quantity')
    expect(correctHeaderTypos('Qauntity')).toBe('quantity') // transposition
    expect(correctHeaderTypos('Total Qantity')).toBe('Total quantity') // only the word fixed
    // Amount / Unit.
    expect(correctHeaderTypos('Amout')).toBe('amount')
    expect(correctHeaderTypos('Uint')).toBe('unit') // transposition
    expect(correctHeaderTypos('Units')).toBe('unit')
    // Left alone — no false positives.
    expect(correctHeaderTypos('Date')).toBe('Date') // one edit from "rate", must NOT become rate
    expect(correctHeaderTypos('Rate Rs.')).toBe('Rate Rs.')
    expect(correctHeaderTypos('No.')).toBe('No.')
    expect(correctHeaderTypos('L')).toBe('L')
  })

  it('resolves a mis-spelled "Qantity" quantity header (real footpath estimate)', async () => {
    const { extractEstimateItems } = await import('../core/estimateExtract')
    const header = ['Sl. No.', 'Description of Work', 'No.', 'L', 'B', 'D', 'Qantity', 'Rate', 'Unit', 'Amount']
    const grid = [
      header,
      ['1', 'Earth work excavation', '', '', '', '', '111.78', '350.28', 'cum', '39154'],
      ['2', 'Supply of precast CC blocks', '', '', '', '', '2070', '200', 'each', '414000']
    ]
    const items = extractEstimateItems(grid, 0)
    expect(items.map((i) => ({ q: i.quantity, r: i.rate, u: i.unit }))).toEqual([
      { q: '111.78', r: '350.28', u: 'cum' },
      { q: '2070', r: '200', u: 'each' }
    ])
  })

  it('resolves the unit column whether it is headed Unit, UOM, or Per', async () => {
    const { extractEstimateItems } = await import('../core/estimateExtract')
    for (const unitHeader of ['Unit', 'Units', 'UOM']) {
      const grid = [['S.No', 'Description', 'Qty', 'Rate', unitHeader], ['1', 'Earth work excavation', '237.06', '308.31', 'Cum']]
      expect(extractEstimateItems(grid, 0)[0]).toMatchObject({ quantity: '237.06', rate: '308.31', unit: 'Cum' })
    }
  })

  it('uses the "Total Qty" column, not a blank "Qty per Day", in a day-rate estimate', async () => {
    const { extractEstimateItems } = await import('../core/estimateExtract')
    // A day-rate estimate carries both a "Qty per Day" column (blank on most
    // rows — the qty is built from nos × days × months) and a "Total Qty"
    // column with the final figure. The plain /qty/ pattern used to grab the
    // leftmost ("Qty per Day"), dropping every item whose per-day cell is empty.
    const header = ['Sl. No.', 'Description of Item', 'Qty per Day', 'nos', 'No. of days', 'No of Months', 'Total Qty', 'Rate', 'Unit', 'Amount']
    const grid = [
      header,
      ['1', 'Engaging One Tractor', '', '1', '10', '4', '', '', '', ''],
      ['', '', '', '', '', '', '320', '512', 'hour', '163840'],
      ['2', 'Providing UNSKILLED WORKMEN', '', '1', '10', '4', '', '', '', ''],
      ['', '', '', '', '', '', '120', '847', 'Day', '101640']
    ]
    const items = extractEstimateItems(grid, 0)
    expect(items).toMatchObject([
      { description: 'Engaging One Tractor', quantity: '320', rate: '512', unit: 'hour' },
      { description: 'Providing UNSKILLED WORKMEN', quantity: '120', rate: '847', unit: 'Day' }
    ])
  })

  it('does not stop at a bare "Total" label used as an intra-item / sub-work subtotal, only at a Total in the Qty column', async () => {
    const { extractEstimateItems } = await import('../core/estimateExtract')
    const header = ['S.No', 'Description of Item', 'Nos', 'L', 'B', 'D', 'Quantity', 'Rate', 'Unit', 'Amount']
    const grid = [
      header,
      // Sub-work 1
      ['1', 'Earth work excavation', '', '', '', '', '34.92', '308.31', 'Cum', '10766'],
      // An item whose measured parts are summed under a bare "Total" label
      // (description column), with the real Qty/Rate a couple of rows below.
      ['2', 'Flooring with Shabad stones', '', '', '', '', '', '', '', ''],
      ['', 'Total', '', '', '', '', '', '', '', ''],
      ['', 'Waiting room', '1', '1', '9.23', '4.23', '39.04', '', '', ''],
      ['', '', '', '', '', '', '39.04', '894.61', 'Sqm', '34926'],
      // Sub-work 1's own total (amount in Amount column, Qty blank).
      ['', '', '', '', '', '', 'Total:', '', '', '45692'],
      // Sub-work 2 starts — a bare "Total" label alone must not have ended things.
      ['Construction of Toilet Block', '', '', '', '', '', '', '', '', ''],
      ['1', 'RCC M25 for toilet slab', '', '', '', '', '4.36', '7453.8', 'Cum', '32499'],
      // True end of items: the final "Total" sits in the Quantity column.
      ['', '', '', '', '', '', 'Total', '', '', '78191'],
      ['', 'Provision towards GST @ 18%', '', '', '', '', '', '', '', '14074'],
      ['', 'Total', '', '', '', '', '', '', '', '92265']
    ]
    const items = extractEstimateItems(grid, 0)
    expect(items.map((i) => i.description)).toEqual([
      'Earth work excavation',
      'Flooring with Shabad stones',
      'RCC M25 for toilet slab'
    ])
  })

  it('prefixes lettered sub-parts with their parent item spec, and leaves self-contained items unprefixed', async () => {
    const { extractEstimateItems } = await import('../core/estimateExtract')
    const header = ['S.No', 'Description of Item', 'Nos', 'L', 'B', 'D', 'Quantity', 'Rate', 'Unit', 'Amount']
    const spec = 'Supply and placing of the M-25 Design Mix Concrete corresponding to IS 456 using weigh batcher'
    const grid = [
      header,
      // Parent item: full spec, no measurement of its own.
      ['3', spec, '', '', '', '', '', '', '', ''],
      ['a', 'Footings', '', '', '', '', '', '', '', ''],
      ['', 'F1', '1', '2', '1.3', '1.3', '0.85', '', '', ''],
      ['', '', '', '', '', '', '0.85', '11031.32', 'Cum', '9377'],
      ['b', 'Pedastals', '', '', '', '', '', '', '', ''],
      ['', '', '1', '4', '0.45', '0.45', '0.49', '11031.32', 'Cum', '5405'],
      // A self-contained full-spec item ends the sub-part run.
      ['4', 'Filling with useful available excavated earth in trenches', '', '', '', '', '10.8', '43.76', 'Cum', '473'],
      // A short label after a self-contained item must NOT inherit item 3's spec.
      ['5', 'Extra soil disposal', '', '', '', '', '5', '100', 'Cum', '500']
    ]
    const items = extractEstimateItems(grid, 0)
    expect(items.map((i) => i.description)).toEqual([
      `${spec} - Footings`,
      `${spec} - Pedastals`,
      'Filling with useful available excavated earth in trenches',
      'Extra soil disposal'
    ])
  })

  it('swaps a per-row Rate/Unit inversion where the unit token sits in the Rate column', async () => {
    const { extractEstimateItems } = await import('../core/estimateExtract')
    const header = ['S.No', 'Description of Item', 'Nos', 'L', 'B', 'D', 'Quantity', 'Rate', 'Unit', 'Amount']
    const grid = [
      header,
      // "Say" summary row: unit "Cum" landed in Rate, the rate number in Unit.
      ['7', 'RCC roof beam for arch', '', '', '', '', '0.43', 'Cum', '7453.8', '3205']
    ]
    const items = extractEstimateItems(grid, 0)
    expect(items[0]).toMatchObject({ description: 'RCC roof beam for arch', quantity: '0.43', rate: '7453.8', unit: 'Cum' })
  })

  it('starts a new item on a spec row whose S.No was left blank (description begins with a schedule code)', async () => {
    const { extractEstimateItems } = await import('../core/estimateExtract')
    const header = ['S.No', 'Description of Item', 'Nos', 'L', 'B', 'D', 'Quantity', 'Rate', 'Unit', 'Amount']
    const door = 'TBSC-L.III-02: Providing and fixing 30mm thick factory made PVC door shutter'
    const window = 'TBSC-M.II-02: Providing and fixing of two shutter sliding windows'
    const grid = [
      header,
      ['19', door, '', '', '', '', '', '', '', ''],
      ['', 'Door D1', '1', '8', '0.75', '2.1', '', '', '', ''],
      ['', '', '', '', '', '', '16.8', '2618', 'sqm', '43982'],
      // Next item's spec, but its S.No cell was left blank in the source.
      ['', window, '', '', '', '', '', '', '', ''],
      ['', 'window', '1', '2', '1.2', '1.2', '', '', '', ''],
      ['', '', '', '', '', '', '2.88', '6120.44', 'sqm', '17627']
    ]
    const items = extractEstimateItems(grid, 0)
    expect(items.map((i) => ({ description: i.description, quantity: i.quantity }))).toEqual([
      { description: door, quantity: '16.8' },
      { description: window, quantity: '2.88' }
    ])
  })

  it('flags items with a quantity and rate but a blank or zero Amount (itemsMissingEstimateAmount)', async () => {
    const { extractEstimateItems, itemsMissingEstimateAmount } = await import('../core/estimateExtract')
    const header = ['S.No', 'Description of Item', 'Nos', 'L', 'B', 'D', 'Quantity', 'Rate', 'Unit', 'Amount']
    const grid = [
      header,
      ['1', 'Impervious coat to RCC roof slab', '', '', '', '', '59.04', '533.41', 'Sqm', ''], // Amount blank
      ['2', 'Rolling shutter for wood storage', '', '', '', '', '5.48', '4017.43', 'Sqm', '0'], // Amount zero
      ['3', 'Earth work excavation', '', '', '', '', '34.92', '308.31', 'Cum', '10766'], // costed
      ['4', 'Lintel not executed', '', '', '', '', '0', '13916.2', 'Cum', ''] // zero qty — genuinely no work
    ]
    const items = extractEstimateItems(grid, 0)
    const missing = itemsMissingEstimateAmount(items)
    expect(missing.map((i) => i.description)).toEqual(['Impervious coat to RCC roof slab', 'Rolling shutter for wood storage'])
  })

  it('reports no missing-amount items when the estimate has no Amount column', async () => {
    const { extractEstimateItems, itemsMissingEstimateAmount } = await import('../core/estimateExtract')
    const header = ['S.No', 'Description', 'Qty', 'Rate', 'Unit']
    const grid = [header, ['1', 'Earth work excavation', '237.06', '308.31', 'Cum']]
    const items = extractEstimateItems(grid, 0)
    expect(items[0].estimateAmount).toBeUndefined()
    expect(itemsMissingEstimateAmount(items)).toEqual([])
  })

  it('splits a drilling item with depth-range sub-rows into one item per depth, tagging each with its range', async () => {
    const { extractEstimateItems } = await import('../core/estimateExtract')
    const header = ['S.No', 'Description', 'No', 'Length', 'Qty', 'Rate', 'Unit', 'Amount']
    const desc =
      'Drilling of tube wells by down the hole hammer drilling, for a finalised dix of 150mm of in all formations suitable for down the hole hammer drilling back or medium rack etc., and drilling in overburden to a dia not less than 180mm etc complete for finished item of work.'
    const grid = [
      header,
      ['2', desc, '', '', '', '', '', ''],
      ['', '0 to 30 mtrs', '1', '3', '90.00', '298.00', 'Rmt', '26820.00'],
      ['', '30 to 60 mtrs', '1', '3', '90.00', '373.00', 'Rmt', '33570.00'],
      ['', '60 to 90 mtrs', '1', '3', '90.00', '447.00', 'Rmt', '40230.00']
    ]
    const items = extractEstimateItems(grid, 0)
    expect(items).toHaveLength(3)
    expect(items[0]).toMatchObject({ description: `${desc}( 0 to 30 mtrs)`, quantity: '90.00', rate: '298.00', unit: 'Rmt' })
    expect(items[1]).toMatchObject({ description: `${desc}( 30 to 60 mtrs)`, quantity: '90.00', rate: '373.00', unit: 'Rmt' })
    expect(items[2]).toMatchObject({ description: `${desc}( 60 to 90 mtrs)`, quantity: '90.00', rate: '447.00', unit: 'Rmt' })
  })

  it('still resolves a normal block with intermediate dimension rows and one final summary row', async () => {
    const { extractEstimateItems } = await import('../core/estimateExtract')
    const header = ['S.No', 'Description', 'Length', 'Breadth', 'Qty', 'Rate', 'Unit']
    const grid = [
      header,
      ['3', 'Plain Cement concrete', '', '', '', '', ''],
      ['', '', '10', '2', '', '', ''],
      ['', '', '', '', '1.55', '6587.04', 'Cum']
    ]
    const items = extractEstimateItems(grid, 0)
    expect(items).toMatchObject([{ description: 'Plain Cement concrete', quantity: '1.55', rate: '6587.04', unit: 'Cum' }])
  })

  it('tracks cell positions: normal item points at the lead row, a variant points at its own row', async () => {
    const { extractEstimateItems } = await import('../core/estimateExtract')
    const header = ['S.No', 'Description', 'Length', 'Breadth', 'Qty', 'Rate', 'Unit']
    const grid = [
      header,
      ['3', 'Plain Cement concrete', '', '', '', '', ''],
      ['', '', '10', '2', '', '', ''],
      ['', '', '', '', '1.55', '6587.04', 'Cum']
    ]
    const items = extractEstimateItems(grid, 0)
    expect(items[0]).toMatchObject({ descRow: 1, descCol: 1, rateRow: 3, rateCol: 5, isVariant: false })

    const desc = 'Drilling of tube wells'
    const variantGrid = [
      header,
      ['2', desc, '', '', '', '', ''],
      ['', '0 to 30 mtrs', '', '', '90.00', '298.00', 'Rmt'],
      ['', '30 to 60 mtrs', '', '', '90.00', '373.00', 'Rmt']
    ]
    const variantItems = extractEstimateItems(variantGrid, 0)
    expect(variantItems[0]).toMatchObject({ descRow: 2, descCol: 1, rateRow: 2, rateCol: 5, isVariant: true })
    expect(variantItems[1]).toMatchObject({ descRow: 3, descCol: 1, rateRow: 3, rateCol: 5, isVariant: true })
  })

  it('resolves a Quantity column no regex would recognize, via the embedding fallback', async () => {
    const { extractEstimateItems, estimateColumnsMatchedViaEmbedding, ESTIMATE_COLUMN_SPECS } = await import(
      '../core/estimateExtract'
    )
    // "Nos." matches none of Serial Number/Quantity/Rate/Unit's regexes.
    const header = ['S.No', 'Description', 'Nos.', 'Rate', 'Unit']
    const grid = [header, ['1', 'Earth work excavation', '10', '100', 'Cum']]

    expect(() => extractEstimateItems(grid, 0)).toThrow(/Could not find/)

    // One embedding per header (S.No, Description, Nos., Rate, Unit), then one per spec
    // (Serial Number, Quantity, Rate, Unit) — "Nos." is deliberately closest to Quantity's.
    const embeddings = {
      headerVectors: [
        [0, 1, 0, 0],
        [0, 0, 1, 0],
        [1, 0, 0, 0],
        [0.9, 0, 0, 0.1],
        [0, 0, 0, 1]
      ],
      labelVectors: ESTIMATE_COLUMN_SPECS.map((spec) =>
        spec.label === 'Quantity' ? [1, 0, 0, 0] : spec.label === 'Rate' ? [0.9, 0, 0, 0.1] : [0, 0, 0, 0]
      )
    }

    const items = extractEstimateItems(grid, 0, embeddings)
    expect(items).toMatchObject([{ description: 'Earth work excavation', quantity: '10', rate: '100', unit: 'Cum' }])
    expect(estimateColumnsMatchedViaEmbedding(grid, 0, embeddings)).toEqual(['Quantity'])
  })

  it('finds the real unit column when a "Per" multiplier column (always "1") sits before it and its own header is blank', async () => {
    const { extractEstimateItems } = await import('../core/estimateExtract')
    // Real department template shape: Rate | Per (bare "1") | <blank header, real unit text> | Amount.
    const header = ['S.No', 'Description', 'Qty', 'Rate', 'Per', '', 'Amount']
    const grid = [header, ['1', 'Earth work excavation', '285.20', '308.31', '1', 'Cum', '87930']]
    const items = extractEstimateItems(grid, 0)
    expect(items).toMatchObject([{ description: 'Earth work excavation', quantity: '285.20', rate: '308.31', unit: 'Cum' }])
  })

  it('finds the unit column from the data when the estimate has no Unit header at all (detailed No.s/L/B/D layout)', async () => {
    const { extractEstimateItems } = await import('../core/estimateExtract')
    // The real CMC detailed-estimate shape: no Unit/UOM header anywhere — the
    // unit (Cum/Sqm/…) sits in an unlabelled column between Rate and Amount,
    // filled only on each item's summary line. Previously threw
    // "Could not find S.No / Qty / Rate / Unit columns in the estimate."
    const header = ['Sl. No', 'Description', 'No.s', '', '', 'L', 'B', 'D', 'Qty', 'Rate', '', 'Amount']
    const grid = [
      header,
      ['1', 'Earth work excavation', '', '', '', '', '', '', '', '', '', ''],
      ['', '', '1', 'X', '1', '462.00', '6.50', '0.150', '450.45', '', '', ''],
      ['', '', '', '', '', '', '', '', '450.45', '308.31', 'Cum', '138878.00'],
      ['2', 'WMM base course', '', '', '', '', '', '', '', '', '', ''],
      ['', '', '1', 'X', '1', '462.00', '6.30', '0.100', '291.06', '', '', ''],
      ['', '', '', '', '', '', '', '', '291.06', '2086.57', 'sqm', '607317.00']
    ]
    const items = extractEstimateItems(grid, 0)
    expect(items).toMatchObject([
      { description: 'Earth work excavation', quantity: '450.45', rate: '308.31', unit: 'Cum' },
      { description: 'WMM base course', quantity: '291.06', rate: '2086.57', unit: 'sqm' }
    ])
  })

  it('keeps "Per" as the unit column when it is not overwhelmingly numeric (no multiplier-column quirk present)', async () => {
    const { extractEstimateItems } = await import('../core/estimateExtract')
    const header = ['S.No', 'Description', 'Qty', 'Rate', 'Per']
    const grid = [header, ['1', 'Earth work excavation', '285.20', '308.31', 'Cum']]
    const items = extractEstimateItems(grid, 0)
    expect(items).toMatchObject([{ unit: 'Cum' }])
  })

  it('reads No\'s/L/B/D from their own dimension row, not just Qty/Rate/Unit', async () => {
    const { extractEstimateItems } = await import('../core/estimateExtract')
    const header = ['S.No', 'Description', "No's", 'L', 'B', 'D', 'Qty', 'Rate', 'Unit']
    const grid = [
      header,
      ['1', 'Providing wet mix macadam', '', '', '', '', '', '', ''],
      ['', '', '5', '10', '2', '0.3', '', '', ''],
      ['', '', '', '', '', '', '30.00', '2000.00', 'Cum']
    ]
    const items = extractEstimateItems(grid, 0)
    expect(items).toMatchObject([
      { description: 'Providing wet mix macadam', quantity: '30.00', rate: '2000.00', unit: 'Cum', nos: '5', l: '10', b: '2', d: '0.3' }
    ])
  })

  it('keeps each depth-range variant\'s own dimensions separate, not bled from an earlier variant', async () => {
    const { extractEstimateItems } = await import('../core/estimateExtract')
    const header = ['S.No', 'Description', "No's", 'L', 'B', 'D', 'Qty', 'Rate', 'Unit']
    const grid = [
      header,
      ['2', 'Drilling of tube wells', '', '', '', '', '', '', ''],
      ['', '0 to 30 mtrs', '1', '30', '', '', '90.00', '298.00', 'Rmt'],
      ['', '30 to 60 mtrs', '1', '30', '', '', '90.00', '373.00', 'Rmt']
    ]
    const items = extractEstimateItems(grid, 0)
    expect(items).toHaveLength(2)
    expect(items[0]).toMatchObject({ nos: '1', l: '30' })
    expect(items[1]).toMatchObject({ nos: '1', l: '30' })
  })
})

describe('extractEstimateItemsFromLines', () => {
  it('extracts qty/rate/unit from a summary line, ignoring title-block and header lines before it', async () => {
    const { extractEstimateItemsFromLines } = await import('../core/estimateExtract')
    const lines = [
      'NIZAMPET MUNICIPAL CORPORATION :: MEDCHAL-MALKAJGIRI DISTRICT',
      'Name of the work: Laying of UGD Line',
      'Sl. Description of work No\'s L B D',
      'Qty Rate Per Amount',
      'Cutting openCC road surface as directed by the departmental officers.',
      '99.00 2345.00 1 Cum 232155.00'
    ]
    const items = extractEstimateItemsFromLines(lines)
    expect(items).toMatchObject([
      {
        description: 'Cutting openCC road surface as directed by the departmental officers.',
        quantity: '99.00',
        rate: '2345.00',
        unit: 'Cum'
      }
    ])
  })

  it('finds service/transport items priced per tonne (a unit beyond the civil-works set)', async () => {
    const { extractEstimateItemsFromLines } = await import('../core/estimateExtract')
    // The real garbage-transport estimate: two items, both measured in "tonne".
    const lines = [
      'DETAILED CUM ABSTRACT ESTIMATE',
      'Name of the work: Transportation of Garbage from Nizampet Transfer Station',
      'Sl. Description of work No L B D Qty Rate/per per Amount',
      'Hiring of JCB for Lifting and Loading the Garbage into tippers per tonne',
      '1 x 1 14960.00 14960.00',
      '14960.00 166.76 tonne 2494730.00',
      'Engaging Tipper for conveyance of Garbage to Dumping Yard at 12 Tonne per trip',
      '1 x 1 14960.00 14960.00',
      '14960.00 372.84 tonne 5577686.00'
    ]
    const items = extractEstimateItemsFromLines(lines)
    expect(items.map((i) => ({ quantity: i.quantity, rate: i.rate, unit: i.unit }))).toEqual([
      { quantity: '14960.00', rate: '166.76', unit: 'tonne' },
      { quantity: '14960.00', rate: '372.84', unit: 'tonne' }
    ])
  })

  it('re-stitches a summary row the OCR split into out-of-order fragments', async () => {
    const { extractEstimateItemsFromLines } = await import('../core/estimateExtract')
    // Exactly what PaddleOCR produced for the second item of the garbage
    // estimate: its Qty/Rate cell and unit/Amount cell became two lines, in the
    // wrong order (the Amount line sorted above the Qty/Rate line).
    const lines = [
      'Sl. Description of work No L B D Qty Rate/per per Amount',
      'Hiring of JCB per tonne',
      '14960.00 166.76 tonne 2494730.00',
      'Engaging Tipper for conveyance of Garbage per tonne',
      '14960.00 14960.00',
      '1x 1',
      'tonne 5577686.00',
      '14960.00 372.84'
    ]
    const items = extractEstimateItemsFromLines(lines)
    expect(items.map((i) => ({ quantity: i.quantity, rate: i.rate, unit: i.unit }))).toEqual([
      { quantity: '14960.00', rate: '166.76', unit: 'tonne' },
      { quantity: '14960.00', rate: '372.84', unit: 'tonne' }
    ])
  })

  it('matches a unit glued directly onto the Per number with no space (a common OCR artifact)', async () => {
    const { extractEstimateItemsFromLines } = await import('../core/estimateExtract')
    const lines = ['Qty Rate Per Amount', 'Earthwork excavation', '619.75 475.5901 Cum 294746.90']
    const items = extractEstimateItemsFromLines(lines)
    expect(items).toMatchObject([{ quantity: '619.75', rate: '475.59', unit: 'Cum' }])
  })

  it('does not let a spec line starting with digits ("600 mm dia...") wipe out the description already accumulated', async () => {
    const { extractEstimateItemsFromLines } = await import('../core/estimateExtract')
    const lines = [
      'Qty Rate Per Amount',
      'MANUFACTURE, SUPPLY AND DELIVERY OF 600mm DIA R.C.C',
      '600 mm dia RCC NP3 Pipes',
      '550.00 3712.24 Rmt 2041732.00'
    ]
    const items = extractEstimateItemsFromLines(lines)
    expect(items).toHaveLength(1)
    expect(items[0].description).toContain('MANUFACTURE, SUPPLY AND DELIVERY OF 600mm DIA R.C.C')
    expect(items[0]).toMatchObject({ quantity: '550.00', rate: '3712.24', unit: 'Rmt' })
  })

  it('extracts all 6 items from a realistic multi-item estimate, matching every Qty/Rate/Unit exactly', async () => {
    const { extractEstimateItemsFromLines } = await import('../core/estimateExtract')
    // A synthetic fixture, shaped after (but not copied from) the OCR
    // artifacts a real line-detecting OCR engine produces on a photographed
    // "Detailed and Abstract Estimate": glued-together numbers with no
    // space ("550.000.900.20"), a unit token glued directly onto the "Per"
    // multiplier ("2345.001Cum"), stray extra digits/characters, and a
    // "600 mm dia..." spec line that starts with a small integer (the kind
    // of line that must NOT be mistaken for a new item's S.No). This
    // document shape — a long, wrapped description on its own line(s), with
    // narrower numeric columns on separate rows below — defeats position-
    // based column reconstruction almost every time, regardless of photo
    // quality or OCR engine's raw text accuracy.
    const lines = [
      'SAMPLE MUNICIPAL CORPORATION::SAMPLE DISTRICT',
      'DETAILED AND ABSTRACT-ESTIMATE',
      'Ne of the work: Laying of a sample pipeline from Point A to Point B in Ward No 1',
      'Sample Municipal Corporation under Municipal General Funds 2025-26',
      'Estimate Amount Rs. 10.00 lakhs',
      "Sl. Description ofwork No's L B D",
      'Qty Rate Per Amount',
      'No.',
      'Cutting open road surface as well as concrete upto 75 mm thick',
      'including stacking of excavated materials for pipe line trench work',
      'as directed by the departmental officers.',
      'for sample line 1x 1500.000.900.20 90.00155472 13916300',
      '90.00 1500.001Cum 135000.000',
      '2 Earthwork excavation in all kinds of soil for Pipeline trenches as',
      'per drawings and technical specifications including setting out',
      'construction of shoring and bracing, removal of stumps and other',
      'deleterious material, dressing of sides and bottom as per spec',
      'for sample line 1x 500.000.901.50 700.50',
      'Deduct Rock Qty 1x 1200.75 -200.75 2500 5000308',
      '500.75 400.5901 Cum 200000.00',
      'Earth work excavation in Hard rock blasting prohibited upto 3 m',
      'depth for foundations and depositing on bank for all lifts',
      'for sample line 500mm Dia 1x 30% 700.50 200.75',
      '200.75 2000.80l1Cum 400000.00',
      'MANUFACTURE, SUPPLY AND DELIVERY OF 500mm DIA R.C.C',
      'SOCKET AND SPIGOT PIPES CONFORMING TO STANDARD SPEC',
      '500 mm dia RCCNP3 Pipes X 500.00 500.00 Hl00.85 250000100',
      '500.00 3000.24 Rmt 1500000.00',
      'Lowering the RCC pipes carefully into the trenches laying them',
      'true to alignment and gradient, jointing with rubber rings',
      '500 mm dia RCC NP3 Pipes X 500.00 500.00 250.06',
      '500.00 250.75 Rmt 125000.00',
      'Supply and Fixing of Rubber Rings',
      '500mm dia RCC X 200.00 200.00',
      '200.00 450.18 Nos 90000.00'
    ]

    const items = extractEstimateItemsFromLines(lines)
    expect(items).toHaveLength(6)
    expect(items.map((i) => ({ quantity: i.quantity, rate: i.rate, unit: i.unit }))).toEqual([
      { quantity: '90.00', rate: '1500.00', unit: 'Cum' },
      { quantity: '500.75', rate: '400.59', unit: 'Cum' },
      { quantity: '200.75', rate: '2000.80', unit: 'Cum' },
      { quantity: '500.00', rate: '3000.24', unit: 'Rmt' },
      { quantity: '500.00', rate: '250.75', unit: 'Rmt' },
      { quantity: '200.00', rate: '450.18', unit: 'Nos' }
    ])
  })

  it('corrects a misread quantity from the printed Amount ÷ Rate (Qty=Amount/Rate)', async () => {
    const { extractEstimateItemsFromLines } = await import('../core/estimateExtract')
    // OCR misread the quantity as 30.00, but Rate 1500.00 and Amount 135000.00
    // are right: 30 × 1500 = 45000 ≠ 135000, so the quantity is corrected back
    // to 135000 / 1500 = 90.00.
    const lines = ['Qty Rate Per Amount', 'Earthwork excavation', '30.00 1500.00 Cum 135000.00']
    const items = extractEstimateItemsFromLines(lines)
    expect(items).toMatchObject([{ quantity: '90.00', rate: '1500.00', unit: 'Cum' }])
  })

  it('picks Qty/Rate by the Qty×Rate=Amount invariant when an extra number sits on the summary line', async () => {
    const { extractEstimateItemsFromLines } = await import('../core/estimateExtract')
    // OCR left a stray extra decimal after the amount, so the printed "last
    // three" (Amount, extra) no longer start on Qty — the product check must
    // still land Qty=90.00, Rate=1500.00 (90 × 1500 = 135000), not the extra.
    const lines = ['Qty Rate Per Amount', 'Earthwork excavation', '90.00 1500.00 135000.00 5.00 Cum']
    const items = extractEstimateItemsFromLines(lines)
    expect(items).toMatchObject([{ quantity: '90.00', rate: '1500.00', unit: 'Cum' }])
  })

  it('captures a No\'s/L/B/D dimension line sitting between the description and the summary line', async () => {
    const { extractEstimateItemsFromLines } = await import('../core/estimateExtract')
    const lines = [
      'Qty Rate Per Amount',
      'Cutting open CC road surface for pipe line trench work',
      '5.00 2.50 1.20 0.75',
      '99.00 2345.00 Cum 232155.00'
    ]
    const items = extractEstimateItemsFromLines(lines)
    expect(items).toMatchObject([
      { quantity: '99.00', rate: '2345.00', unit: 'Cum', nos: '5.00', l: '2.50', b: '1.20', d: '0.75' }
    ])
  })

  it('reads a dimension line whose No\'s is a bare integer and whose depth has 3 decimals', async () => {
    const { extractEstimateItemsFromLines } = await import('../core/estimateExtract')
    const lines = [
      'Qty Rate Per Amount',
      'Laying of CC road',
      '1 250.00 6.00 0.075',
      '112.50 5800.00 Cum 652500.00'
    ]
    const items = extractEstimateItemsFromLines(lines)
    expect(items).toMatchObject([{ nos: '1', l: '250.00', b: '6.00', d: '0.075' }])
  })

  it('reads a count + L + B dimension line (area item, no depth)', async () => {
    const { extractEstimateItemsFromLines } = await import('../core/estimateExtract')
    const lines = ['Qty Rate Per Amount', 'BT road renewal', '2 100.00 7.50', '1500.00 320.00 Sqm 480000.00']
    const items = extractEstimateItemsFromLines(lines)
    expect(items).toMatchObject([{ nos: '2', l: '100.00', b: '7.50' }])
  })

  it('treats a 3-number dimension line as L/B/D with an implicit single count', async () => {
    const { extractEstimateItemsFromLines } = await import('../core/estimateExtract')
    const lines = ['Qty Rate Per Amount', 'Plain Cement concrete', '10.00 2.00 0.30', '6.00 6587.04 Cum 43424.00']
    const items = extractEstimateItemsFromLines(lines)
    expect(items).toMatchObject([{ l: '10.00', b: '2.00', d: '0.30' }])
    expect(items[0].nos).toBeUndefined()
  })

  it('does not mistake a dimension line for a description, and does not carry stale dims into the next item', async () => {
    const { extractEstimateItemsFromLines } = await import('../core/estimateExtract')
    const lines = [
      'Qty Rate Per Amount',
      'Item one with dimensions',
      '2.00 3.00 4.00 5.00',
      '99.00 100.00 Cum 9900.00',
      'Item two with no dimensions shown',
      '50.00 200.00 Cum 10000.00'
    ]
    const items = extractEstimateItemsFromLines(lines)
    expect(items).toHaveLength(2)
    expect(items[0].description).toBe('Item one with dimensions')
    expect(items[0]).toMatchObject({ nos: '2.00', l: '3.00', b: '4.00', d: '5.00' })
    expect(items[1].nos).toBeUndefined()
    expect(items[1].l).toBeUndefined()
    expect(items[1].b).toBeUndefined()
    expect(items[1].d).toBeUndefined()
  })

  it('reconstructs No\'s/L/B/D glued together with no delimiter, alongside real words on the same line, by checking candidate splits against the already-known Qty', async () => {
    const { extractEstimateItemsFromLines } = await import('../core/estimateExtract')
    // Shaped after a real line-detecting OCR engine's output on a photographed
    // estimate: the printed dimension row ("1  550.00  0.90  0.20") comes back
    // glued into one blob with no spaces at all, mixed in with a label phrase
    // and further glued noise (a repeated Qty and an Amount figure) that must
    // NOT be mistaken for real dimension data.
    const lines = [
      'Qty Rate Per Amount',
      'Cutting open road surface for pipe line trench work as directed by the departmental officers.',
      'for UGD line X 1550.000.900.20 99.0026642 2616300',
      '99.00 2345.00 1Cum 232155.00'
    ]
    const items = extractEstimateItemsFromLines(lines)
    expect(items).toMatchObject([
      { quantity: '99.00', rate: '2345.00', unit: 'Cum', nos: '1', l: '550.00', b: '0.90', d: '0.20' }
    ])
  })

  it('leaves No\'s/L/B/D blank when no candidate split\'s product matches Qty, rather than guessing', async () => {
    const { extractEstimateItemsFromLines } = await import('../core/estimateExtract')
    // A pipe-length item: "550.00" appears twice (matching Qty exactly on its
    // own), but a single matching number is exactly the case that must be
    // rejected — with only one piece there's no way to tell it apart from Qty
    // simply being printed again on the same line.
    const lines = [
      'Qty Rate Per Amount',
      'MANUFACTURE, SUPPLY AND DELIVERY OF 600mm DIA R.C.C pipes',
      '600 mm dia RCC NP3 Pipes 550.00 550.00',
      '550.00 3712.24 Rmt 2041732.00'
    ]
    const items = extractEstimateItemsFromLines(lines)
    expect(items).toHaveLength(1)
    expect(items[0].nos).toBeUndefined()
    expect(items[0].l).toBeUndefined()
  })

  it('does not reconstruct dimensions across a Qty-changing deduction line — leaves it blank rather than match against the gross, pre-deduction figure', async () => {
    const { extractEstimateItemsFromLines } = await import('../core/estimateExtract')
    // The dimension line's own arithmetic (550.00 x 0.90 x 1.50 = 742.50) is
    // real, but a "Deduct Rock Qty" line nets it down to the final Qty
    // (519.75) printed on the summary line — the reconstruction must not
    // match against the pre-deduction 742.50, since that's not the item's
    // actual final quantity.
    const lines = [
      'Qty Rate Per Amount',
      'Earthwork excavation in all kinds of soil for Pipeline trenches',
      'for UGD line X 550.00 0.901.50 742.50',
      'Deduct Rock Qty 222.75 -222.75',
      '519.75 475.59 Cum 247188.00'
    ]
    const items = extractEstimateItemsFromLines(lines)
    expect(items).toHaveLength(1)
    expect(items[0]).toMatchObject({ quantity: '519.75', rate: '475.59' })
    expect(items[0].l).toBeUndefined()
  })

  it('recognizes the header even when Qty and Rate land on two separate OCR lines instead of one', async () => {
    const { extractEstimateItemsFromLines } = await import('../core/estimateExtract')
    // A tall/wrapped header row can get split across two detected lines —
    // e.g. "SI. Rate Per Amount" then "Description of work No's L B D Qty" —
    // so neither line alone has both "qty" and "rate".
    const lines = [
      'CYBERABAD MUNICIPAL CORPORATION : QUTHBULLAPUR ZONE',
      'SI. Rate Per Amount',
      "Description of work No's L B D Qty",
      'Earthwork excavation for road way in soil by mechanical means',
      '180.00 73.08 Cum 13154.00'
    ]
    const items = extractEstimateItemsFromLines(lines)
    expect(items).toMatchObject([{ quantity: '180.00', rate: '73.08', unit: 'Cum' }])
  })

  it('reconstructs No\'s/L/B/D when only some of them are glued together and the rest are already clean, separate tokens on the same line', async () => {
    const { extractEstimateItemsFromLines } = await import('../core/estimateExtract')
    // "1100.006.00" is No's=1/L=100.00/B=6.00 glued into one blob; "0.30" (D)
    // is already its own clean, separate token right after it.
    const lines = [
      'Qty Rate Per Amount',
      'Earthwork excavation for road way in soil by mechanical means',
      '1100.006.00 0.30 180.00',
      '180.00 73.08 Cum 13154.00'
    ]
    const items = extractEstimateItemsFromLines(lines)
    expect(items).toMatchObject([
      { quantity: '180.00', rate: '73.08', unit: 'Cum', nos: '1', l: '100.00', b: '6.00', d: '0.30' }
    ])
  })

  it('accepts a 3-decimal dimension value (e.g. a 0.075m slab thickness) on its own clean token, not just the usual 2-decimal convention', async () => {
    const { extractEstimateItemsFromLines } = await import('../core/estimateExtract')
    const lines = [
      'Qty Rate Per Amount',
      'Supply and placing of the Ready Mix Concrete M10 grade',
      '100.005.50 0.075 41.25',
      '41.25 4445.90 Cum 183393.00'
    ]
    const items = extractEstimateItemsFromLines(lines)
    expect(items).toMatchObject([{ quantity: '41.25', rate: '4445.90', l: '100.00', b: '5.50', d: '0.075' }])
  })

  it('reconstructs a 2-value L x B dimension (no depth) for an area-only item, without forcing a spurious D', async () => {
    const { extractEstimateItemsFromLines } = await import('../core/estimateExtract')
    const lines = [
      'Qty Rate Per Amount',
      'Supply and providing of Seperation membrane including cost and conveyance',
      '100.005.50 550.00',
      '550.00 14.00 Cum 7700.00'
    ]
    const items = extractEstimateItemsFromLines(lines)
    expect(items).toMatchObject([{ quantity: '550.00', l: '100.00', b: '5.50' }])
    expect(items[0].d).toBeUndefined()
  })
})

describe('extractWorkName', () => {
  it('reads the name inline after a colon in the same cell', async () => {
    const { extractWorkName } = await import('../core/estimateExtract')
    const grid = [
      ['Detailed Estimate'],
      ['Name of Work: Road from A to B'],
      ['S.No', 'Description', 'Qty', 'Rate', 'Unit']
    ]
    expect(extractWorkName(grid, 2)).toBe('Road from A to B')
  })

  it('reads the name from the next non-empty cell when the label is alone', async () => {
    const { extractWorkName } = await import('../core/estimateExtract')
    const grid = [
      ['Name of the Work', '', 'Road from A to B'],
      ['S.No', 'Description', 'Qty', 'Rate', 'Unit']
    ]
    expect(extractWorkName(grid, 1)).toBe('Road from A to B')
  })

  it('only looks above the header row, and returns undefined when no label is found', async () => {
    const { extractWorkName } = await import('../core/estimateExtract')
    const grid = [
      ['Some other title'],
      ['S.No', 'Description', 'Qty', 'Rate', 'Unit'],
      ['1', 'Name of Work: should not be read from a data row', '1', '1', 'Cum']
    ]
    expect(extractWorkName(grid, 1)).toBeUndefined()
  })
})

describe('Schedule A meta from the Works List', () => {
  const worksTable: ExcelTable = {
    id: 'works',
    name: 'Works database',
    path: '',
    headers: [
      'Name of the work',
      'Amount of estimate',
      'ECV',
      'Contract Amount',
      'Name of the Agency',
      'Tender Percentage'
    ],
    rows: [
      {
        'Name of the work': 'Road from A to B',
        'Amount of estimate': '45',
        'ECV': '4200000',
        'Contract Amount': '40',
        'Name of the Agency': 'ABC Constructions',
        'Tender Percentage': '18'
      }
    ]
  }

  it('findWorksRowByName matches case- and whitespace-insensitively', async () => {
    const { findWorksRowByName } = await import('../core/scheduleA')
    const match = findWorksRowByName(worksTable, '  road   FROM a to b ')
    expect(match?.row).toBe(worksTable.rows[0])
    expect(match?.matchedViaAi).toBe(false)
    expect(findWorksRowByName(worksTable, 'Some unrelated work')).toBeUndefined()
  })

  it('findWorksRowByName falls back to the closest embedding match when the exact name differs', async () => {
    const { findWorksRowByName } = await import('../core/scheduleA')
    // "Road from A to B" (row 0) vs a differently-worded query — no exact
    // match, so this only resolves via the supplied embeddings.
    const match = findWorksRowByName(worksTable, 'Road works between A and B, Ph-1', {
      workNameVector: [1, 0],
      rowNameVectors: [
        [0.99, 0.01], // row 0 — closest
        [0, 1] // row 1 — unrelated
      ]
    })
    expect(match?.row).toBe(worksTable.rows[0])
    expect(match?.matchedViaAi).toBe(true)
  })

  it('findWorksRowByName ignores an embedding match below the threshold', async () => {
    const { findWorksRowByName } = await import('../core/scheduleA')
    const match = findWorksRowByName(worksTable, 'Completely unrelated text', {
      workNameVector: [1, 0],
      rowNameVectors: [
        [0.1, 0.99],
        [0.2, 0.98]
      ]
    })
    expect(match).toBeUndefined()
  })

  it('metaFromWorksRow treats "Name of the Agency" as "Name of the Contractor", and Indian-formats the amounts', async () => {
    const { metaFromWorksRow } = await import('../core/scheduleA')
    const meta = metaFromWorksRow(worksTable.rows[0])
    expect(meta).toEqual({
      nameOfWork: 'Road from A to B',
      // 45/42 Lakhs -> rupees, Indian-grouped + "/-" (no "Rs" — Schedule A's own labels already print "Rs.").
      estimateAmount: '45,00,000/-',
      ecvAmount: '42,00,000/-',
      // Contract Amount is always computed — ECV * (1 - Tender Percentage) = 42,00,000 * (1 - 0.18) = 34,44,000.
      contractAmount: '34,44,000/-',
      contractorName: 'ABC Constructions',
      tenderPercentage: '18'
    })
  })
})
