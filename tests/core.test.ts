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

  it('matches a unit glued directly onto the Per number with no space (a common OCR artifact)', async () => {
    const { extractEstimateItemsFromLines } = await import('../core/estimateExtract')
    const lines = ['Qty Rate Per Amount', 'Earthwork excavation', '619.75 475.5901 Cum 247188.00']
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
      'Estimate Amount ECV',
      'Contract Amount',
      'Name of the Agency',
      'Tender Percentage'
    ],
    rows: [
      {
        'Name of the work': 'Road from A to B',
        'Amount of estimate': '45',
        'Estimate Amount ECV': '42',
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
