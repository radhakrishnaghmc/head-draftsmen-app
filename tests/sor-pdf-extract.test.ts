import { describe, it, expect } from 'vitest'
import { extractSorRateEntries } from '../core/sorPdfExtract'

describe('extractSorRateEntries', () => {
  it('extracts lettered sub-items sharing one group heading, folding a wrapped description and a rate that lands on its own line', () => {
    const lines = [
      'Old Material Rates for Electrical items',
      'Buildings (Part-III) SoR 2026-27',
      'Page 396 of 806',
      'Sl.No Item code Description of item Unit Scheduled Rate',
      'ELEC- 1.9.1 Supply -------------------------------------- Makes : Legrand Arteor / Schneider Zencelo',
      '377 a 6 / 10A 1 Way 1 Module Modular Switch',
      '1 No 144.00',
      '379 c 6 / 10A 2 Pin 2 Module Modular Socket.',
      '1 No 144.00',
      '380 d 6 / 10A 3 Pin 2 Module Modular Socket',
      'with shutter.',
      '1 No 215.00'
    ]
    const entries = extractSorRateEntries(lines, 'SoR 2026-27')
    expect(entries).toHaveLength(3)
    expect(entries[2]).toMatchObject({
      description: 'Supply - 6 / 10A 3 Pin 2 Module Modular Socket with shutter.',
      rate: '215.00',
      unit: 'each',
      code: 'ELEC-1.9.1(d)',
      sheet: 'SoR 2026-27'
    })
  })

  it('extracts a self-contained numbered row that carries its own code (no group heading)', () => {
    const lines = [
      'Plumbing Items',
      'Builings(Part-III) SoR 2026-27',
      'Page 353 of 806',
      'Sl. No Item Code Description of the item Unit Scheduled Rate',
      '409 TBSP-G.I-01 S& F 15 mm Nominal Bore GI pipe Light Grade properties & weight',
      'as per IS 1239 ISI mark with GI fittings including the cost of pipe & its fittings',
      '&labour charges complete',
      '1 RM 211.00'
    ]
    const entries = extractSorRateEntries(lines, 'SoR 2026-27')
    expect(entries).toHaveLength(1)
    expect(entries[0]).toMatchObject({
      description:
        'S& F 15 mm Nominal Bore GI pipe Light Grade properties & weight as per IS 1239 ISI mark with GI fittings including the cost of pipe & its fittings &labour charges complete',
      rate: '211.00',
      unit: 'rmt',
      code: 'TBSP-G.I-01'
    })
  })

  it('drops an unfinished row that never reaches a rate tail before the next serial number starts', () => {
    const lines = [
      '1 a Some truncated row whose rate column got clipped by a page break',
      '2 b A following, well-formed row',
      '1 No 95.00'
    ]
    const entries = extractSorRateEntries(lines)
    expect(entries).toHaveLength(1)
    expect(entries[0].rate).toBe('95.00')
  })

  it('ignores repeated page furniture (headers/footers) without corrupting an open row buffer', () => {
    const lines = [
      'ELEC- 5.15.6 Supply, Transportation and fixing of 450mm length, 12mm thickness and',
      'Old Material Rates for Electrical items',
      'Page 461 of 806',
      '75mm wide MS Clamp set with required nuts and bolts.',
      '1354 a 40 mm dia',
      '1 No 431.00'
    ]
    const entries = extractSorRateEntries(lines)
    expect(entries).toHaveLength(1)
    expect(entries[0]).toMatchObject({
      description: 'Supply, Transportation and fixing of 450mm length, 12mm thickness and 75mm wide MS Clamp set with required nuts and bolts. - 40 mm dia',
      rate: '431.00',
      code: 'ELEC-5.15.6(a)'
    })
  })

  it('yields nothing for lines with no recognizable row/code convention', () => {
    expect(extractSorRateEntries(['SECTION - 5', '5.1 CEILING AND EXHAUST FANS', 'Index'])).toEqual([])
  })
})
