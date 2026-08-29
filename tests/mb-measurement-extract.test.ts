import { describe, expect, it } from 'vitest'
import {
  estimateRowPitch,
  findGridStart,
  generateRowBounds,
  splitMultiValue,
  normalizeNumericCell,
  buildMbMeasurementRows
} from '../core/mbMeasurementExtract'

describe('estimateRowPitch', () => {
  it('averages only the plausible row-height gaps, skipping title-block spacing', () => {
    // Real measured line-Ys from a sample MB page: 87,171 are title-block
    // rules (gaps 84,87), 258..599 are clean table rows (gaps ~48-49), then
    // a gap to 890 from undetected mid-table lines obscured by handwriting.
    const lineYs = [87, 171, 258, 307, 356, 404, 453, 502, 550, 599, 890]
    const pitch = estimateRowPitch(lineYs)
    expect(pitch).toBeCloseTo(48.71, 1)
  })

  it('returns undefined when no gap falls in the plausible range', () => {
    expect(estimateRowPitch([10, 500])).toBeUndefined()
  })
})

describe('findGridStart', () => {
  it('finds the first line whose gap to the next line matches the pitch', () => {
    const lineYs = [87, 171, 258, 307, 356]
    expect(findGridStart(lineYs, 48.7)).toBe(258)
  })

  it('returns undefined when nothing matches', () => {
    expect(findGridStart([87, 171], 48.7)).toBeUndefined()
  })
})

describe('generateRowBounds', () => {
  it('generates bounds arithmetically from a start and pitch up to the bottom limit', () => {
    const bounds = generateRowBounds(258, 48.71, 890)
    expect(bounds[0]).toBe(258)
    expect(bounds[bounds.length - 1]).toBeCloseTo(890, -1)
    expect(bounds.length).toBe(14) // 13 rows -> 14 boundaries
  })

  it('caps at maxRows even if bottomLimit is far away', () => {
    const bounds = generateRowBounds(0, 50, 100000, 5)
    expect(bounds.length).toBe(6)
  })
})

describe('splitMultiValue', () => {
  it('keeps a single value in the first slot only', () => {
    expect(splitMultiValue('105')).toEqual(['105', '', ''])
  })

  it('splits multiple stacked lengths into l1/l2/l3 style slots', () => {
    expect(splitMultiValue('1.2\n2.4\n3.6')).toEqual(['1.2', '2.4', '3.6'])
  })

  it('splits on commas/semicolons too', () => {
    expect(splitMultiValue('1.2, 2.4; 3.6')).toEqual(['1.2', '2.4', '3.6'])
  })

  it('ignores blank lines and caps at 3 values', () => {
    expect(splitMultiValue('1\n\n2\n3\n4')).toEqual(['1', '2', '3'])
  })

  it('returns blanks for empty input', () => {
    expect(splitMultiValue('')).toEqual(['', '', ''])
  })
})

describe('normalizeNumericCell', () => {
  it('fixes common OCR decimal-separator confusions', () => {
    expect(normalizeNumericCell('1:5')).toBe('1.5')
    expect(normalizeNumericCell('1、5')).toBe('1.5')
  })

  it('leaves plain numbers and multiplier notation alone', () => {
    expect(normalizeNumericCell('1x1')).toBe('1x1')
    expect(normalizeNumericCell(' 105.00 ')).toBe('105.00')
  })
})

describe('buildMbMeasurementRows', () => {
  it('maps a 7-column grid into typed rows with l/b/d split into sub-slots', () => {
    const grid = [
      ['16.03.2025', 'Road No 7', '1x1', '105', '-', '-', '105.00'],
      ['', 'Wall segment', '1', '1.2\n2.4', '0.5', '', '90.00']
    ]
    const rows = buildMbMeasurementRows(grid)
    expect(rows[0]).toEqual({
      date: '16.03.2025',
      description: 'Road No 7',
      no: '1x1',
      l1: '105',
      l2: '',
      l3: '',
      b1: '-',
      b2: '',
      b3: '',
      d1: '-',
      d2: '',
      d3: '',
      contents: '105.00'
    })
    expect(rows[1].l1).toBe('1.2')
    expect(rows[1].l2).toBe('2.4')
    expect(rows[1].b1).toBe('0.5')
  })
})
