import { describe, expect, it } from 'vitest'
import {
  digitConfusionVariants,
  formatMbFigure,
  mbAverageVariants,
  mbDimensionCandidates,
  mbFigureCandidates,
  mbMappedDigitCount,
  normalizeMbFigure,
  parseMbCount,
  readMbCount,
  readMbFigure,
  reconcileMbRow,
  splitAverageNumerator
} from '../core/mbNumeric'

describe('normalizeMbFigure', () => {
  it('maps letter shapes back to the digits they were written as', () => {
    expect(normalizeMbFigure('IS.00')).toBe('15.00')
    expect(normalizeMbFigure('O.3O')).toBe('0.30')
  })

  it('maps every decimal-point shape the model returns', () => {
    expect(normalizeMbFigure('6-30')).toBe('6.30')
    expect(normalizeMbFigure('6·30')).toBe('6.30')
    expect(normalizeMbFigure('6,30')).toBe('6.30')
  })

  it('drops unit words and ditto ticks around the figure', () => {
    expect(normalizeMbFigure('31.05um')).toBe('31.05')
  })
})

describe('mbFigureCandidates', () => {
  it('offers the reading as given first when it already has 2 decimals', () => {
    expect(mbFigureCandidates('28.35')[0]).toBe('28.35')
  })

  it('trims a third decimal, which is the clerk unit tick read as a digit', () => {
    expect(mbFigureCandidates('29.704')[0]).toBe('29.7')
  })

  it('puts a lost decimal point back 2 from the right', () => {
    expect(mbFigureCandidates('28351')[0]).toBe('283.51')
    expect(mbFigureCandidates('28351')).toContain('28.35')
  })

  it('returns nothing for a cell holding no real digit', () => {
    // Description text leaning into a figure column, or invention from paper grain.
    expect(mbFigureCandidates('dorprisn')).toEqual([])
    expect(readMbFigure('pwell')).toBe('')
  })
})

describe('splitAverageNumerator', () => {
  it('splits on the plus signs when they survived, keeping each part 2-decimal', () => {
    expect(splitAverageNumerator('6.40+6.50', 2)).toEqual(['6.4', '6.5'])
    expect(splitAverageNumerator('640+6.50', 2)).toEqual(['6.4', '6.5'])
  })

  it('splits structurally when the plus signs were lost or read as digits', () => {
    expect(splitAverageNumerator('630630', 2)).toEqual(['6.3', '6.3'])
    expect(splitAverageNumerator('6304630', 2)).toEqual(['6.3', '6.3'])
  })

  it('refuses a digit run that cannot hold the stated number of figures', () => {
    expect(splitAverageNumerator('630', 2)).toBeUndefined()
  })
})

describe('mbDimensionCandidates', () => {
  it('reads a cell written as an average over a denominator', () => {
    expect(mbDimensionCandidates('6.70+6.50\n2')[0]).toBe('6.6')
  })

  it('reads a cell as an average even when the fraction bar and denominator were lost', () => {
    expect(mbDimensionCandidates('6.20+6.50')[0]).toBe('6.35')
  })
})

describe('digit repair candidates', () => {
  it('swaps one digit for a shape the model confuses it with', () => {
    expect(digitConfusionVariants(['6.30'])).toContain('0.3')
  })

  it("swaps every position at once, for a clerk's zero read as a six throughout", () => {
    expect(digitConfusionVariants(['15.66'])).toContain('15')
  })

  it('repairs an averaged cell on its parts, not on the mean', () => {
    // 6.20 was written 6.70; the average moves 6.35 -> 6.60.
    expect(mbAverageVariants(['6.20', '6.50'], 2)).toContain('6.6')
  })
})

describe('reconcileMbRow', () => {
  const row = (over: Partial<Parameters<typeof reconcileMbRow>[0]>) =>
    reconcileMbRow({ count: 1, l: [], b: [], d: [], contents: [], ...over })

  it('reports OK when the best reading of every cell multiplies out', () => {
    const result = row({ l: ['15.00'], b: ['6.30'], d: ['0.30'], contents: ['28.35'] })
    expect(result.check).toBe('OK')
    expect(result.computed).toBe('28.35')
  })

  it('picks the weaker reading of a cell when only that one satisfies the row', () => {
    const result = row({ l: ['15.00'], b: ['6.30'], d: ['0.30'], contents: ['283.51', '28.35'] })
    expect(result.contents).toBe('28.35')
    expect(result.check).toBe('Corrected')
  })

  it('reports Recheck without rewriting anything when nothing reconciles', () => {
    const result = row({ l: ['15.00'], b: ['6.30'], d: ['0.30'], contents: ['29.35'] })
    expect(result.check).toBe('Recheck')
    expect(result.b).toBe('6.30')
    expect(result.contents).toBe('29.35')
  })

  it('accepts a digit repair in one cell when the row then multiplies out', () => {
    const result = row({
      l: ['15.00'],
      b: ['6.30'],
      d: ['6.30'],
      contents: ['28.35'],
      alternates: { l: [], b: [], d: ['0.30'] }
    })
    expect(result.d).toBe('0.30')
    expect(result.check).toBe('Corrected')
  })

  it('will not repair two cells on an approximate fit, which is arithmetic coincidence', () => {
    // 15.00 x 6.36 x 0.30 = 28.62, two digits away from what was read and 0.02
    // off the Contents figure: close, but not the row the clerk wrote.
    const result = row({
      l: ['15.00'],
      b: ['6.30'],
      d: ['6.30'],
      contents: ['28.60'],
      alternates: { l: [], b: ['6.36'], d: ['0.30'] }
    })
    expect(result.check).toBe('Recheck')
  })

  it('repairs two cells when the row then multiplies out exactly', () => {
    // Both readings are one digit shape away from the page, and only these two
    // together reproduce Contents to the book's own rounding.
    const result = row({
      l: ['15.00'],
      b: ['6.30'],
      d: ['6.30'],
      contents: ['28.58'],
      alternates: { l: [], b: ['6.35'], d: ['0.30'] }
    })
    expect(result.b).toBe('6.35')
    expect(result.d).toBe('0.30')
    expect(result.check).toBe('Corrected')
  })

  it('solves a blank cell from the rest of the row', () => {
    const result = row({ l: ['15.00'], b: [], d: ['0.30'], contents: ['29.03'] })
    expect(result.b).toBe('6.45')
    expect(result.check).toBe('Corrected')
  })

  it('does not solve from a blank cell, whose missing figure is not a 1', () => {
    const result = row({ l: ['15.00'], b: [], d: [], contents: ['29.03'] })
    expect(result.check).toBe('Recheck')
  })

  it('does not solve a cell that was read, however badly the row adds up', () => {
    const result = row({ l: ['15.00'], b: ['2.06'], d: ['0.30'], contents: ['31.05'] })
    expect(result.b).toBe('2.06')
    expect(result.check).toBe('Recheck')
  })

  it('solves a cell whose ink the caller flagged as unreadable', () => {
    const result = row({
      l: ['15.00'],
      b: ['6.22'],
      d: ['0.30'],
      contents: ['30.15'],
      trusted: { l: true, b: false, d: true }
    })
    expect(result.b).toBe('6.7')
    expect(result.check).toBe('Corrected')
  })

  it('multiplies by the No. column', () => {
    const result = row({ count: 2, l: ['15.00'], b: ['6.30'], d: ['0.30'], contents: ['56.70'] })
    expect(result.check).toBe('OK')
  })
})

describe('count and formatting', () => {
  it('reads the No. column as a multiplier', () => {
    expect(readMbCount('1x1')).toBe('1x1')
    expect(parseMbCount('2x3')).toBe(6)
    expect(parseMbCount('')).toBe(1)
  })

  it('treats a No. cell with no digit as noise, not as a count', () => {
    expect(readMbCount('pwell')).toBe('')
  })

  it('writes figures to the 2 decimals the book is kept in', () => {
    expect(formatMbFigure('6.7')).toBe('6.70')
    expect(formatMbFigure('')).toBe('')
  })

  it('counts letter-shaped digits when measuring how much a cell holds', () => {
    expect(mbMappedDigitCount('IS.00')).toBe(4)
    expect(mbMappedDigitCount('6046-40\n2846')).toBe(10)
  })
})
