import { describe, expect, it } from 'vitest'
import { reconstructGrid } from '../core/ocrTableReconstruct'
import type { OcrWord } from '../core/ocrTableReconstruct'

function word(text: string, x0: number, y0: number, w = 40, h = 20): OcrWord {
  return { text, bbox: { x0, y0, x1: x0 + w, y1: y0 + h } }
}

describe('reconstructGrid', () => {
  it('returns an empty grid for no words', () => {
    expect(reconstructGrid([])).toEqual([])
  })

  it('groups a simple two-row, four-column table correctly', () => {
    const words: OcrWord[] = [
      // Header row (y ~ 20)
      word('Sl', 10, 20),
      word('No', 60, 20),
      word('Description', 200, 20, 120),
      word('Qty', 400, 20),
      word('Rate', 500, 20),
      // Data row (y ~ 70, same column bands)
      word('1', 10, 70),
      word('Excavation', 200, 70, 120),
      word('25.50', 400, 70),
      word('350.00', 500, 70)
    ]
    const grid = reconstructGrid(words)
    expect(grid).toHaveLength(2)
    // "Sl" and "No" are close enough in x to fall in the same column band
    // (small gap), landing joined in one cell — that's fine, it's still
    // recognizable as the serial-number column.
    expect(grid[0]).toEqual(['Sl No', 'Description', 'Qty', 'Rate'])
    expect(grid[1]).toEqual(['1', 'Excavation', '25.50', '350.00'])
  })

  it('joins multiple words in the same row and column (a wrapped description)', () => {
    const words: OcrWord[] = [
      word('1', 10, 20),
      word('Excavation', 100, 20, 80),
      word('in', 185, 20),
      word('earth', 230, 20),
      word('cum', 400, 20)
    ]
    const grid = reconstructGrid(words)
    expect(grid).toHaveLength(1)
    // "Excavation", "in", "earth" are all close together (small gaps) so
    // they land in the same description column, joined with spaces.
    expect(grid[0][1]).toBe('Excavation in earth')
  })

  it('keeps rows in top-to-bottom order even if words are given out of order', () => {
    const words: OcrWord[] = [word('second', 10, 100), word('first', 10, 20), word('third', 10, 180)]
    const grid = reconstructGrid(words)
    expect(grid.map((r) => r[0])).toEqual(['first', 'second', 'third'])
  })

  it('treats words on the same visual line (small y differences) as one row', () => {
    // Real OCR word boxes on the same printed line are rarely at the exact
    // same y0 — a few pixels of jitter is normal and shouldn't split a row.
    const words: OcrWord[] = [word('A', 10, 20), word('B', 200, 24), word('C', 400, 18)]
    const grid = reconstructGrid(words)
    expect(grid).toHaveLength(1)
    expect(grid[0]).toEqual(['A', 'B', 'C'])
  })

  it('separates genuinely different rows even when close in y', () => {
    const words: OcrWord[] = [word('row1', 10, 20, 40, 20), word('row2', 10, 200, 40, 20)]
    const grid = reconstructGrid(words)
    expect(grid).toHaveLength(2)
  })
})
