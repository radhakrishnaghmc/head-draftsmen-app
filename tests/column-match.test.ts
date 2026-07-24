import { describe, expect, it } from 'vitest'
import { resolveColumns } from '../core/columnMatch'
import type { ColumnSpec, ColumnEmbeddings } from '../core/columnMatch'

const QTY_RATE_UNIT_SPECS: ColumnSpec[] = [
  { label: 'Quantity', patterns: [/qty|quantity/i] },
  { label: 'Rate', patterns: [/rate/i] },
  { label: 'Unit', patterns: [/unit|^per$/i] }
]

describe('resolveColumns', () => {
  it('resolves every spec by regex alone, no embeddings needed', () => {
    const headers = ['S.No', 'Description', 'Qty', 'Rate', 'Unit']
    const { indexByLabel, viaEmbedding } = resolveColumns(headers, QTY_RATE_UNIT_SPECS)
    expect(indexByLabel.Quantity).toBe(2)
    expect(indexByLabel.Rate).toBe(3)
    expect(indexByLabel.Unit).toBe(4)
    expect(viaEmbedding).toEqual([])
  })

  it('throws naming the first unresolvable spec when there are no embeddings to fall back on', () => {
    const headers = ['S.No', 'Description', 'Rate', 'Unit'] // no quantity column at all
    expect(() => resolveColumns(headers, QTY_RATE_UNIT_SPECS)).toThrow(/"Quantity"/)
  })

  it('falls back to embedding similarity for a column no regex anticipated', () => {
    // "Nos." matches none of the Quantity/Rate/Unit regexes at all.
    const headers = ['S.No', 'Description', 'Nos.', 'Rate', 'Unit']
    const embeddings: ColumnEmbeddings = {
      // One vector per header (S.No, Description, Nos., Rate, Unit).
      headerVectors: [
        [0, 0, 1],
        [0, 1, 0],
        [1, 0, 0], // "Nos." — closest to the Quantity label vector below
        [0.9, 0.1, 0],
        [0.1, 0.9, 0]
      ],
      // One vector per spec (Quantity, Rate, Unit), matching QTY_RATE_UNIT_SPECS order.
      labelVectors: [
        [1, 0, 0], // Quantity
        [0.9, 0.1, 0], // Rate
        [0.1, 0.9, 0] // Unit
      ]
    }
    const { indexByLabel, viaEmbedding } = resolveColumns(headers, QTY_RATE_UNIT_SPECS, embeddings)
    expect(indexByLabel.Quantity).toBe(2) // "Nos."
    expect(indexByLabel.Rate).toBe(3) // resolved by regex
    expect(indexByLabel.Unit).toBe(4) // resolved by regex
    expect(viaEmbedding).toEqual(['Quantity']) // only Quantity needed the embedding fallback
  })

  it('never resolves two different specs to the same header via the embedding fallback', () => {
    const headers = ['Nos.', 'RateCol']
    const specs: ColumnSpec[] = [
      { label: 'Quantity', patterns: [/^nomatch$/] },
      { label: 'Rate', patterns: [/^nomatch$/] }
    ]
    // Both specs' own top preference is header 0 — Rate must fall back to its
    // second-best (header 1, still above the confidence threshold) instead of
    // reusing the header Quantity already claimed.
    const embeddings: ColumnEmbeddings = {
      headerVectors: [
        [1, 0],
        [0.8, 0.6]
      ],
      labelVectors: [
        [1, 0.05], // Quantity: header0 ≈0.999, header1 ≈0.829
        [0.95, 0.15] // Rate: header0 ≈0.988, header1 ≈0.884
      ]
    }
    const { indexByLabel } = resolveColumns(headers, specs, embeddings)
    expect(indexByLabel.Quantity).toBe(0)
    expect(indexByLabel.Rate).toBe(1)
  })

  it('throws when even the embedding fallback scores below the confidence threshold', () => {
    const headers = ['S.No', 'Description']
    const specs: ColumnSpec[] = [{ label: 'Quantity', patterns: [/^nomatch$/] }]
    const embeddings: ColumnEmbeddings = {
      headerVectors: [
        [1, 0],
        [0, 1]
      ],
      labelVectors: [[0, -1]] // orthogonal/opposed to both headers — no real match
    }
    expect(() => resolveColumns(headers, specs, embeddings)).toThrow(/"Quantity"/)
  })
})
