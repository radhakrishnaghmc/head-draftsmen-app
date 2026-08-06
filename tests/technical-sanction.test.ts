import { describe, it, expect } from 'vitest'
import {
  matchEstimateItems,
  buildCellEdits,
  buildRateAnalysisSheet,
  isAutoResolved,
  needsReview,
  extractItemCode,
  extractGrade
} from '../core/technicalSanction'
import type { EstimateWorkItem } from '../core/estimateExtract'
import type { RateEntry } from '../core/rateDatabase'
import type { MatchCandidate } from '../core/technicalSanction'

function item(overrides: Partial<EstimateWorkItem>): EstimateWorkItem {
  return { description: '', quantity: '1', rate: '', unit: 'Cum', ...overrides }
}

function entry(description: string, rate: string, extra: Partial<RateEntry> = {}): RateEntry {
  return { description, rate, sheet: 'test', breakdown: [], ...extra }
}

function candidate(overrides: Partial<MatchCandidate>): MatchCandidate {
  return { description: 'x', rate: '1', breakdown: [], ...overrides }
}

describe('matchEstimateItems', () => {
  const rateIndex: RateEntry[] = [
    entry('Earthwork in excavation for structures as per drawing and technical specifications clause 305.1.', '308.31'),
    entry('Earthwork in excavation for structures as per drawing and technical specifications clause 305.1.', '141.22'),
    entry('Brick masonry work in cement mortar excluding pointing and plastering as per drawing.', '6720.10')
  ]

  it('auto-resolves a description with exactly one candidate rate', () => {
    const items = [item({ description: 'Brick masonry work in cement mortar excluding pointing and plastering as per drawing.' })]
    const [match] = matchEstimateItems(items, rateIndex)
    expect(match.candidates).toHaveLength(1)
    expect(match.candidates[0].rate).toBe('6720.10')
    expect(match.semanticOnly).toBe(false)
  })

  it('surfaces every candidate when the best-matching description has several rates', () => {
    const items = [item({ description: 'Earthwork in excavation for structures as per drawing and technical specifications clause 305.1.' })]
    const [match] = matchEstimateItems(items, rateIndex)
    expect(match.candidates.map((c) => c.rate).sort()).toEqual(['141.22', '308.31'])
  })

  it('leaves an unrelated description unresolved (no candidates)', () => {
    const items = [item({ description: 'Supply and installation of solar photovoltaic panels on rooftop.' })]
    const [match] = matchEstimateItems(items, rateIndex)
    expect(match.candidates).toEqual([])
  })

  it('resolves via embeddings when TF-IDF alone would score below threshold (paraphrase with no shared vocabulary)', () => {
    // A completely different wording for the same underlying item — TF-IDF
    // (pure word-overlap) sees no shared vocabulary at all and would score 0,
    // but a semantic embedding correctly recognizes it as the same concept.
    const items = [item({ description: 'Building a wall out of bricks bonded together, minus rendering.' })]
    const brickEntryIndex = rateIndex.findIndex((e) => e.description.startsWith('Brick masonry'))
    // Synthetic vectors: the paraphrase's vector is deliberately made close
    // to the brick-masonry entry's vector and far from the others.
    const itemVectors = [[0.9, 0.1, 0]]
    const entryVectors = rateIndex.map((_, i) => (i === brickEntryIndex ? [0.95, 0.05, 0] : [0, 0.1, 0.9]))
    const [match] = matchEstimateItems(items, rateIndex, { itemVectors, entryVectors })
    expect(match.candidates).toHaveLength(1)
    expect(match.candidates[0].rate).toBe('6720.10')
    // Found only via the embedding — no keyword overlap contributed at all —
    // so callers should treat this as needing a manual look, not auto-apply.
    expect(match.semanticOnly).toBe(true)
  })

  it('without embeddings supplied, behaves exactly as TF-IDF-only (no regression)', () => {
    const items = [item({ description: 'Brick masonry work in cement mortar excluding pointing and plastering as per drawing.' })]
    const withoutEmbeddings = matchEstimateItems(items, rateIndex)
    const withUndefinedEmbeddings = matchEstimateItems(items, rateIndex, undefined)
    expect(withoutEmbeddings).toEqual(withUndefinedEmbeddings)
  })

  it('resolves an item by its cited index code, even against a differently-worded description', () => {
    const index = [
      entry('Earthwork excavation for road way in soil by mechanical means (Hydraulic Excavator)', '73.08', {
        code: 'RBR-EECD-8(c)',
        unit: 'cum'
      }),
      entry('Excavation for roadway in soil by mechanical means (Dozer)', '55.07', { code: 'RBR-EECD-8(a)', unit: 'cum' })
    ]
    const items = [item({ description: 'RBR-EECD-8(c): Some paraphrased earthwork wording that barely overlaps.' })]
    const [match] = matchEstimateItems(items, index)
    expect(match.candidates).toHaveLength(1)
    expect(match.candidates[0].rate).toBe('73.08')
    expect(isAutoResolved(match)).toBe(true)
  })

  it('resolves a concrete-grade item to exactly the entry carrying that grade', () => {
    const index = [
      entry('Providing concrete ... Transit Mixer and Concrete Pump (RCC) Design Mix M10', '4445.91', {
        code: 'RBR-FNDN-4(M10)',
        unit: 'cum'
      }),
      entry('Providing concrete ... Transit Mixer and Concrete Pump (RCC) Design Mix M30', '5128.80', {
        code: 'RBR-FNDN-4(M30)',
        unit: 'cum'
      })
    ]
    const items = [item({ description: 'Supply and placing of the Ready Mix Concrete (RMC) M30 grade design mix ...' })]
    const [match] = matchEstimateItems(items, index)
    expect(match.candidates).toHaveLength(1)
    expect(match.candidates[0].rate).toBe('5128.80')
  })

  it('drops a candidate whose unit is incompatible with the item, leaving it unresolved', () => {
    // A per-Sqm plastering rate must not be applied to a per-Rmt item, even if
    // the wording overlaps enough to score above the match threshold.
    const index = [entry('Plastering finishing work with sound edges and corners as directed', '41', { unit: 'sqm' })]
    const items = [item({ description: 'Plastering finishing work with sound edges and corners as directed', unit: 'Rmt' })]
    const [match] = matchEstimateItems(items, index)
    expect(match.candidates).toEqual([])
  })
})

describe('isAutoResolved / needsReview', () => {
  const base = (overrides: Partial<ItemMatchLike>): ItemMatchLike => ({
    item: item({}),
    candidates: [{ description: 'x', rate: '1', breakdown: [] }],
    bestScore: 1,
    semanticOnly: false,
    ...overrides
  })
  type ItemMatchLike = Parameters<typeof isAutoResolved>[0]

  it('auto-resolves a confident single candidate', () => {
    const m = base({ bestScore: 0.95 })
    expect(isAutoResolved(m)).toBe(true)
    expect(needsReview(m)).toBe(false)
  })

  it('holds back a medium-confidence single candidate for review (does not auto-apply)', () => {
    const m = base({ bestScore: 0.65 })
    expect(isAutoResolved(m)).toBe(false)
    expect(needsReview(m)).toBe(true)
  })

  it('holds back a semantic-only single candidate even at a high score', () => {
    const m = base({ bestScore: 0.99, semanticOnly: true })
    expect(isAutoResolved(m)).toBe(false)
    expect(needsReview(m)).toBe(true)
  })

  it('treats several candidates as needing a pick, and no candidates as neither', () => {
    const many = base({ candidates: [{ description: 'a', rate: '1', breakdown: [] }, { description: 'a', rate: '2', breakdown: [] }] })
    expect(isAutoResolved(many)).toBe(false)
    expect(needsReview(many)).toBe(true)
    const none = base({ candidates: [] })
    expect(isAutoResolved(none)).toBe(false)
    expect(needsReview(none)).toBe(false)
  })
})

describe('extractItemCode / extractGrade', () => {
  it('reads a departmental code with sub-variant from the start of a description', () => {
    expect(extractItemCode('RBR-EECD-8(c): Earthwork excavation for road way ...')).toBe('RBR-EECD-8(c)')
    expect(extractItemCode('BLD-CSTN-1-1 Cement Mortar')).toBe('BLD-CSTN-1-1')
  })

  it('does not mistake a mid-sentence clause number for the item code', () => {
    expect(extractItemCode('Earthwork in excavation as per clause 305.1 complete')).toBeUndefined()
  })

  it('reads a concrete grade only in a concrete context', () => {
    expect(extractGrade('Supply and placing of the Ready Mix Concrete (RMC) M30 grade')).toBe('M30')
    expect(extractGrade('Structural steel member M20 bolt of grade 8.8')).toBeUndefined()
  })
})

describe('buildCellEdits', () => {
  it('replaces both rate and description, colored green, for a resolved normal item', () => {
    const est = item({ descRow: 5, descCol: 1, rateRow: 5, rateCol: 3, isVariant: false })
    const edits = buildCellEdits([{ item: est, chosen: candidate({ description: 'Matched description', rate: '308.31' }) }])
    expect(edits).toEqual([
      { row: 5, col: 3, value: 308.31 },
      { row: 5, col: 1, value: 'Matched description', color: 'green' }
    ])
  })

  it('only touches the rate and the color for a resolved variant item — never overwrites its shared description text', () => {
    const est = item({ descRow: 7, descCol: 1, rateRow: 7, rateCol: 5, isVariant: true })
    const edits = buildCellEdits([{ item: est, chosen: candidate({ description: 'Matched description', rate: '141.22' }) }])
    expect(edits).toEqual([
      { row: 7, col: 5, value: 141.22 },
      { row: 7, col: 1, color: 'green' }
    ])
  })

  it('colors the description red and leaves the rate untouched when unresolved', () => {
    const est = item({ descRow: 9, descCol: 1, rateRow: 9, rateCol: 3 })
    const edits = buildCellEdits([{ item: est, chosen: null }])
    expect(edits).toEqual([{ row: 9, col: 1, color: 'red' }])
  })

  it('skips items with no known cell position', () => {
    const est = item({})
    const edits = buildCellEdits([{ item: est, chosen: candidate({}) }])
    expect(edits).toEqual([])
  })

  it('rounds the written rate to two decimal places', () => {
    const est = item({ descRow: 1, descCol: 1, rateRow: 1, rateCol: 3 })
    const edits = buildCellEdits([{ item: est, chosen: candidate({ rate: '4445.908823' }) }])
    expect(edits[0]).toEqual({ row: 1, col: 3, value: 4445.91 })
  })
})

describe('buildRateAnalysisSheet', () => {
  it('builds a block per resolved item: a header line followed by its breakdown rows', () => {
    const est1 = item({ descRow: 1, descCol: 1, rateRow: 1, rateCol: 3 })
    const est2 = item({ descRow: 2, descCol: 1, rateRow: 2, rateCol: 3 })
    const resolved = [
      {
        item: est1,
        chosen: candidate({
          description: 'Earthwork in excavation',
          rate: '308.31',
          breakdown: [
            ['RBR-FNDN-1', '1', 'Earthwork in excavation'],
            ['', '', 'Rate per cum', '308.31']
          ]
        })
      },
      { item: est2, chosen: null }
    ]
    const sheet = buildRateAnalysisSheet(resolved)
    expect(sheet).toEqual([
      ['Rate Analysis: Earthwork in excavation'],
      ['RBR-FNDN-1', '1', 'Earthwork in excavation'],
      ['', '', 'Rate per cum', '308.31'],
      []
    ])
  })

  it('contributes nothing for unresolved items or matches with no captured breakdown', () => {
    const est = item({ descRow: 1, descCol: 1, rateRow: 1, rateCol: 3 })
    expect(buildRateAnalysisSheet([{ item: est, chosen: null }])).toEqual([])
    expect(buildRateAnalysisSheet([{ item: est, chosen: candidate({ breakdown: [] }) }])).toEqual([])
  })
})
