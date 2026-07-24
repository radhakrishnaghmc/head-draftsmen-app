import { describe, expect, it } from 'vitest'
import { classifyHeaders, mismatchHint, docKindPrototypeTexts } from '../src/docClassify'
import { WORKS_COLUMNS } from '../src/worksSchema'
import { BOQ_HEADERS } from '../core/boqHeaders'

describe('classifyHeaders', () => {
  it('recognizes an exact Works List header set', () => {
    const [top] = classifyHeaders(WORKS_COLUMNS)
    expect(top.kind).toBe('works-list')
  })

  it('recognizes an exact BOQ header set', () => {
    const [top] = classifyHeaders(BOQ_HEADERS)
    expect(top.kind).toBe('boq')
  })

  it('recognizes a typical estimate header set', () => {
    const [top] = classifyHeaders(['S.No', 'Description', 'Qty', 'Rate', 'Unit'])
    expect(top.kind).toBe('estimate')
  })
})

describe('mismatchHint', () => {
  it('stays silent when the headers match the expected kind', () => {
    expect(mismatchHint(WORKS_COLUMNS, 'works-list')).toBeNull()
  })

  it('flags a BOQ uploaded where an Estimate was expected', () => {
    const hint = mismatchHint(BOQ_HEADERS, 'estimate')
    expect(hint).toContain('BOQ')
    expect(hint).toContain('Estimate')
  })

  it('flags an Estimate uploaded where a BOQ was expected', () => {
    const hint = mismatchHint(['S.No', 'Description', 'Qty', 'Rate', 'Unit'], 'boq')
    expect(hint).toContain('Estimate')
  })

  it('stays silent on a low-confidence, ambiguous header set rather than guessing', () => {
    expect(mismatchHint(['Column A', 'Column B'], 'estimate')).toBeNull()
  })
})

describe('embedding-assisted classification', () => {
  it('uses the embedding score when it is stronger than the keyword overlap', () => {
    const headers = ['Approx Qty', 'Cost per Unit', 'Measure'] // deliberately low keyword overlap with any prototype
    const texts = docKindPrototypeTexts()
    expect(texts).toHaveLength(3) // works-list, estimate, boq — keeps this test honest if a profile is ever added/removed

    // A fake embedding where the uploaded document's vector is closest to the "estimate" prototype.
    const embeddings = {
      documentVector: [1, 0, 0],
      profileVectors: [
        [0, 1, 0], // works-list
        [0.95, 0.05, 0], // estimate
        [0, 0, 1] // boq
      ]
    }
    const [top] = classifyHeaders(headers, embeddings)
    expect(top.kind).toBe('estimate')
  })
})
