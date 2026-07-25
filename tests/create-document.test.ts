import { describe, it, expect } from 'vitest'
import { findPlaceholders, matchPlaceholdersToColumns } from '../core/createDocument'

describe('findPlaceholders', () => {
  it('finds a single placeholder', () => {
    expect(findPlaceholders('Name of Work: {{Name of the work}}')).toEqual(['Name of the work'])
  })

  it('finds multiple distinct placeholders in first-seen order', () => {
    const html = '<p>{{Name of the work}}</p><p>Dated {{Agmt Date}}</p>'
    expect(findPlaceholders(html)).toEqual(['Name of the work', 'Agmt Date'])
  })

  it('dedupes a placeholder that appears more than once', () => {
    const html = '{{Name of the work}} ... {{Name of the work}}'
    expect(findPlaceholders(html)).toEqual(['Name of the work'])
  })

  it('trims stray whitespace inside the braces', () => {
    expect(findPlaceholders('{{  Tender Percentage  }}')).toEqual(['Tender Percentage'])
  })

  it('returns an empty array when there are no placeholders', () => {
    expect(findPlaceholders('<p>Just some plain pasted text.</p>')).toEqual([])
  })

  it('reads a placeholder split across several tags as one clean label, not markup noise', () => {
    // Real bug: Word's clipboard HTML frequently splits what looks like one
    // run of text into several <span>/<b> elements (spell-check, language
    // tagging, bidi font sizing) even mid-placeholder. Matching {{...}}
    // against the raw HTML string picked up everything between — the
    // intervening tags included — as the "label", producing something like
    // "Amount&nbsp;</span></b><b style=\"...\">Estimate" that could never
    // resolve to a real column.
    const html =
      '{{Amount&nbsp;</span></b><b style="font-size: 12pt;"><span lang="EN-US" style="font-size:11.0pt">&nbsp;of&nbsp;</span></b><b style="font-size: 12pt;"><span lang="EN-US" style="font-size:11.0pt">Estimate}}'
    expect(findPlaceholders(html)).toEqual(['Amount of Estimate'])
  })

  it('decodes &nbsp; to a real space when reading a label', () => {
    expect(findPlaceholders('{{Name&nbsp;of&nbsp;the&nbsp;work}}')).toEqual(['Name of the work'])
  })
})

describe('matchPlaceholdersToColumns', () => {
  it('auto-resolves a confident single match via embeddings', () => {
    const labels = ['Agency Name']
    const columns = ['Name of the Contractor', 'WIN Code']
    const embeddings = {
      labelVectors: [[1, 0]],
      columnVectors: [
        [0.95, 0.05],
        [0, 1]
      ]
    }
    const [match] = matchPlaceholdersToColumns(labels, columns, embeddings)
    expect(match.column).toBe('Name of the Contractor')
  })

  it('leaves a label unresolved when nothing scores above the threshold', () => {
    const labels = ['Zzyx qwibblonium']
    const columns = ['Name of the work', 'WIN Code']
    const embeddings = {
      labelVectors: [[1, 0]],
      columnVectors: [
        [0, 1],
        [0, -1]
      ]
    }
    const [match] = matchPlaceholdersToColumns(labels, columns, embeddings)
    expect(match.column).toBeNull()
  })

  it('falls back to token overlap when no embeddings are supplied', () => {
    const labels = ['Name of Work']
    const columns = ['Name of the work', 'WIN Code']
    const [match] = matchPlaceholdersToColumns(labels, columns)
    expect(match.column).toBe('Name of the work')
  })

  it('returns null for every label when there are no columns', () => {
    const [match] = matchPlaceholdersToColumns(['Name of Work'], [])
    expect(match.column).toBeNull()
  })

  it('takes whichever signal is stronger — a weak embedding score does not drag down a strong keyword match', () => {
    const labels = ['Name of Work']
    const columns = ['Name of the work', 'WIN Code']
    // Embeddings score both columns weakly/negatively (a poor model output), but the
    // literal word overlap with "Name of the work" is strong enough to win on its own.
    const embeddings = {
      labelVectors: [[1, 0]],
      columnVectors: [
        [0, 1],
        [0, -1]
      ]
    }
    const [match] = matchPlaceholdersToColumns(labels, columns, embeddings)
    expect(match.column).toBe('Name of the work')
  })

  it('takes whichever signal is stronger — a weak keyword score does not drag down a strong embedding match', () => {
    const labels = ['Agency Name']
    const columns = ['Name of the Contractor', 'WIN Code']
    // "Agency Name" vs "Name of the Contractor" shares only one token, but the
    // embedding score is confident — that alone should be enough to resolve it.
    const embeddings = {
      labelVectors: [[1, 0]],
      columnVectors: [
        [0.95, 0.05],
        [0, 1]
      ]
    }
    const [match] = matchPlaceholdersToColumns(labels, columns, embeddings)
    expect(match.column).toBe('Name of the Contractor')
    expect(match.score).toBeGreaterThan(0.9) // the embedding score, not the weak keyword overlap
  })
})

