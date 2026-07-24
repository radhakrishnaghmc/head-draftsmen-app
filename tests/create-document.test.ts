import { describe, it, expect } from 'vitest'
import { findPlaceholders, matchPlaceholdersToColumns, fillDocumentHtml, bakeFixedPlaceholders } from '../core/createDocument'
import type { PlaceholderMatch } from '../core/createDocument'

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

describe('fillDocumentHtml', () => {
  it('replaces a resolved placeholder with the row value', () => {
    const resolved: PlaceholderMatch[] = [{ label: 'Name of the work', column: 'Name of Work', score: 0.9 }]
    const row = { 'Name of Work': 'Road repair, Ward 12' }
    const html = fillDocumentHtml('<p>{{Name of the work}}</p>', resolved, row)
    expect(html).toBe('<p>Road repair, Ward 12</p>')
  })

  it('blanks an unresolved placeholder rather than leaving the token in place', () => {
    const resolved: PlaceholderMatch[] = [{ label: 'Mystery Field', column: null, score: 0 }]
    const html = fillDocumentHtml('Value: {{Mystery Field}}.', resolved, {})
    expect(html).toBe('Value: .')
  })

  it('replaces every occurrence of a repeated placeholder', () => {
    const resolved: PlaceholderMatch[] = [{ label: 'Name of the work', column: 'Name of Work', score: 0.9 }]
    const row = { 'Name of Work': 'Bridge work' }
    const html = fillDocumentHtml('{{Name of the work}} ... {{Name of the work}}', resolved, row)
    expect(html).toBe('Bridge work ... Bridge work')
  })

  it('HTML-escapes the substituted value', () => {
    const resolved: PlaceholderMatch[] = [{ label: 'Agency', column: 'Agency', score: 0.9 }]
    const row = { Agency: 'M/s A & B <Contractors>' }
    const html = fillDocumentHtml('{{Agency}}', resolved, row)
    expect(html).toBe('M/s A &amp; B &lt;Contractors&gt;')
  })

  it('collapses a placeholder split across several tags into a single plain replacement', () => {
    const resolved: PlaceholderMatch[] = [{ label: 'Amount of Estimate', column: 'Amount of Estimate', score: 0.9 }]
    const row = { 'Amount of Estimate': 'Rs. 8,00,000' }
    const html =
      '<p>{{Amount&nbsp;</span></b><b style="font-size: 12pt;"><span lang="EN-US">&nbsp;of&nbsp;</span></b><b style="font-size: 12pt;"><span lang="EN-US">Estimate}}</p>'
    expect(fillDocumentHtml(html, resolved, row)).toBe('<p>Rs. 8,00,000</p>')
  })

  it('leaves the surrounding markup outside the placeholder span untouched', () => {
    const resolved: PlaceholderMatch[] = [{ label: 'Name of the work', column: 'Name of Work', score: 0.9 }]
    const row = { 'Name of Work': 'Road repair' }
    const html = '<p class="MsoNormal" style="margin:0 0 8pt">{{Name of the work}}</p>'
    expect(fillDocumentHtml(html, resolved, row)).toBe('<p class="MsoNormal" style="margin:0 0 8pt">Road repair</p>')
  })
})

describe('bakeFixedPlaceholders', () => {
  it('replaces only the given labels, matched case-insensitively', () => {
    const html = '<p>{{Circle}} / {{circle}} / {{Zone}} / {{CNO}}</p>'
    const out = bakeFixedPlaceholders(html, { zone: 'Cyberabad', circle: 'Gajularamaram Circle-57', cno: '57' })
    expect(out).toBe('<p>Gajularamaram Circle-57 / Gajularamaram Circle-57 / Cyberabad / 57</p>')
  })

  it('leaves every other placeholder untouched for later per-row resolution', () => {
    const html = '{{Circle}} — {{Name of the work}} — {{Estimate Amount}}'
    const out = bakeFixedPlaceholders(html, { circle: 'Gajularamaram Circle-57' })
    expect(out).toBe('Gajularamaram Circle-57 — {{Name of the work}} — {{Estimate Amount}}')
  })

  it('is idempotent — running it again after the label is gone changes nothing', () => {
    const once = bakeFixedPlaceholders('{{Zone}}', { zone: 'Cyberabad' })
    const twice = bakeFixedPlaceholders(once, { zone: 'Cyberabad' })
    expect(twice).toBe('Cyberabad')
  })

  it('HTML-escapes the baked-in value', () => {
    const out = bakeFixedPlaceholders('{{Circle}}', { circle: 'M/s A & B' })
    expect(out).toBe('M/s A &amp; B')
  })
})
