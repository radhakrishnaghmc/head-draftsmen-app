import { describe, expect, it } from 'vitest'
import { findHtmlPlaceholders, fillHtmlPlaceholders } from '../core/htmlPlaceholders'

describe('findHtmlPlaceholders', () => {
  it('finds every distinct {{Label}}, trimmed, in first-appearance order', () => {
    const html = '<p>{{ Name of the work }}</p><p>{{Date}}</p><p>{{Name of the work}}</p>'
    expect(findHtmlPlaceholders(html)).toEqual(['Name of the work', 'Date'])
  })

  it('returns an empty array when the template has no placeholders', () => {
    expect(findHtmlPlaceholders('<p>Plain HTML, nothing to fill</p>')).toEqual([])
  })
})

describe('fillHtmlPlaceholders', () => {
  it('replaces every occurrence of a label with its resolved value', () => {
    const html = '<p>Work: {{Name of the work}}</p><footer>{{Name of the work}}</footer>'
    const out = fillHtmlPlaceholders(html, { 'Name of the work': 'Laying of UGD at Nizampet' })
    expect(out).toBe('<p>Work: Laying of UGD at Nizampet</p><footer>Laying of UGD at Nizampet</footer>')
  })

  it('replaces an unresolved placeholder with an empty string rather than leaving the token', () => {
    const html = '<p>{{Date}}</p>'
    expect(fillHtmlPlaceholders(html, {})).toBe('<p></p>')
  })

  it('tolerates stray whitespace inside the {{ }} tokens', () => {
    const html = '<p>{{  Date  }}</p>'
    expect(fillHtmlPlaceholders(html, { Date: '12.07.2026' })).toBe('<p>12.07.2026</p>')
  })
})
