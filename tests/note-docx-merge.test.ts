import { describe, it, expect } from 'vitest'
import { noteDocxHtml } from '../src/notePagination'

describe('noteDocxHtml', () => {
  it('rejoins a table the preview split across two pages into one table', () => {
    const t = (rows: string) => `<table style="w"><tbody>${rows}</tbody></table>`
    const html = noteDocxHtml([
      [{ n: 0, html: t('<tr><td>1</td></tr><tr><td>2</td></tr>'), first: true, last: false }],
      [{ n: 0, html: t('<tr><td>3</td></tr>'), first: false, last: true }]
    ])
    expect(html.match(/<table/g)).toHaveLength(1)
    expect(html).toContain('<td>2</td></tr><tr><td>3</td>')
  })
})
