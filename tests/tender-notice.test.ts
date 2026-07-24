import { describe, expect, it } from 'vitest'
import { readFileSync } from 'fs'
import { resolve } from 'path'
import { fillTenderNotice } from '../core/tenderNotice'
import { listParagraphs } from '../core/docx-edit'

const TEMPLATE_PATH = resolve(__dirname, '../resources/tender-notice-template.docx')

describe('fillTenderNotice', () => {
  it('keeps the NIT No. and Dated values in their own paragraph text, not merged', () => {
    const buffer = readFileSync(TEMPLATE_PATH)
    const out = fillTenderNotice(buffer, {
      nitNo: '16/DB/EE/Gajularamaram Circle-57/CMC/2026-27',
      startDate: '18.08.2026',
      endDate: '25.08.2026',
      today: '21.07.2026',
      items: [{ name: 'Test work item', amount: '1.00' }]
    })

    const paragraphs = listParagraphs(out)
    const nitLine = paragraphs.find((p) => p.startsWith('NIT No. '))
    expect(nitLine).toBeDefined()
    expect(nitLine).toBe('NIT No. 16/DB/EE/Gajularamaram Circle-57/CMC/2026-27Dated:21.07.2026')

    const lrLine = paragraphs.find((p) => p.startsWith('Lr.No:'))
    expect(lrLine).toBe('Lr.No:16/DB/EE/Gajularamaram Circle-57/CMC/2026-27Date:21.07.2026')
  })

  it('shows the summary table work count instead of listing every work name', () => {
    const buffer = readFileSync(TEMPLATE_PATH)
    const out = fillTenderNotice(buffer, {
      nitNo: '16/DB/EE/Gajularamaram Circle-57/CMC/2026-27',
      startDate: '18.08.2026',
      endDate: '25.08.2026',
      today: '21.07.2026',
      items: [
        { name: 'Work one', amount: '1.00' },
        { name: 'Work two', amount: '2.00' },
        { name: 'Work three', amount: '3.00' }
      ]
    })

    const paragraphs = listParagraphs(out)
    expect(paragraphs).toContain('No of works: 3')
    expect(paragraphs.some((p) => p.includes('Work one'))).toBe(true) // still in the item table
    expect(paragraphs.some((p) => p === 'Work one; Work two; Work three')).toBe(false)
  })
})
