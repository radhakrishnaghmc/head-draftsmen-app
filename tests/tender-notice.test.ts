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
      today: '21.07.2026',
      groups: [
        { startDate: '18.08.2026', endDate: '25.08.2026', items: [{ name: 'Test work item', amount: '1.00' }] }
      ]
    })

    const paragraphs = listParagraphs(out)
    const nitLine = paragraphs.find((p) => p.startsWith('NIT No. '))
    expect(nitLine).toBeDefined()
    expect(nitLine).toBe('NIT No. 16/DB/EE/Gajularamaram Circle-57/CMC/2026-27Dated:21.07.2026')

    const lrLine = paragraphs.find((p) => p.startsWith('Lr.No:'))
    expect(lrLine).toBe('Lr.No:16/DB/EE/Gajularamaram Circle-57/CMC/2026-27Date:21.07.2026')
  })

  it('shows the summary table a work-type count breakdown instead of listing every work name', () => {
    const buffer = readFileSync(TEMPLATE_PATH)
    const out = fillTenderNotice(buffer, {
      nitNo: '16/DB/EE/Gajularamaram Circle-57/CMC/2026-27',
      today: '21.07.2026',
      groups: [
        {
          startDate: '18.08.2026',
          endDate: '25.08.2026',
          items: [
            { name: 'Work one', amount: '1.00' },
            { name: 'Work two', amount: '2.00' },
            { name: 'Work three', amount: '3.00' }
          ]
        }
      ]
    })

    const paragraphs = listParagraphs(out)
    expect(paragraphs).toContain('Maintenance: 3')
    expect(paragraphs.some((p) => p.includes('Work one'))).toBe(true) // still in the item table
    expect(paragraphs.some((p) => p === 'Work one; Work two; Work three')).toBe(false)
  })

  it('breaks the count down by work type inferred from each work name', () => {
    const buffer = readFileSync(TEMPLATE_PATH)
    const out = fillTenderNotice(buffer, {
      nitNo: '16/DB/EE/Gajularamaram Circle-57/CMC/2026-27',
      today: '21.07.2026',
      groups: [
        {
          startDate: '18.08.2026',
          endDate: '25.08.2026',
          items: [
            { name: 'Laying of CC road from X to Y', amount: '1.00' },
            { name: 'Laying of CC road near Z', amount: '1.00' },
            { name: 'Construction of graveyard at Z', amount: '1.00' },
            { name: 'Maintenance of BT roads in ward 5', amount: '1.00' },
            { name: 'Maintenance of CC roads in ward 6', amount: '1.00' },
            { name: 'Desilting works in ward 7', amount: '1.00' },
            { name: 'Desilting works in ward 8', amount: '1.00' },
            { name: 'Desilting works in ward 9', amount: '1.00' }
          ]
        }
      ]
    })

    const paragraphs = listParagraphs(out)
    expect(paragraphs).toContain('Graveyard: 1, Desilting: 3, Maintenance: 2, CC Roads: 2')
  })

  it('counts a work under its Works List "Type of work" value, not the name-inferred one, when both disagree', () => {
    const buffer = readFileSync(TEMPLATE_PATH)
    const out = fillTenderNotice(buffer, {
      nitNo: '16/DB/EE/Gajularamaram Circle-57/CMC/2026-27',
      today: '21.07.2026',
      groups: [
        {
          startDate: '18.08.2026',
          endDate: '25.08.2026',
          items: [
            // Name reads like CC ROADS, but the Works List says it's really HALLS — that wins.
            { name: 'Laying of CC road from X to Y', amount: '1.00', typeOfWork: 'HALLS' },
            // No Works List value — falls back to the name-inferred category.
            { name: 'Construction of graveyard at Z', amount: '1.00' },
            // A free-text Works List value that isn't one of the known categories is kept as its own bucket.
            { name: 'Some other work', amount: '1.00', typeOfWork: 'Street Lighting' }
          ]
        }
      ]
    })

    const paragraphs = listParagraphs(out)
    expect(paragraphs).toContain('Graveyard: 1, Halls: 1, Street Lighting: 1')
  })

  it('rewrites the template circle name+number to the issuing office when a different circle is chosen', () => {
    const buffer = readFileSync(TEMPLATE_PATH)
    const out = fillTenderNotice(buffer, {
      nitNo: '22/DB/EE/Nizampet Circle-58/CMC/2026-27',
      today: '21.07.2026',
      groups: [
        { startDate: '18.08.2026', endDate: '25.08.2026', items: [{ name: 'Test work item', amount: '1.00' }] }
      ],
      circle: 'Nizampet',
      circleNumber: '58',
      zone: 'Quthbullapur'
    })

    const text = listParagraphs(out).join('\n')
    // The template's own circle identity must be gone EVERYWHERE — the name, the
    // "Circle-57"/"Circle 57" forms, and the EE e-mail that embeds the number
    // ("eec57…"), which is the spot that used to keep leaking circle 57.
    expect(text).not.toContain('Gajularamaram')
    expect(text).not.toContain('Circle-57')
    expect(text).not.toContain('Circle 57')
    expect(text).not.toContain('eec57')
    // …replaced by the issuing office's throughout the boilerplate.
    expect(text).toContain('Nizampet Circle-58')
    expect(text.includes('Nizampet Circle 58') || text.includes('EE – Nizampet Circle 58')).toBe(true)
    expect(text).toContain('eec58')
  })

  it('swaps the e-mail and both mobile numbers when supplied, and derives the e-mail circle when not', () => {
    const buffer = readFileSync(TEMPLATE_PATH)
    const out = fillTenderNotice(buffer, {
      nitNo: '22/DB/EE/Nizampet Circle-58/CMC/2026-27',
      today: '21.07.2026',
      groups: [
        { startDate: '18.08.2026', endDate: '25.08.2026', items: [{ name: 'Test work item', amount: '1.00' }] }
      ],
      circle: 'Nizampet',
      circleNumber: '58',
      zone: 'Quthbullapur',
      email: 'nizampet.ee@gmail.com',
      eePhone: '9000000001',
      hdPhone: '9000000002'
    })
    const text = listParagraphs(out).join('\n')
    expect(text).toContain('nizampet.ee@gmail.com')
    expect(text).toContain('9000000001')
    expect(text).toContain('9000000002')
    // The template's own contact values are gone.
    expect(text).not.toContain('eec57')
    expect(text).not.toContain('7893066262')
    expect(text).not.toContain('9063836115')
  })

  it('derives the circle from the NIT number when no office circle fields are passed', () => {
    const buffer = readFileSync(TEMPLATE_PATH)
    // Office fields omitted entirely — only a Nizampet NIT No is given.
    const out = fillTenderNotice(buffer, {
      nitNo: '16/DB/EE/Nizampet Circle-58/CMC/2026-27',
      today: '21.07.2026',
      groups: [
        { startDate: '18.08.2026', endDate: '25.08.2026', items: [{ name: 'Test work item', amount: '1.00' }] }
      ]
    })
    const text = listParagraphs(out).join('\n')
    expect(text).not.toContain('Gajularamaram')
    expect(text).not.toContain('Circle-57')
    expect(text).not.toContain('Circle 57')
    expect(text).not.toContain('eec57')
    expect(text).toContain('Nizampet Circle-58')
  })

  it('leaves the template place names untouched for a genuine Gajularamaram Circle-57 notice', () => {
    const buffer = readFileSync(TEMPLATE_PATH)
    const out = fillTenderNotice(buffer, {
      nitNo: '16/DB/EE/Gajularamaram Circle-57/CMC/2026-27',
      today: '21.07.2026',
      groups: [
        { startDate: '18.08.2026', endDate: '25.08.2026', items: [{ name: 'Test work item', amount: '1.00' }] }
      ]
    })
    const text = listParagraphs(out).join('\n')
    expect(text).toContain('Gajularamaram Circle-57')
  })

  it('supports multiple download-date groups — one Schedule block + consolidation row per group, labelled "Sl No.X to Y"', () => {
    const buffer = readFileSync(TEMPLATE_PATH)
    const out = fillTenderNotice(buffer, {
      nitNo: '24/DB/EE/Nizampet Circle-58/CMC/2026-27',
      today: '11.09.2026',
      groups: [
        {
          startDate: '15.09.2026',
          endDate: '18.09.2026',
          items: [
            { name: 'Work one', amount: '15.00' },
            { name: 'Work two', amount: '43.00' },
            { name: 'Work three', amount: '15.00' }
          ]
        },
        {
          startDate: '15.09.2026',
          endDate: '22.09.2026',
          items: [
            { name: 'Work four', amount: '6.59' },
            { name: 'Work five', amount: '49.96' }
          ]
        },
        {
          startDate: '15.09.2026',
          endDate: '17.09.2026',
          items: [{ name: 'Work six', amount: '20.00' }]
        }
      ]
    })

    const paragraphs = listParagraphs(out)
    expect(paragraphs).toContain('Sl No.01 to 03')
    expect(paragraphs).toContain('Sl No.04 to 05')
    expect(paragraphs).toContain('Sl No.06')
    expect(paragraphs).toContain('Maintenance: 3')
    expect(paragraphs).toContain('Maintenance: 2')
    expect(paragraphs).toContain('Maintenance: 1')

    const text = paragraphs.join('\n')
    expect(text).toContain('22.09.2026')
    expect(text).toContain('17.09.2026')

    // The item table itself stays one continuous list, Sl.No 1..6 — only the
    // Schedule/consolidation blocks below it are split by group.
    expect(paragraphs.filter((p) => p === 'Work one').length).toBe(1)
    expect(paragraphs.some((p) => p.includes('Work six'))).toBe(true)
  })
})
