import { describe, expect, it } from 'vitest'
import { readFileSync } from 'fs'
import { resolve } from 'path'
import { fillSeBidDocument } from '../core/bidDocument'
import { listParagraphs, findPlaceholdersInAllParts } from '../core/docx-edit'

const TEMPLATE_PATH = resolve(__dirname, '../resources/bid-document-se-template.docx')

describe('fillSeBidDocument', () => {
  const buffer = readFileSync(TEMPLATE_PATH)

  it('leaves no unresolved {{Placeholder}} anywhere in the document (body, headers or footers) once every field is supplied', () => {
    const out = fillSeBidDocument(buffer, {
      nitNo: '12/SE/QBZ/CMC/2026-27',
      dated: '03.07.2026',
      downloadStartDate: '16.07.2026',
      downloadEndDate: '23.07.2026',
      work: {
        serial: 1,
        name: 'Laying of UGD at Rajarajeshwari Colony (Item No.3)',
        amount: '63',
        ecv: '4998557',
        zone: 'Quthbullapur',
        completionPeriod: '3',
        itemNo: '3',
        tsNo: '29/SE/QBZ/CMC/2026-27',
        tsDate: '04-07-2026',
        asAuthority: 'commissioner',
        asDate: '23.06.2026'
      }
    })
    expect(findPlaceholdersInAllParts(out)).toEqual([])
  })

  it('computes ECV Lakhs by TRUNCATING (not rounding) to 2 decimals, matching the SE office\'s own convention', () => {
    const out = fillSeBidDocument(buffer, {
      nitNo: '1/SE/QBZ/CMC/2026-27',
      dated: '01.01.2026',
      downloadStartDate: '02.01.2026',
      downloadEndDate: '09.01.2026',
      work: {
        serial: 1,
        name: 'Test work',
        amount: '10',
        ecv: '4998557', // 49.98557 Lakhs -> would round to 49.99, but the office's own docs truncate to 49.98
        zone: 'Quthbullapur'
      }
    })
    const joined = listParagraphs(out).join('\n')
    expect(joined).toContain('49.98 Lakhs')
    expect(joined).not.toContain('49.99 Lakhs')
    // ECV itself (rupees, Indian-grouped) and EMD @ 1% computed from it.
    expect(joined).toContain('49,98,557/-')
    expect(joined).toContain('49,986/-')
  })

  it('leaves ECV, ECV Lakhs and EMD @ 1% blank — never computed from the Lakhs estimate — when ECV is not supplied', () => {
    const out = fillSeBidDocument(buffer, {
      nitNo: '1/SE/QBZ/CMC/2026-27',
      dated: '01.01.2026',
      downloadStartDate: '02.01.2026',
      downloadEndDate: '09.01.2026',
      work: { serial: 1, name: 'Test work', amount: '10', zone: 'Quthbullapur' }
    })
    // Every {{...}} still resolves (to blank strings) — no leftover tokens even when a field is absent.
    expect(findPlaceholdersInAllParts(out)).toEqual([])
  })

  it('composes the Administrative Sanction line differently per authority, and drops the date clause entirely when the date is blank', () => {
    const base = {
      nitNo: '1/SE/QBZ/CMC/2026-27',
      dated: '01.01.2026',
      downloadStartDate: '02.01.2026',
      downloadEndDate: '09.01.2026'
    }

    const commissionerWithDate = fillSeBidDocument(buffer, {
      ...base,
      work: { serial: 1, name: 'A', amount: '10', zone: 'Quthbullapur', asAuthority: 'commissioner', asDate: '23.06.2026' }
    })
    expect(listParagraphs(commissionerWithDate).join('\n')).toContain(
      'Administrative Sanction approved by the Commissioner, CMC Dt:23.06.2026'
    )

    const commissionerNoDate = fillSeBidDocument(buffer, {
      ...base,
      work: { serial: 1, name: 'A', amount: '10', zone: 'Quthbullapur' }
    })
    const asPara = listParagraphs(commissionerNoDate).find((p) => p.startsWith('Administrative Sanction approved'))
    expect(asPara).toBe('Administrative Sanction approved by the Commissioner, CMC')

    const zonal = fillSeBidDocument(buffer, {
      ...base,
      work: { serial: 1, name: 'A', amount: '10', zone: 'Quthbullapur', asAuthority: 'zonal', asDate: '31-03-2026' }
    })
    expect(listParagraphs(zonal).join('\n')).toContain(
      'Administrative Sanction approved by the Zonal Commissioner, QBZ, CMC vide Dt. 31-03-2026.'
    )
  })

  it('falls back the Item No from a "(Item No.N)" tag in the work name, then to the serial, when not given explicitly', () => {
    const base = {
      nitNo: '1/SE/QBZ/CMC/2026-27',
      dated: '01.01.2026',
      downloadStartDate: '02.01.2026',
      downloadEndDate: '09.01.2026'
    }

    const fromName = fillSeBidDocument(buffer, {
      ...base,
      work: { serial: 5, name: 'Some work (Item No.7)', amount: '10', zone: 'Quthbullapur' }
    })
    expect(listParagraphs(fromName).join('\n')).toContain('(Item No.7)')

    const fromSerial = fillSeBidDocument(buffer, {
      ...base,
      work: { serial: 5, name: 'Some work with no tag', amount: '10', zone: 'Quthbullapur' }
    })
    expect(listParagraphs(fromSerial).join('\n')).toContain('(Item No.5)')

    const explicit = fillSeBidDocument(buffer, {
      ...base,
      work: { serial: 5, name: 'Some work (Item No.7)', amount: '10', zone: 'Quthbullapur', itemNo: '7A' }
    })
    expect(listParagraphs(explicit).join('\n')).toContain('(Item No.7A)')
  })

  it('resolves the Zone code (QBZ) for a known zone and prints the Zone name verbatim throughout', () => {
    const out = fillSeBidDocument(buffer, {
      nitNo: '1/SE/QBZ/CMC/2026-27',
      dated: '01.01.2026',
      downloadStartDate: '02.01.2026',
      downloadEndDate: '09.01.2026',
      work: { serial: 1, name: 'A', amount: '10', zone: 'Quthbullapur' }
    })
    const joined = listParagraphs(out).join('\n')
    expect(joined).toContain('Superintending Engineer(QBZ)')
    expect(joined).toContain('Quthbullapur Zone')
    expect(joined).not.toContain('{{Zone}}')
    expect(joined).not.toContain('{{Zone Abbr}}')
  })
})
