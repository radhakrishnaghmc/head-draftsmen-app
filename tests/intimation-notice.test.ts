import { describe, expect, it } from 'vitest'
import { readFileSync } from 'fs'
import { parseIntimationNotice, parseIntimationNoticeLines } from '../core/intimationNotice'

// A trimmed copy of the real Telangana portal "viewIntimationNotice" page.
const SAMPLE = `
<p align="center"><b><u>DATE: Friday, July 24, 2026</u></b></p>
<p>To
<br>M V S CONSTRUCTIONS
<br>13/B, Allwyn Colony, Phase 2, Kukatpally
<br>
<br>Hyderabad -500072
<br>Telangana</p>
<p>Sir/Madam,</p>
<p align="justify">This is notify you that the bid submitted by you for execution of the
<b>NIT No. 12/DB/EE/Nizampet Circle-58/CMC/2026-27 item 1 Dated:15.07.2026</b> at contract price of Rs.
<b> 1416455.93 ( Fourteen Lakh Sixteen Thousand Four Hundred and Fifty Five Rupees Ninety Three Paisa)</b>
as corrected and modified.</p>
<table>
  <thead><tr><th>Company Name</th><th>Estimated Contract Value</th><th>Corpus Fund @ 0.04 %</th></tr></thead>
  <tbody><tr><td align="center">M V S CONSTRUCTIONS</td><td align="center">1593493.00</td><td align="center">638.00</td></tr></tbody>
</table>
`

describe('parseIntimationNotice', () => {
  it('extracts the agency name and full address from the "To" block', () => {
    const r = parseIntimationNotice(SAMPLE)
    expect(r.agencyName).toBe('M V S CONSTRUCTIONS')
    expect(r.address).toBe('13/B, Allwyn Colony, Phase 2, Kukatpally, Hyderabad -500072, Telangana')
  })

  it('extracts the NIT No without the trailing item/date tail', () => {
    expect(parseIntimationNotice(SAMPLE).nitNo).toBe('12/DB/EE/Nizampet Circle-58/CMC/2026-27')
  })

  it('extracts the accepted contract value in rupees', () => {
    expect(parseIntimationNotice(SAMPLE).contractRupees).toBe(1416455.93)
  })

  it('extracts the Estimated Contract Value from the summary table', () => {
    expect(parseIntimationNotice(SAMPLE).ecvRupees).toBe(1593493)
  })

  it('returns an empty object for HTML with none of the expected structure', () => {
    expect(parseIntimationNotice('<html><body><p>Nothing here</p></body></html>')).toEqual({})
  })

  it('parses the real saved portal page if present', () => {
    const path = '/Users/radhakrishnapodugu/Downloads/viewIntimationNoticealeadp circle.html'
    let html: string
    try {
      html = readFileSync(path, 'utf8')
    } catch {
      return // file not present in CI/other machines — the synthetic cases above cover the logic
    }
    const r = parseIntimationNotice(html)
    expect(r.agencyName).toBe('M V S CONSTRUCTIONS')
    expect(r.address).toContain('Allwyn Colony')
    expect(r.address).toContain('Hyderabad')
    expect(r.nitNo).toContain('Nizampet Circle-58')
    expect(r.contractRupees).toBeGreaterThan(0)
    expect(r.ecvRupees).toBe(1593493)
  })
})

describe('parseIntimationNoticeLines (PDF-printout text)', () => {
  it('reads agency + address from a "To" block with the label on its own line', () => {
    const lines = [
      'DATE: Friday, July 24, 2026',
      'To',
      'M V S CONSTRUCTIONS',
      '13/B, Allwyn Colony, Phase 2, Kukatpally',
      'Hyderabad -500072',
      'Telangana',
      'Sir/Madam,',
      'This is notify you that the bid submitted by you ...'
    ]
    const r = parseIntimationNoticeLines(lines)
    expect(r.agencyName).toBe('M V S CONSTRUCTIONS')
    expect(r.address).toBe('13/B, Allwyn Colony, Phase 2, Kukatpally, Hyderabad -500072, Telangana')
  })

  it('reads agency + address when "To" is inline with the agency name', () => {
    const lines = ['To M V S CONSTRUCTIONS', '13/B, Allwyn Colony', 'Hyderabad -500072', 'Sir/Madam,']
    const r = parseIntimationNoticeLines(lines)
    expect(r.agencyName).toBe('M V S CONSTRUCTIONS')
    expect(r.address).toBe('13/B, Allwyn Colony, Hyderabad -500072')
  })

  it('does not mistake a word starting with "To" (e.g. "Toilets") for the address block', () => {
    const lines = ['Toilets at MPPS in ward 12', 'To', 'ABC BUILDERS', 'Road No 1, Hyderabad', 'Sir,']
    const r = parseIntimationNoticeLines(lines)
    expect(r.agencyName).toBe('ABC BUILDERS')
    expect(r.address).toBe('Road No 1, Hyderabad')
  })

  it('returns empty when there is no address block', () => {
    expect(parseIntimationNoticeLines(['Some heading', 'Random text'])).toEqual({})
  })
})
