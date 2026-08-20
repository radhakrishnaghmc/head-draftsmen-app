import { describe, expect, it } from 'vitest'
import { readFileSync } from 'fs'
import { parseIntimationNotice, parseIntimationNoticeText } from '../core/intimationNotice'

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

  it('parses the printed Intimation / Letter of Acceptance PDF text', () => {
    // The reconstructed text lines pdfToTextLines produces for the office's
    // printed Intimation letter (a trimmed copy of the real one).
    const lines = [
      'CYBERABAD MUNICIPAL CORPORATION',
      'Lr. No: /EE/ Nizampet/CIR-58/QBZ/CMC/2026-27Date:15.07.2026',
      'I N T I M A T I O N',
      'To,',
      'M V S CONSTRUCTIONS',
      '13/B, Allwyn Colony, Phase 2, Kukatpally, Hyderabad -500072, Telangana',
      'Phone No:',
      'Sub: -CMC – Quthbullapur Zone – Nizampet – Circle-58 –Works “Junction Improvement…”',
      'Ref : 1). E-Proc Nit No : 12/DB/EE/Nizampet Circle-58/CMC/2026-27',
      '2) Your Tender price bid opened Date: 15.07.2026',
      '3).Amount of Estimate: Rs.20.00 Lakhs',
      'You are hereby informed that your tender for the execution of the above mentioned work has been',
      'accepted at (-)11.11% less than the estimated value Rs 1593493.00/-, with a contract value of ₹',
      '1416455.93/- subject to the terms and conditions of the tender.'
    ]
    const r = parseIntimationNoticeText(lines)
    expect(r.agencyName).toBe('M V S CONSTRUCTIONS')
    expect(r.address).toBe('13/B, Allwyn Colony, Phase 2, Kukatpally, Hyderabad -500072, Telangana')
    expect(r.nitNo).toBe('12/DB/EE/Nizampet Circle-58/CMC/2026-27')
    expect(r.ecvRupees).toBe(1593493)
    expect(r.contractRupees).toBe(1416455.93)
  })

  it('stitches a NIT No that wraps across two lines in the printed LOA', () => {
    // Real Gajularamaram LOA: the NIT No wraps mid-code across two lines.
    const lines = [
      'tender.telangana.gov.in/viewIntimationNotice.html',
      'To,',
      'Sri Duggi Parvathalu',
      'Gajularamaram, Hyderabad',
      'This is notify you that the bid submitted by you for execution of the NIT No.02/DB/EE/Gajularamaram',
      'Circle-57/QBZ/CMC/2026-27 Dt.07.05.2026 (Item No.02) at contract price of Rs. 83008.32 ( Eighty',
      'Three Thousand Eight and Thirty Two paise only )'
    ]
    const r = parseIntimationNoticeText(lines)
    expect(r.nitNo).toBe('02/DB/EE/Gajularamaram Circle-57/QBZ/CMC/2026-27')
  })

  it('parses a printed LOA that carries the NIT code with no "NIT No" label at all', () => {
    // Real GHMC Nizampet Circle-58 "View Intimation Notice" letter, printed to
    // PDF: "…for execution of the E1/06/11/DB/EE/Nizampet Circle-58/CMC/2026-27,
    // dt: 18.06.2026 at contract price of Rs. 806133.90 (…)". No "NIT No"
    // label, an "E1/06/" item-number prefix, "contract price of" (not "contract
    // value of"), and the ECV sits only in the summary table, not in a sentence.
    // Reported bug: the tender fields came out wrong/blank in the Intimation
    // and Agreement workspace for this letter format.
    const lines = [
      'DATE: Wednesday, July 01, 2026',
      'To',
      'SP CONSTRCUTIONS',
      'PLOT NO 113,JANGIDIPURAM COLONY,WANAPARTHY',
      'HYDERABAD -509103',
      'Telangana',
      'Sir/Madam,',
      'This is notify you that the bid submitted by you for execution of the E1/06/11/DB/EE/Nizampet Circle-',
      '58/CMC/2026-27, dt: 18.06.2026 at contract price of Rs. 806133.90 ( Eight Lakh Six Thousand One',
      'Hundred and Thirty Three Rupees Ninety Paisa) as corrected and modified in accordance with the',
      'instructions to the bidders is here by considered as successful bid .',
      'Company Name Estimated Contract Value Corpus Fund @ 0.04 %',
      'SP CONSTRCUTIONS 1033505.00 414.00',
      'Yours Faithfully'
    ]
    const r = parseIntimationNoticeText(lines)
    expect(r.agencyName).toBe('SP CONSTRCUTIONS')
    expect(r.nitNo).toBe('E1/06/11/DB/EE/Nizampet Circle-58/CMC/2026-27')
    expect(r.nitDate).toBe('18.06.2026')
    expect(r.ecvRupees).toBe(1033505)
    expect(r.contractRupees).toBe(806133.9)
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
