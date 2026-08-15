import { describe, expect, it } from 'vitest'
import { parseBalanceEmdReceipt } from '../core/balanceEmdReceipt'

// The reconstructed text lines pdfToTextLines produces for a real CURE portal
// "Balance EMD payment Receipt" PDF — note "Name of the work" wraps across
// several rows with its own label landing mid-wrap (a tall-row quirk of the
// source PDF), and the ECV prints with a stray extra comma
// ("7,7,70,240.00") that toNumber must still parse correctly once commas are
// stripped.
const LINES = [
  'CORE URBAN REGION ECONOMY (CURE)',
  'Administration Department',
  'Balance EMD payment Receipt',
  'Receipt No. : 3859427072026144459848',
  'TenderID : 717465',
  'Tender Notice No. : 14/SE/QBZ/CMC/2026-27, Dated.17.07.2026',
  'Zone Name : Quthbullapur',
  'Circle Name : 58-Nizampet',
  ': Transportation of Garbage from Nizampet',
  'Circle Transfer station to jawahar nagar Transfer',
  'Name of the work',
  'station from July to September in Nizampet',
  'circle-58, Quthbullapur Zone, CMC',
  'Type of the work : General Work',
  'Name of the Agency : PATNAM SRISAILAM',
  'In Favour of : CURE',
  'Estimated Contract Value : Rs. 7,7,70,240.00',
  'Tender Percentage : -7.67 %',
  '1.5% / 2.5% Balance EMD : Rs. 1,16,554.00',
  'ASD : Rs. .00',
  'Total Amount : Rs. 1,16,554.00',
  'Payment ID : pay_TITRIQJMJueKuX',
  'Date of Payment : 27-07-2026',
  'Transaction Status : SUCCESS',
  '*It is a computerized Receipt, signature not required.',
  '*Note: Payment gateway charges levied separately.'
]

describe('parseBalanceEmdReceipt', () => {
  it('extracts the receipt/tender identifiers', () => {
    const r = parseBalanceEmdReceipt(LINES)
    expect(r.receiptNo).toBe('3859427072026144459848')
    expect(r.tenderId).toBe('717465')
    expect(r.paymentId).toBe('pay_TITRIQJMJueKuX')
  })

  it('splits the Tender Notice No. from its own trailing date', () => {
    const r = parseBalanceEmdReceipt(LINES)
    expect(r.noticeNo).toBe('14/SE/QBZ/CMC/2026-27')
    expect(r.noticeDate).toBe('17.07.2026')
  })

  it('extracts Zone/Circle and the agency name', () => {
    const r = parseBalanceEmdReceipt(LINES)
    expect(r.zone).toBe('Quthbullapur')
    expect(r.circle).toBe('58-Nizampet')
    expect(r.agencyName).toBe('PATNAM SRISAILAM')
  })

  it('reassembles the wrapped Name of the work, dropping the label that lands mid-wrap', () => {
    const r = parseBalanceEmdReceipt(LINES)
    expect(r.nameOfWork).toBe(
      'Transportation of Garbage from Nizampet Circle Transfer station to jawahar nagar Transfer station from July to September in Nizampet circle-58, Quthbullapur Zone, CMC'
    )
  })

  it('parses the amounts, stripping the extra comma in the ECV', () => {
    const r = parseBalanceEmdReceipt(LINES)
    expect(r.ecvRupees).toBe(7770240)
    expect(r.tenderPercentage).toBe(-7.67)
    expect(r.balanceEmdRupees).toBe(116554)
    expect(r.totalRupees).toBe(116554)
  })

  it('reads a blank ASD amount ("Rs. .00") as zero, not undefined', () => {
    expect(parseBalanceEmdReceipt(LINES).asdRupees).toBe(0)
  })

  it('extracts the payment date and SUCCESS status', () => {
    const r = parseBalanceEmdReceipt(LINES)
    expect(r.paymentDate).toBe('27-07-2026')
    expect(r.status).toBe('SUCCESS')
  })

  it('returns an empty object for unrelated text', () => {
    expect(parseBalanceEmdReceipt(['Nothing here', 'Just some other document'])).toEqual({})
  })
})
