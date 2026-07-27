export interface IntimationNotice {
  /** Contractor/company name, from the "To" address block. */
  agencyName?: string
  /** The agency's postal address — the address-block lines below the name, joined with ", ". */
  address?: string
  /** NIT / e-Proc tender notice number (without the trailing "item N"/"Dated" tail). */
  nitNo?: string
  /** Estimated Contract Value, in rupees (the portal lists it in rupees, not Lakhs). */
  ecvRupees?: number
  /** Accepted contract value, in rupees. */
  contractRupees?: number
}

function decodeEntities(s: string): string {
  return s
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&nbsp;/g, ' ')
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
}

function clean(s: string): string {
  return decodeEntities(s.replace(/<[^>]+>/g, ' ')).replace(/\s+/g, ' ').trim()
}

function toNumber(s: string): number | undefined {
  const n = Number(s.replace(/,/g, '').trim())
  return Number.isFinite(n) ? n : undefined
}

/**
 * Parses the Telangana e-procurement portal's "View Intimation Notice" page
 * (saved as .html) for the fields the Intimation letter needs that aren't
 * already on the Works List — above all the agency's postal address, which
 * lives only here. Everything is best-effort against that page's specific
 * markup; any field it can't find is left undefined (the caller falls back
 * to the Works List row, and every field is editable regardless).
 *
 * The address block is a single <p> of the shape:
 *   To <br> M V S CONSTRUCTIONS <br> 13/B, Allwyn Colony… <br> Hyderabad… <br> Telangana
 * — the first line after "To" is the agency name, the rest are the address.
 */
export function parseIntimationNotice(html: string): IntimationNotice {
  const result: IntimationNotice = {}

  // Address block: the <p> that starts with "To" followed by <br>-separated lines.
  const toBlock = /<p[^>]*>\s*To\b([\s\S]*?)<\/p>/i.exec(html)
  if (toBlock) {
    const lines = toBlock[1]
      .split(/<br\s*\/?>/i)
      .map((l) => clean(l))
      .filter((l) => l.length > 0)
    if (lines.length > 0) result.agencyName = lines[0]
    if (lines.length > 1) result.address = lines.slice(1).join(', ')
  }

  // NIT No. — captured up to the "item"/"Dated" tail that follows it.
  const nit = /NIT\s*No\.?\s*([\s\S]*?)\s*(?:item\b|Dated\b|<\/b>)/i.exec(html)
  if (nit) {
    const value = clean(nit[1])
    if (value) result.nitNo = value
  }

  // Accepted contract value: "contract price of Rs. … <b> 1416455.93 ( … )".
  const contract = /contract\s+price\s+of\s+Rs\.?[\s\S]*?<b>\s*([\d,]+(?:\.\d+)?)/i.exec(html)
  if (contract) result.contractRupees = toNumber(contract[1])

  // Estimated Contract Value: the data cell under the "Estimated Contract
  // Value" column of the summary table (the row's 2nd cell, after the
  // company-name cell).
  const tbody = /<tbody>([\s\S]*?)<\/tbody>/i.exec(html)
  if (tbody) {
    const cells = [...tbody[1].matchAll(/<td[^>]*>([\s\S]*?)<\/td>/gi)].map((m) => clean(m[1]))
    if (cells.length >= 2) result.ecvRupees = toNumber(cells[1])
  }

  return result
}

/**
 * Parses the *printed* Intimation / "Letter of Acceptance" PDF (the office's
 * own letter, saved as a text-bearing PDF) for the same fields
 * parseIntimationNotice pulls from the portal HTML — so the Give Intimation
 * and Work Order / Agreement tabs accept the Online Intimation as either an
 * .html portal page or a .pdf letter. Input is the reconstructed text lines
 * (see src/pdfToText.ts's pdfToTextLines). The letter's layout:
 *
 *   To,
 *   M V S CONSTRUCTIONS                          <- agency name
 *   13/B, Allwyn Colony, … Telangana             <- address (until "Phone No"/"Sub")
 *   ...
 *   Ref : 1). E-Proc Nit No : 12/DB/EE/…/2026-27 <- NIT No
 *   ... accepted at (-)11.11% less than the estimated value Rs 1593493.00/-,
 *       with a contract value of ₹ 1416455.93/- ...
 *
 * Each field is best-effort; anything not found is left undefined (the caller
 * falls back to the Works List row, and every field stays editable).
 */
export function parseIntimationNoticeText(lines: string[]): IntimationNotice {
  const result: IntimationNotice = {}
  const joined = lines.join(' ').replace(/\s+/g, ' ').trim()

  // Agency name + address from the "To," block: the first line after "To,"
  // is the agency, the lines below it (until "Phone No"/"Sub:"/"Ref:") are
  // the postal address.
  const toIdx = lines.findIndex((l) => /^to\s*[,:]?\s*$/i.test(l))
  if (toIdx >= 0) {
    const after = lines.slice(toIdx + 1).map((l) => l.trim()).filter((l) => l.length > 0)
    if (after.length > 0) result.agencyName = after[0]
    const addr: string[] = []
    for (const l of after.slice(1)) {
      if (/^(phone\s*no|sub\s*[:.]|ref\s*[:.]|sir|madam)/i.test(l)) break
      addr.push(l)
    }
    if (addr.length > 0) result.address = addr.join(', ')
  }

  // NIT No — the "…Nit No : <code>" on the Ref line; capture the code up to a
  // following "2)" ref, a "Date:" tail, or the end of that line.
  const nitLine = lines.find((l) => /Nit\s*No/i.test(l)) ?? ''
  const nit = /Nit\s*No\.?\s*:?\s*(.+?)\s*(?:\d\s*\)|Your\s+Tender|Date\s*:|$)/i.exec(nitLine)
  if (nit) {
    const value = nit[1].replace(/\s+/g, ' ').trim()
    if (value) result.nitNo = value
  }

  // ECV — "…estimated value Rs 1593493.00…" (may drop the paise).
  const ecv = /estimated\s+value\s+(?:of\s+)?Rs\.?\s*([\d,]+(?:\.\d+)?)/i.exec(joined)
  if (ecv) result.ecvRupees = toNumber(ecv[1])

  // Accepted contract value — "…contract value of ₹ 1416455.93…" (the "₹"/
  // "Rs." and the amount often land on different lines, so read from joined).
  const contract = /contract\s+value\s+of\s*[₹Rs.\s]*([\d,]+(?:\.\d+)?)/i.exec(joined)
  if (contract) result.contractRupees = toNumber(contract[1])

  return result
}
