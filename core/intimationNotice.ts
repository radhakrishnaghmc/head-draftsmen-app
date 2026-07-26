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
 * Extract just the agency name + postal address from an intimation notice
 * rendered as *text lines* — the same "To / <agency> / <address…> / Sir,
 * Madam" block, but from a PDF printout's extracted lines (see
 * src/pdfToText.ts) rather than HTML markup. Used when the intimation is
 * saved as a .pdf instead of a .html. Tolerant of the "To" label being on
 * its own line or inline with the agency name, and stops the address block
 * at the "Sir/Madam" salutation (or after a few lines, as a safety cap).
 */
export function parseIntimationNoticeLines(lines: string[]): { agencyName?: string; address?: string } {
  const rows = lines.map((l) => l.trim()).filter(Boolean)
  let toIdx = -1
  let inlineAgency = ''
  for (let i = 0; i < rows.length; i++) {
    const m = /^to\b[.,:]?\s*(.*)$/i.exec(rows[i])
    if (m) {
      toIdx = i
      inlineAgency = m[1].trim()
      break
    }
  }
  if (toIdx === -1) return {}

  const block: string[] = []
  if (inlineAgency) block.push(inlineAgency)
  for (let i = toIdx + 1; i < rows.length && block.length < 8; i++) {
    if (/^(sir|madam|sir\s*\/?\s*madam|dear)\b/i.test(rows[i])) break
    block.push(rows[i])
  }
  if (block.length === 0) return {}
  return {
    agencyName: block[0],
    address: block.length > 1 ? block.slice(1).join(', ') : undefined
  }
}
