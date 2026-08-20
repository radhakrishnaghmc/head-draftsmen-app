export interface IntimationNotice {
  /** Contractor/company name, from the "To" address block. */
  agencyName?: string
  /** The agency's postal address — the address-block lines below the name, joined with ", ". */
  address?: string
  /** NIT / e-Proc tender notice number (without the trailing "item N"/"Dated" tail). */
  nitNo?: string
  /** The NIT's own date, from the "…Dated:24.07.2026" tail right after the NIT No. */
  nitDate?: string
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

  // NIT's own date, in the same "…item N Dated:24.07.2026" tail the NIT No
  // capture above stops before. Searched in a short window right after "NIT
  // No" (not the whole page) so an unrelated "Dated" further down the letter
  // is never picked up instead.
  const nitIdx = html.search(/NIT\s*No\b/i)
  if (nitIdx >= 0) {
    const window = clean(html.slice(nitIdx, nitIdx + 400))
    const nitDate = /(?<![A-Za-z])(?:Dated|Dt)\b\.?\s*:?\s*(\d{1,2}[.\-/]\d{1,2}[.\-/]\d{2,4})/i.exec(window)
    if (nitDate) result.nitDate = nitDate[1]
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

  // NIT No — the "…Nit No[.:] <code>" line; capture the code up to a following
  // "2)" ref, a "Date:"/"Dt"/"Item" tail, "at contract", or the line's end.
  // In the printed LOA the code wraps onto the next line ("…Gajularamaram" then
  // "Circle-57/QBZ/CMC/2026-27 Dt.…"), so when the code on the NIT line alone
  // looks truncated (no 20XX-XX year yet), stitch the following line first.
  // Also stitch when the NIT line has nothing at all after "NIT No" (e.g.
  // "…for execution of the NIT No" ends the line, the code starts on the
  // next) — checking `!value` alone, not `value &&`, would otherwise skip
  // the stitch precisely when it's needed most.
  const nitIdx = lines.findIndex((l) => /Nit\s*No/i.test(l))
  if (nitIdx >= 0) {
    const nitRe = /Nit\s*No\.?\s*:?\s*(.+?)\s*(?:\d\s*\)|Your\s+Tender|Date\s*:|Dt\b|Item\b|at\s+contract|$)/i
    const capture = (text: string): string => nitRe.exec(text)?.[1].replace(/\s+/g, ' ').trim() ?? ''
    let value = capture(lines[nitIdx])
    if (!/20\d\d\s*-\s*\d{2}/.test(value) && lines[nitIdx + 1]) {
      const stitched = capture(`${lines[nitIdx]} ${lines[nitIdx + 1]}`)
      if (stitched) value = stitched
    }
    if (value) result.nitNo = value

    // NIT's own date, e.g. "…ITEM 7Dated:24.07.2026" right after the NIT No —
    // read from the same one/two-line window used above (not the whole
    // letter) so an unrelated "Dated" further down the body — e.g. a GO
    // reference's own date — is never picked up instead. No \b before the
    // anchor: the item number often runs straight into "Dated" with no space,
    // which a word boundary would reject (digit-into-letter isn't a
    // boundary) — the lookbehind instead just rules out landing inside
    // another word ("Updated"/"Mandated").
    const dateWindow = `${lines[nitIdx]} ${lines[nitIdx + 1] ?? ''}`
    const nitDate = /(?<![A-Za-z])(?:Dated|Dt)\b\.?\s*:?\s*(\d{1,2}[.\-/]\d{1,2}[.\-/]\d{2,4})/i.exec(dateWindow)
    if (nitDate) result.nitDate = nitDate[1]
  }

  // Some "View Intimation Notice" letters carry the code with no "NIT No"
  // label at all — "…for execution of the E1/06/11/DB/EE/Nizampet
  // Circle-58/CMC/2026-27, dt: 18.06.2026 at contract price of Rs. …". Fall
  // back to the code's own canonical shape (same as the L1 sheet's fallback
  // in tenderEvaluationPdf.ts): "<code>/DB/EE/<place> Circle-<circleNo>/
  // [QBZ/]CMC/<year>", with an optional 1-2 segment item-number prefix
  // ("E1/06/") each segment carrying a digit so it can't swallow a label word.
  if (!result.nitNo) {
    const codePrefix = '(?:[A-Za-z]{0,2}\\d{1,3}\\/){0,2}'
    const canonical = new RegExp(
      `${codePrefix}\\d{1,3}\\/DB\\/EE\\/.+?Circle\\s*-\\s*\\d{1,3}\\/(?:QBZ\\/)?CMC\\/\\d{4}\\s*-\\s*\\d{2,4}`,
      'i'
    ).exec(joined)
    if (canonical) {
      result.nitNo = canonical[0].replace(/\s*([/\-])\s*/g, '$1').replace(/\s+/g, ' ').trim()
      // The date right after the code — "…2026-27, dt: 18.06.2026 at contract…".
      const window = joined.slice(canonical.index, canonical.index + canonical[0].length + 60)
      const date = /(?<![A-Za-z])(?:Dated|Dt)\b\.?\s*:?\s*(\d{1,2}[.\-/]\d{1,2}[.\-/]\d{2,4})/i.exec(window)
      if (date) result.nitDate = date[1]
    }
  }

  // ECV — "…estimated value Rs 1593493.00…" (may drop the paise).
  const ecv = /estimated\s+value\s+(?:of\s+)?Rs\.?\s*([\d,]+(?:\.\d+)?)/i.exec(joined)
  if (ecv) result.ecvRupees = toNumber(ecv[1])
  // Fallback: the letters that skip the "estimated value" sentence instead
  // carry the ECV in a summary table — "Company Name Estimated Contract
  // Value Corpus Fund @ 0.04 %" then "<agency> <ecv> <corpus>" on the next
  // line. Take the first of the row's two trailing amounts (the ECV; the
  // second is the corpus fund).
  if (result.ecvRupees == null) {
    const headerIdx = lines.findIndex((l) => /Estimated\s+Contract\s+Value/i.test(l) && /Corpus\s+Fund/i.test(l))
    if (headerIdx >= 0) {
      const row = /([\d,]+\.\d{2})\s+([\d,]+\.\d{2})\s*$/.exec(lines[headerIdx + 1]?.trim() ?? '')
      if (row) result.ecvRupees = toNumber(row[1])
    }
  }

  // Accepted contract value — "…contract value of ₹ 1416455.93…" (the "₹"/
  // "Rs." and the amount often land on different lines, so read from joined)
  // — or the shorter LOA wording "…at contract price of Rs. 806133.90 ( …)".
  const contract = /contract\s+(?:value|price)\s+of\s*[₹Rs.\s]*([\d,]+(?:\.\d+)?)/i.exec(joined)
  if (contract) result.contractRupees = toNumber(contract[1])

  return result
}
