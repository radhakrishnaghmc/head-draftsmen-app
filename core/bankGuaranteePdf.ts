/**
 * Parses a Bank Guarantee (BG) certificate — uploaded in place of an online
 * EMD 1.5%/2.5% or ASD payment receipt when the agency furnishes a BG
 * instead of paying online. Different banks issue BGs on very different
 * letterheads/layouts, and a single upload commonly bundles two separate BG
 * certificates back to back (one for EMD, one for ASD) with no explicit
 * label saying which is which — so this only pulls the handful of details
 * that are printed on every bank's BG regardless of layout (BG No, issue
 * date, guaranteed amount) and leaves matching each one to EMD vs ASD to the
 * caller, which knows the work's expected amounts (see
 * src/components/WorkOrderAgreementTab.tsx).
 */
export interface BankGuaranteeInfo {
  bgNo?: string
  /** dd.mm.yyyy / dd/mm/yyyy / dd-mm-yyyy as printed. */
  issueDate?: string
  amountRupees?: number
}

function toNumber(s: string): number | undefined {
  const n = Number(s.replace(/,/g, '').trim())
  return Number.isFinite(n) ? n : undefined
}

// Captures the rest of the line: some banks wrap the code across a stray
// space ("BG No: 14141 GI3D3244026") that a plain word-boundary capture
// would truncate at.
const BG_NO_RE = /\bB\.?\s*G\.?\s*No\.?\s*:?\s*(.+)/i
const DATE_RE = /\b(\d{1,2}[.\-/]\d{1,2}[.\-/]\d{2,4})\b/
const AMOUNT_RE = /Rs\.?\s*([\d,]+(?:\.\d{1,2})?)/i
// Preferred over a bare "Rs." figure, which can pick up an unrelated amount
// (a stamp duty, a reference number formatted like money, etc.) elsewhere on
// the certificate — these phrasings are the guarantee's own operative sum
// clause and are close to universal across bank BG formats.
const AMOUNT_PREFERRED_RES = [
  /(?:shall not exceed|not exceeding|liability.{0,20}exceed)\s*Rs\.?\s*([\d,]+(?:\.\d{1,2})?)/i,
  /(?:total of|sum of|value of|guarantee[d]?\s*(?:for|amount))\s*Rs\.?\s*([\d,]+(?:\.\d{1,2})?)/i
]

/**
 * One entry per distinct BG No found. Each "BG No:" line anchors a search
 * for a nearby date and a forward window (the certificate body that follows
 * the header) for the guaranteed amount. A BG's header/amount is often
 * repeated verbatim on a later page of the same certificate (renewal
 * schedule, T&C page) — those repeats are collapsed into the first sighting
 * by amount (a steadier match key than the BG No text, which scanned/OCR'd
 * certificates sometimes render slightly differently on each repeat).
 */
export function parseBankGuarantees(lines: string[]): BankGuaranteeInfo[] {
  const found: BankGuaranteeInfo[] = []

  for (let idx = 0; idx < lines.length; idx++) {
    const m = BG_NO_RE.exec(lines[idx])
    if (!m) continue
    const bgNo = m[1].replace(/\s+/g, '').trim()
    // A real BG No is a long alphanumeric code. A short capture means the
    // value got split across a page-reflow (the rest landing on an
    // unrelated earlier/later line) — too unreliable to trust, so this
    // sighting is skipped rather than surfaced half-wrong.
    if (bgNo.length < 6) continue

    let issueDate: string | undefined
    const sameLine = DATE_RE.exec(lines[idx])
    if (sameLine) issueDate = sameLine[1]
    for (let d = 1; d <= 6 && !issueDate; d++) {
      const near = lines[idx + d]
      if (near && /issue|dated|date/i.test(near)) {
        const dm = DATE_RE.exec(near)
        if (dm) issueDate = dm[1]
      }
    }

    const windowEnd = Math.min(lines.length, idx + 80)
    let amountRupees: number | undefined
    for (const re of AMOUNT_PREFERRED_RES) {
      for (let i = idx; i < windowEnd && amountRupees === undefined; i++) {
        const am = re.exec(lines[i])
        if (am) amountRupees = toNumber(am[1])
      }
      if (amountRupees !== undefined) break
    }
    if (amountRupees === undefined) {
      for (let i = idx; i < windowEnd; i++) {
        const am = AMOUNT_RE.exec(lines[i])
        if (am) {
          amountRupees = toNumber(am[1])
          break
        }
      }
    }

    found.push({ bgNo, issueDate, amountRupees })
  }

  // Collapse repeats of the same certificate (same amount, printed again on
  // a later page) — a real second BG in the same upload has its own,
  // different guaranteed amount.
  const result: BankGuaranteeInfo[] = []
  for (const bg of found) {
    const dup = bg.amountRupees !== undefined && result.some((r) => r.amountRupees === bg.amountRupees)
    if (!dup) result.push(bg)
  }
  return result
}

// Standard IBA bank-guarantee boilerplate ("And whereas we Union Bank of
// India, a Body Corporate...", "We, State Bank of India, a body corporate
// constituted...") is close to identical across issuing banks regardless of
// letterhead layout, so this one anchor covers most uploads. Best-effort —
// left undefined (the composed EMD Details clause just omits "of <bank>")
// rather than risk printing the wrong bank's name.
const BANK_NAME_RES = [
  /\bwe,?\s+([A-Z][A-Za-z.&' ]{2,60}?Bank[A-Za-z.&' ]{0,20}?),?\s+a\s+(?:body corporate|banking company|company)/i
]

export function extractBankName(lines: string[]): string | undefined {
  const text = lines.join(' ')
  for (const re of BANK_NAME_RES) {
    const m = re.exec(text)
    if (m) return m[1].trim().replace(/\s+/g, ' ')
  }
  return undefined
}
