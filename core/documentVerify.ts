/**
 * No-AI, rule-based cross-check for a generated document: does the text that
 * actually got baked into the .docx still match what the app currently
 * computes from its source data? Catches blank/missing fields, leftover
 * {{placeholders}}, EMD/ASD arithmetic that doesn't tie out, a wrong/leftover
 * corporation name, and reserved-tag inconsistency — the classes of mistake a
 * human proofreader would look for, without sending the document anywhere.
 */

export interface VerifyIssue {
  field: string
  message: string
}

export interface VerifyResult {
  ok: boolean
  issues: VerifyIssue[]
}

function normalizeText(s: string): string {
  return s.replace(/\s+/g, ' ').trim()
}

/** Every non-empty resolved value must appear verbatim (whitespace-normalized) in the document text; every value the caller marks required must be non-empty. */
export function verifyPlaceholderCoverage(
  docText: string,
  values: Record<string, string>,
  requiredLabels: string[] = Object.keys(values)
): VerifyIssue[] {
  const issues: VerifyIssue[] = []
  const haystack = normalizeText(docText)
  if (/\{\{[^}]+\}\}/.test(docText)) {
    const leftover = [...new Set(docText.match(/\{\{[^}]+\}\}/g) ?? [])]
    for (const ph of leftover) issues.push({ field: ph, message: `Unfilled placeholder ${ph} left in the document` })
  }
  for (const label of requiredLabels) {
    const value = normalizeText(values[label] ?? '')
    if (!value) {
      issues.push({ field: label, message: `${label} is blank` })
      continue
    }
    if (!haystack.includes(value)) {
      issues.push({ field: label, message: `${label} ("${values[label]}") not found in the generated document` })
    }
  }
  return issues
}

/** Recomputes ECV/Contract/EMD/ASD independently from source figures and confirms they tie out with what's in the document — catches a wrong source cell or a stale re-download, not just a formula bug (the formula is re-derived here, not imported). */
export function verifyAmountMath(opts: {
  docText: string
  ecv?: number | null
  contract?: number | null
  tenderPct?: number | null
  // e.g. 0.015 or 0.025 — omitted for documents whose text never prints a
  // computed EMD/ASD figure (only ECV/Contract), so there's nothing to find.
  emdPct?: number
}): VerifyIssue[] {
  const issues: VerifyIssue[] = []
  const { docText, ecv, contract, tenderPct, emdPct } = opts
  const haystack = normalizeText(docText)
  const rupeeGroup = (n: number) => Math.round(n).toLocaleString('en-IN')

  if (ecv != null && contract != null && tenderPct != null) {
    const expectedContract = ecv * (1 - tenderPct / 100)
    if (Math.abs(expectedContract - contract) > Math.max(2, ecv * 0.0005)) {
      issues.push({
        field: 'Contract Amount',
        message: `Tender Contract Value Rs.${rupeeGroup(contract)} doesn't match ECV Rs.${rupeeGroup(ecv)} less ${tenderPct}% (expected ~Rs.${rupeeGroup(expectedContract)})`
      })
    }
  }
  if (ecv != null && emdPct != null) {
    const emd = Math.round(ecv * emdPct)
    if (!haystack.includes(rupeeGroup(emd))) {
      issues.push({ field: 'EMD', message: `Expected EMD Rs.${rupeeGroup(emd)} (${emdPct * 100}% of ECV) not found in the document` })
    }
    if (tenderPct != null && tenderPct > 25) {
      const asd = Math.round((ecv * (tenderPct - 25)) / 100)
      if (!haystack.includes(rupeeGroup(asd))) {
        issues.push({ field: 'ASD', message: `Expected ASD Rs.${rupeeGroup(asd)} (${(tenderPct - 25).toFixed(2)}% of ECV) not found in the document` })
      }
    }
  }
  return issues
}

/** Flags any *other* known corporation's name showing up in the text — a leftover from a previous office/template (e.g. "GHMC" baked into an MMC letter). */
export function verifyCorporationWording(docText: string, expectedCorporation: string, allCorporations: string[]): VerifyIssue[] {
  const issues: VerifyIssue[] = []
  for (const corp of allCorporations) {
    if (corp === expectedCorporation) continue
    const re = new RegExp(`\\b${corp}\\b`)
    if (re.test(docText)) {
      issues.push({ field: 'Corporation', message: `Document mentions "${corp}" but the selected office's corporation is "${expectedCorporation}"` })
    }
  }
  return issues
}

/** Confirms the reserved-work tag is present when the work is reserved, and absent when it isn't. */
export function verifyReservedTag(docText: string, isReserved: boolean): VerifyIssue[] {
  const mentionsReserved = /reserved for\s+\S+/i.test(docText)
  if (isReserved && !mentionsReserved) {
    return [{ field: 'Reserved Tag', message: 'Work is marked reserved but no "(Reserved for ...)" tag was found in the document' }]
  }
  if (!isReserved && mentionsReserved) {
    return [{ field: 'Reserved Tag', message: 'Work is not marked reserved but the document mentions "Reserved for ..."' }]
  }
  return []
}

export function combineVerify(...issueLists: VerifyIssue[][]): VerifyResult {
  const issues = issueLists.flat()
  return { ok: issues.length === 0, issues }
}
