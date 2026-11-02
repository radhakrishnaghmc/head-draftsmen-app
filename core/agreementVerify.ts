import type { VerifyIssue } from './documentVerify'
import type { MonitoringFormatWorkRow } from './monitoringFormat'
import { findMonitoringWork } from './massNotices'
import { normalizeWorkNameForMatch } from './tenderAgents/nameOfWork'

/**
 * Agreement-workspace Verify: cross-checks the figures printed in every filled document against their sources —
 * the L-1 sheet (name of work, ECV, contract, tender %), the Works List database (estimate amount) and the
 * Monitoring Format. No AI; it only reports. Document-level checks run per document on its extracted text,
 * work-level checks run once on the source figures.
 */

/** What a filled document must carry, expressed as source figures (never read back from the document itself). */
export interface DocFigures {
  /** Name of work as the L-1 sheet gives it. */
  workName: string
  /** Amount of estimate in Lakhs, from the Works List database row for this work. */
  estimateLakhs: number | null
  ecv: number | null
  contract: number | null
  tenderPct: number | null
  wincode: string
  completionMonths: string
  /** Reserved works are EMD-exempt (documents print "Exempted"), but still owe ASD. */
  reserved: boolean
}

const LABEL = {
  workName: /^name (of )?(the )?work$/i,
  estimate: /^estimate amount$/i,
  ecv: /^ecv$/i,
  contract: /^contract amount$/i,
  tenderPct: /^tender percentage$/i,
  wincode: /^wincode$/i,
  period: /^period (of )?completion$/i,
  // A bare "EMD" placeholder is the 1% figure (Contract Deed etc.).
  emd1: /^(emd|emd 1 percent|emd 1%)$/i,
  emd1_5: /^(emd 1\.5 percent|emd 1\.5%)$/i,
  asd: /^asd( details)?$/i
}

const norm = (s: string) => s.replace(/\s+/g, ' ').trim()
const has = (labels: string[], re: RegExp) => labels.some((l) => re.test(l.trim()))

/** "1,23,456" — the Indian-grouped integer part, which every document spelling of a rupee amount contains. */
function indian(n: number): string {
  return Math.round(n).toLocaleString('en-IN')
}

function containsAmount(haystack: string, n: number): boolean {
  return haystack.includes(indian(n)) || haystack.includes(indian(Math.floor(n)))
}

/**
 * Per-document checks, limited to the figures the document's template actually has a placeholder for
 * (`labels`) — so a document that never prints an estimate isn't flagged for lacking one.
 */
export function verifyDocFigures(docText: string, labels: string[], f: DocFigures): VerifyIssue[] {
  const issues: VerifyIssue[] = []
  const hay = norm(docText)
  const hayName = normalizeWorkNameForMatch(docText)

  if (has(labels, LABEL.workName)) {
    const want = normalizeWorkNameForMatch(f.workName)
    if (!want) issues.push({ field: 'Name of Work', message: 'Name of work is blank (not on the L-1 sheet)' })
    else if (!hayName.includes(want))
      issues.push({ field: 'Name of Work', message: 'Name of work in the document differs from the L-1 sheet' })
  }
  if (has(labels, LABEL.estimate)) {
    if (f.estimateLakhs == null)
      issues.push({ field: 'Estimate', message: 'Amount of Estimate is blank (not in the Works List database)' })
    else if (!hay.includes(f.estimateLakhs.toFixed(2)))
      issues.push({
        field: 'Estimate',
        message: `Estimate amount ${f.estimateLakhs.toFixed(2)} Lakhs (Works List database) not found in the document`
      })
  }
  if (has(labels, LABEL.ecv)) {
    if (f.ecv == null) issues.push({ field: 'ECV', message: 'ECV is blank (not on the L-1 sheet)' })
    else if (!containsAmount(hay, f.ecv))
      issues.push({ field: 'ECV', message: `ECV Rs.${indian(f.ecv)} (L-1 sheet) not found in the document` })
  }
  if (has(labels, LABEL.contract)) {
    if (f.contract == null)
      issues.push({ field: 'Contract Amount', message: 'Contract amount is blank (not on the L-1 sheet)' })
    else if (!containsAmount(hay, f.contract))
      issues.push({
        field: 'Contract Amount',
        message: `Contract amount Rs.${indian(f.contract)} (L-1 sheet) not found in the document`
      })
  }
  if (has(labels, LABEL.tenderPct)) {
    if (f.tenderPct == null)
      issues.push({ field: 'Tender %', message: 'Tender percentage is blank (not on the L-1 sheet)' })
    else {
      const abs = Math.abs(f.tenderPct)
      if (f.tenderPct !== 0 && !hay.includes(abs.toFixed(2)) && !hay.includes(String(abs)))
        issues.push({
          field: 'Tender %',
          message: `Tender percentage ${abs.toFixed(2)}% (L-1 sheet) not found in the document`
        })
    }
  }
  if (has(labels, LABEL.wincode)) {
    const w = f.wincode.trim()
    if (!w) issues.push({ field: 'Wincode', message: 'Wincode is not available' })
    else if (!hay.includes(w)) issues.push({ field: 'Wincode', message: `Wincode ${w} not found in the document` })
  }
  // EMD 1% / 1.5% of ECV, and ASD ((tender % - 25)% of ECV) once the tender exceeds 25% — recomputed here, not read back.
  if (f.ecv != null && !f.reserved) {
    if (has(labels, LABEL.emd1) && !containsAmount(hay, f.ecv * 0.01))
      issues.push({ field: 'EMD 1%', message: `EMD @1% Rs.${indian(f.ecv * 0.01)} (of ECV) not found in the document` })
    if (has(labels, LABEL.emd1_5) && !containsAmount(hay, f.ecv * 0.015))
      issues.push({
        field: 'EMD 1.5%',
        message: `EMD @1.5% Rs.${indian(f.ecv * 0.015)} (of ECV) not found in the document`
      })
  }
  if (f.ecv != null && f.tenderPct != null && f.tenderPct > 25 && has(labels, LABEL.asd)) {
    const asd = f.ecv * ((f.tenderPct - 25) / 100)
    if (!containsAmount(hay, asd))
      issues.push({
        field: 'ASD',
        message: `ASD Rs.${indian(asd)} (${(f.tenderPct - 25).toFixed(2)}% of ECV) not found in the document`
      })
  }
  if (has(labels, LABEL.period) && !f.completionMonths.trim()) {
    issues.push({ field: 'Period of Completion', message: 'Period of completion is not available' })
  }
  return issues
}

export interface WorkLevelInput extends DocFigures {
  /** The estimate on the Works List row the documents are filling from — compared with the database row found by the L-1's work name. */
  fieldsEstimateLakhs: number | null
  /** Whether the L-1's work was found in the Works List at all. */
  worksRowFound: boolean
  /** EMD 1% / EMD 1.5% / ASD cells on the Works List row, as typed (blank = not entered yet). */
  rowEmdAsd?: { emd1?: string; emd1_5?: string; asd?: string }
  /** Imported Monitoring Format works for this office, if any. */
  monitoringWorks?: MonitoringFormatWorkRow[]
}

/** Once-per-work checks on the source figures. `skipped` lists checks that couldn't run (and why). */
export function verifyWorkFigures(i: WorkLevelInput): { issues: VerifyIssue[]; skipped: string[] } {
  const issues: VerifyIssue[] = []
  const skipped: string[] = []
  const est = i.estimateLakhs != null ? i.estimateLakhs * 100000 : null

  // 1. Estimate: the Works List database vs what every document fills.
  if (!i.worksRowFound) skipped.push('Estimate amount — this work was not found in the Works List database')
  else if (i.estimateLakhs == null)
    issues.push({ field: 'Estimate', message: 'Works List database has no Amount of Estimate for this work' })
  else if (i.fieldsEstimateLakhs == null || Math.abs(i.fieldsEstimateLakhs - i.estimateLakhs) > 0.005)
    issues.push({
      field: 'Estimate',
      message: `Documents are filled with estimate ${i.fieldsEstimateLakhs ?? 'blank'} Lakhs but the Works List database has ${i.estimateLakhs.toFixed(2)} Lakhs for this work`
    })

  // 2/3. ECV, contract, tender % must all be on the L-1 sheet and agree with each other.
  if (i.ecv == null) issues.push({ field: 'ECV', message: 'ECV is not on the L-1 sheet' })
  if (i.contract == null) issues.push({ field: 'Contract Amount', message: 'Contract amount is not on the L-1 sheet' })
  if (i.tenderPct == null) issues.push({ field: 'Tender %', message: 'Tender percentage is not on the L-1 sheet' })
  if (i.ecv != null && i.contract != null && i.tenderPct != null) {
    const expected = i.ecv * (1 - i.tenderPct / 100)
    if (Math.abs(expected - i.contract) > Math.max(2, i.ecv * 0.0005))
      issues.push({
        field: 'Contract Amount',
        message: `Contract Rs.${indian(i.contract)} doesn't match ECV Rs.${indian(i.ecv)} at ${i.tenderPct}% (expected ~Rs.${indian(expected)})`
      })
  }

  // 4. ECV below the estimate; contract at/below ECV for a "Less" tender (at/above it for an "Above" one).
  if (est != null && i.ecv != null && i.ecv >= est)
    issues.push({ field: 'ECV', message: `ECV Rs.${indian(i.ecv)} should be less than the Estimate Rs.${indian(est)}` })
  if (i.ecv != null && i.contract != null && i.tenderPct != null) {
    if (i.tenderPct > 0 && i.contract > i.ecv + 0.005)
      issues.push({
        field: 'Contract Amount',
        message: `Tender is ${i.tenderPct}% Less, so the contract Rs.${indian(i.contract)} should not exceed ECV Rs.${indian(i.ecv)}`
      })
    if (i.tenderPct < 0 && i.contract < i.ecv - 0.005)
      issues.push({
        field: 'Contract Amount',
        message: `Tender is ${Math.abs(i.tenderPct)}% Above, so the contract Rs.${indian(i.contract)} should not be below ECV Rs.${indian(i.ecv)}`
      })
  }

  // EMD @1% / @1.5% and ASD: recomputed from the L-1's ECV / tender %, then compared with the Works List row's own columns.
  if (i.ecv != null) {
    const asd = i.tenderPct != null && i.tenderPct > 25 ? i.ecv * ((i.tenderPct - 25) / 100) : 0
    const want: [string, string | undefined, number][] = [
      ['EMD 1%', i.rowEmdAsd?.emd1, i.ecv * 0.01],
      ['EMD 1.5%', i.rowEmdAsd?.emd1_5, i.ecv * 0.015],
      ['ASD', i.rowEmdAsd?.asd, asd]
    ]
    for (const [label, cell, v] of want) {
      const have = Number((cell ?? '').replace(/[^0-9.]/g, ''))
      if ((cell ?? '').trim() && Number.isFinite(have) && Math.abs(have - v) > 1.5)
        issues.push({
          field: label,
          message: `${label} in the Works List is Rs.${indian(have)} but ${label === 'ASD' ? '(tender % - 25)% of ECV' : 'the % of ECV'} gives Rs.${indian(v)}`
        })
    }
  }

  // 5/7. Wincode and period of completion.
  if (!i.wincode.trim()) issues.push({ field: 'Wincode', message: 'Wincode is not available' })
  if (!i.completionMonths.trim())
    issues.push({ field: 'Period of Completion', message: 'Period of completion is not available' })

  // 8. The work must be in the Monitoring Format.
  if (!i.monitoringWorks || i.monitoringWorks.length === 0)
    skipped.push('Monitoring Format — none imported for this office')
  else if (!findMonitoringWork(i.monitoringWorks, i.workName))
    issues.push({ field: 'Monitoring Format', message: 'Name of work was not found in the Monitoring Format' })

  return { issues, skipped }
}
