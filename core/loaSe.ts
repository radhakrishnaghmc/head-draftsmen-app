/**
 * Superintending-Engineer "Letter of Acceptance" (LOA) intimation — the format
 * used when the office is zone-level (a Zone chosen with no Circle, i.e. the
 * SE / Zonal office rather than an EE / Circle office). It fills a different
 * bundled template (`loa-se-template.docx`) with its own placeholder set and a
 * few SE-specific derivations (LOA letter no, EMD/E-corpus figures, amount in
 * words) plus three manually-entered fields the source data can't supply
 * (Admin Sanction Value, Period of Completion, Item No).
 */

/** Zone → the 3-letter code used on SE documents (e.g. "Quthbullapur" → "QBZ"). */
const ZONE_ABBR: Record<string, string> = {
  quthbullapur: 'QBZ',
  kukatpally: 'KPZ',
  serilingampally: 'SLP'
  // Other CMC/GHMC/MMC zones' official codes are added here as they're confirmed.
}

/**
 * The zone's document code. Known zones use their official code; an unknown
 * zone falls back to its first two letters + "Z" (e.g. "Kukatpally" → "KUZ") so
 * the letter still reads sensibly until the real code is added above.
 */
export function zoneAbbr(zone: string | undefined): string {
  const z = (zone ?? '').trim()
  if (!z) return ''
  const known = ZONE_ABBR[z.toLowerCase()]
  if (known) return known
  return (z.slice(0, 2) + 'Z').toUpperCase()
}

/** Indian financial year ("2025-26") for a dd.mm.yyyy date; blank date → current FY. */
export function financialYearFromDate(ddmmyyyy: string | undefined): string {
  const m = /^(\d{1,2})[./-](\d{1,2})[./-](\d{4})$/.exec((ddmmyyyy ?? '').trim())
  const d = m ? new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1])) : new Date()
  const y = d.getFullYear()
  const startY = d.getMonth() >= 3 ? y : y - 1
  return `${startY}-${String((startY + 1) % 100).padStart(2, '0')}`
}

/** Group an integer with the Indian thousand/lakh/crore comma pattern: 16491978 → "1,64,91,978". */
function groupIndian(intPart: string): string {
  const neg = intPart.startsWith('-')
  const digits = neg ? intPart.slice(1) : intPart
  if (digits.length <= 3) return (neg ? '-' : '') + digits
  const last3 = digits.slice(-3)
  const rest = digits.slice(0, -3)
  const grouped = rest.replace(/\B(?=(\d{2})+(?!\d))/g, ',')
  return (neg ? '-' : '') + grouped + ',' + last3
}

/**
 * Format an amount the way the SE letter shows money:
 *  - decimals 2 → "1,64,91,978.00" (values block)
 *  - decimals 0 → "1,64,920" (EMD / E-corpus, rounded)
 * Returns '' for null/undefined/non-finite.
 */
export function formatIndianAmount(n: number | null | undefined, decimals: 0 | 2 = 2): string {
  if (n == null || !Number.isFinite(n)) return ''
  const fixed = Math.abs(n).toFixed(decimals)
  const [intPart, frac] = fixed.split('.')
  const grouped = groupIndian(intPart)
  const sign = n < 0 ? '-' : ''
  return decimals === 0 ? sign + grouped : sign + grouped + '.' + frac
}

const ONES = [
  '', 'One', 'Two', 'Three', 'Four', 'Five', 'Six', 'Seven', 'Eight', 'Nine', 'Ten',
  'Eleven', 'Twelve', 'Thirteen', 'Fourteen', 'Fifteen', 'Sixteen', 'Seventeen', 'Eighteen', 'Nineteen'
]
const TENS = ['', '', 'Twenty', 'Thirty', 'Forty', 'Fifty', 'Sixty', 'Seventy', 'Eighty', 'Ninety']

/** 0-99 in words ("Eighty Nine" — CMC letters space, not hyphen); 0 → ''. */
function twoDigits(n: number): string {
  if (n < 20) return ONES[n]
  const t = Math.floor(n / 10)
  const o = n % 10
  return o ? `${TENS[t]} ${ONES[o]}` : TENS[t]
}

/** 0-999 in words ("Six Hundred Eighty-Nine"). */
function threeDigits(n: number): string {
  const h = Math.floor(n / 100)
  const rest = n % 100
  const parts: string[] = []
  if (h) parts.push(`${ONES[h]} Hundred`)
  if (rest) parts.push(twoDigits(rest))
  return parts.join(' ')
}

/** Whole number in Indian words the CMC way: 3976612 → "Thirty Nine Lakhs Seventy Six Thousand Six Hundred Twelve" (Lakhs/Crores pluralised as in the letters). */
export function integerToIndianWords(num: number): string {
  if (!Number.isFinite(num) || num <= 0) return num === 0 ? 'Zero' : ''
  let n = Math.floor(num)
  const crore = Math.floor(n / 10000000)
  n %= 10000000
  const lakh = Math.floor(n / 100000)
  n %= 100000
  const thousand = Math.floor(n / 1000)
  n %= 1000
  const parts: string[] = []
  if (crore) parts.push(`${integerToIndianWords(crore)} ${crore > 1 ? 'Crores' : 'Crore'}`)
  if (lakh) parts.push(`${threeDigits(lakh)} ${lakh > 1 ? 'Lakhs' : 'Lakh'}`)
  if (thousand) parts.push(`${threeDigits(thousand)} Thousand`)
  if (n) parts.push(threeDigits(n))
  return parts.join(' ')
}

/**
 * The parenthetical amount-in-words the letter wraps as "(Rupees … only)".
 * 3976612.18 → "Thirty Nine Lakhs Seventy Six Thousand Six Hundred Twelve and
 * Eighteen Paise"; a whole amount drops the paise clause.
 */
export function amountInWords(amount: number | null | undefined): string {
  if (amount == null || !Number.isFinite(amount)) return ''
  const rupees = Math.floor(amount)
  const paise = Math.round((amount - rupees) * 100)
  const rupeeWords = integerToIndianWords(rupees)
  if (paise > 0) return `${rupeeWords} and ${integerToIndianWords(paise)} Paise`
  return rupeeWords
}
