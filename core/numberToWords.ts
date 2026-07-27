// Human-readable renderings of amounts and dates for the Work Order /
// Agreement documents — the agreement spells the contract value out in words
// ("(Rupees Fourteen Lakh …)") and states its own date in words ("made on
// this 4th day of July 2026"). Indian numbering (Lakh/Crore), not the
// short scale.

const ONES = [
  '',
  'One',
  'Two',
  'Three',
  'Four',
  'Five',
  'Six',
  'Seven',
  'Eight',
  'Nine',
  'Ten',
  'Eleven',
  'Twelve',
  'Thirteen',
  'Fourteen',
  'Fifteen',
  'Sixteen',
  'Seventeen',
  'Eighteen',
  'Nineteen'
]
const TENS = ['', '', 'Twenty', 'Thirty', 'Forty', 'Fifty', 'Sixty', 'Seventy', 'Eighty', 'Ninety']

/** 0–99 in words ("", "One", … "Ninety Nine"). */
function twoDigits(n: number): string {
  if (n < 20) return ONES[n]
  const t = TENS[Math.floor(n / 10)]
  const o = ONES[n % 10]
  return o ? `${t} ${o}` : t
}

/** 0–999 in words ("Two Hundred Five"). */
function threeDigits(n: number): string {
  const h = Math.floor(n / 100)
  const rest = n % 100
  const parts: string[] = []
  if (h) parts.push(`${ONES[h]} Hundred`)
  if (rest) parts.push(twoDigits(rest))
  return parts.join(' ')
}

/**
 * A whole rupee count in Indian-numbering words: "Fourteen Lakh Forty Five
 * Thousand Nine Hundred Eighty Three". 0 -> "Zero". No "Rupees"/"Only"
 * wrapper — the caller adds those (the agreement template already prints
 * "Rupees" before this value).
 */
export function integerToIndianWords(value: number): string {
  let n = Math.floor(Math.abs(value))
  if (n === 0) return 'Zero'
  const crore = Math.floor(n / 10000000)
  n %= 10000000
  const lakh = Math.floor(n / 100000)
  n %= 100000
  const thousand = Math.floor(n / 1000)
  const rest = n % 1000
  const parts: string[] = []
  if (crore) parts.push(`${integerToIndianWords(crore)} Crore`)
  if (lakh) parts.push(`${threeDigits(lakh)} Lakh`)
  if (thousand) parts.push(`${threeDigits(thousand)} Thousand`)
  if (rest) parts.push(threeDigits(rest))
  return parts.join(' ')
}

/**
 * A rupees-and-paise amount as the office spells it in an agreement:
 * "Fourteen Lakh Forty Five Thousand Nine Hundred Eighty Three Rupees and
 * Seventeen Paise Only" (paise dropped when zero). NaN/negative -> "".
 */
export function amountToWords(rupees: number): string {
  if (!Number.isFinite(rupees) || rupees < 0) return ''
  const whole = Math.floor(rupees)
  const paise = Math.round((rupees - whole) * 100)
  const rupeeWords = `${integerToIndianWords(whole)} Rupees`
  if (paise > 0) return `${rupeeWords} and ${twoDigits(paise)} Paise Only`
  return `${rupeeWords} Only`
}

const MONTHS = [
  'January',
  'February',
  'March',
  'April',
  'May',
  'June',
  'July',
  'August',
  'September',
  'October',
  'November',
  'December'
]

/** 1 -> "1st", 2 -> "2nd", 3 -> "3rd", 4 -> "4th", 11 -> "11th", 21 -> "21st". */
export function ordinal(day: number): string {
  const mod100 = day % 100
  if (mod100 >= 11 && mod100 <= 13) return `${day}th`
  switch (day % 10) {
    case 1:
      return `${day}st`
    case 2:
      return `${day}nd`
    case 3:
      return `${day}rd`
    default:
      return `${day}th`
  }
}

/**
 * Parse the app's own date strings — "dd.mm.yyyy", "dd/mm/yyyy",
 * "dd-mm-yyyy", or an ISO "yyyy-mm-dd" — to a {day, month, year}. Returns
 * null for anything it can't read (blank, half-typed), so the caller leaves
 * the date-in-words placeholder blank rather than printing "NaN".
 */
export function parseDateParts(value: string): { day: number; month: number; year: number } | null {
  const s = value.trim()
  if (!s) return null
  const iso = /^(\d{4})-(\d{1,2})-(\d{1,2})$/.exec(s)
  if (iso) {
    const year = Number(iso[1])
    const month = Number(iso[2])
    const day = Number(iso[3])
    return isValidDmy(day, month, year) ? { day, month, year } : null
  }
  const dmy = /^(\d{1,2})[.\/-](\d{1,2})[.\/-](\d{2,4})$/.exec(s)
  if (dmy) {
    const day = Number(dmy[1])
    const month = Number(dmy[2])
    let year = Number(dmy[3])
    if (year < 100) year += 2000
    return isValidDmy(day, month, year) ? { day, month, year } : null
  }
  return null
}

function isValidDmy(day: number, month: number, year: number): boolean {
  return month >= 1 && month <= 12 && day >= 1 && day <= 31 && year >= 1900 && year <= 2999
}

/** "04.07.2026" -> "4th day of July 2026". Unparseable -> "". */
export function dateToWords(value: string): string {
  const parts = parseDateParts(value)
  if (!parts) return ''
  return `${ordinal(parts.day)} day of ${MONTHS[parts.month - 1]} ${parts.year}`
}

/** A Date -> the app's canonical "dd.mm.yyyy" date string. */
export function formatDmy(d: Date): string {
  const dd = String(d.getDate()).padStart(2, '0')
  const mm = String(d.getMonth() + 1).padStart(2, '0')
  return `${dd}.${mm}.${d.getFullYear()}`
}
