import { stripDecorativeWorkNameTags } from './nameOfWork'

/** 1 -> "1st", 2 -> "2nd", 3 -> "3rd", 4 -> "4th", 11 -> "11th", 21 -> "21st". */
export function ordinal(n: number): string {
  const rem100 = n % 100
  if (rem100 >= 11 && rem100 <= 13) return `${n}th`
  switch (n % 10) {
    case 1:
      return `${n}st`
    case 2:
      return `${n}nd`
    case 3:
      return `${n}rd`
    default:
      return `${n}th`
  }
}

/** "" for round 0 (no recall), else " (Nth Recall)" — space-prefixed so it appends directly after a name. */
export function recallTagFor(round: number): string {
  return round > 0 ? ` (${ordinal(round)} Recall)` : ''
}

/**
 * Work name with any decorative recall/reserved tag it already carries
 * stripped, then the dialog-chosen recall round re-appended — so a name
 * re-tagged in the Works List by hand (e.g. "...CMC(5TH Recall)") never
 * ends up double-tagged once the Tender Notice dialog's own round counter
 * also appends one.
 */
export function applyRecallTag(name: string, round: number): string {
  return stripDecorativeWorkNameTags(name) + recallTagFor(round)
}
