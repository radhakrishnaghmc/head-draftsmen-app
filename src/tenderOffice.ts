import { circleFromNit } from '../core/workOrderAgreement'
import { entriesOf, resolveFromDirectory } from './zoneCircleDirectory'

/**
 * The office a saved tender belongs to, decided by its NIT number alone: "EE" in the NIT is an Executive
 * Engineer's (circle-level) office, "SE" a Superintending Engineer's (zone-level) one.
 *
 * - SE: the zone-level office of `zone` (the NIT only abbreviates the zone, e.g. "QBZ"), key "Corp|Zone|".
 * - EE: the circle the work name states ("… Nizampet Circle-58, Quthbullapur Zone"). Only when the work name
 *   names none does the NIT's own circle serve, by name ("…/EE/Nizampet Circle-58/…") or by number alone
 *   ("26/EE-59/QBZ/…" is circle 59).
 *
 * Undefined when the NIT is missing, names neither level, or an EE tender's circle can't be found anywhere —
 * the caller keeps the tender in its current office then.
 */
export function officeKeyFromNit(nit: string | undefined, corporation: string, zone: string, workName?: string): string | undefined {
  const text = nit ?? ''
  // "EE"/"SE" as a whole NIT segment ("/SE/", "/EE-59/", " SE "), never inside another word.
  const level = /(?:^|[\/\s])(EE|SE)(?=[\/\s-]|$)/i.exec(text)?.[1].toUpperCase()
  if (level === 'SE') return zone ? [corporation, zone, ''].join('|') : undefined
  if (level !== 'EE') return undefined

  const entries = entriesOf(corporation)
  const fromWork = resolveFromDirectory(workName, entries)
  if (fromWork.zone && fromWork.circle) return [corporation, fromWork.zone, fromWork.circle].join('|')
  const named = circleFromNit(text).circle
  const byName = named ? entries.find((e) => e.circle.toLowerCase() === named.toLowerCase()) : undefined
  const num = /\bEE[\s-]*(\d+)/i.exec(text)?.[1]
  const entry = byName ?? (num ? entries.find((e) => e.cno === num) : undefined)
  if (entry) return [corporation, entry.zone, entry.circle].join('|')
  return named && zone ? [corporation, zone, named].join('|') : undefined
}
