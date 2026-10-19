import type { QcOfficeParties, QcPartyDetails } from './types'

const blank = (p: QcPartyDetails | undefined): boolean => !p || !(p.name.trim() || p.address.trim() || p.phone.trim())

/**
 * Combine the cloud's and this device's saved department QC / 3rd / 4th-party agencies, per office
 * and per party. The cloud copy wins where it has anything entered; a party the
 * cloud has blank or missing keeps the local entry. The cloud push is
 * fire-and-forget, so right after entering agencies and logging out / quitting it
 * may still be empty — without this the login pull would overwrite the cache and
 * the agencies would vanish.
 */
export function mergeQcParties(
  remote: Record<string, QcOfficeParties> | undefined,
  local: Record<string, QcOfficeParties> | undefined
): Record<string, QcOfficeParties> | undefined {
  if (!remote && !local) return undefined
  const out: Record<string, QcOfficeParties> = { ...(remote ?? {}) }
  for (const [office, l] of Object.entries(local ?? {})) {
    const r = out[office]
    out[office] = {
      ...(blank(r?.department) && blank(l.department) ? {} : { department: blank(r?.department) ? l.department : r!.department }),
      third: blank(r?.third) ? l.third : r!.third,
      fourth: blank(r?.fourth) ? l.fourth : r!.fourth
    }
  }
  return out
}
