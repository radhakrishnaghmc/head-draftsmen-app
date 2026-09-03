import { cnoOf, corporationOf } from './zoneCircleDirectory'

/**
 * The Head Draughtsman's chosen office — Corporation → Zone → Circle — selected in
 * the sidebar (not from the login). This is what drives which circle/zone the
 * documents and Works List validation are prepared for. Circle (and its number)
 * are optional so a zone-level Head Draughtsman can pick just Corporation + Zone.
 */
export interface Office {
  corporation?: string
  zone?: string
  circle?: string
  /** Circle number (CNO), derived from the directory for the chosen circle. */
  circleNumber?: string
}

// Remembered across app restarts (localStorage) so the office survives a
// relaunch — and a logout: the office is a property of this machine's Head
// Draughtsman, not of the login session, so it's never re-asked once picked
// (it can still be changed from the sidebar). In dev, where the login screen is
// skipped, it also gives an identity to stamp the Works List with.
export const OFFICE_KEYS = {
  corporation: 'hda-corporation',
  zone: 'hda-zone',
  circle: 'hda-circle',
  circleNumber: 'hda-circle-number'
} as const

/** Read the remembered office from localStorage. */
export function loadOffice(): Office {
  const office: Office = {
    corporation: localStorage.getItem(OFFICE_KEYS.corporation) ?? undefined,
    zone: localStorage.getItem(OFFICE_KEYS.zone) ?? undefined,
    circle: localStorage.getItem(OFFICE_KEYS.circle) ?? undefined,
    circleNumber: localStorage.getItem(OFFICE_KEYS.circleNumber) ?? undefined
  }
  // An office saved before the Corporation field existed (old login-based flow)
  // has a Zone/Circle but no Corporation — infer it from the directory so it
  // shows and validates correctly without forcing a re-pick.
  if (!office.corporation && (office.zone || office.circle)) {
    office.corporation = corporationOf(office.zone, office.circle)
  }
  return normalizeOffice(office)
}

/** Persist (or clear) the office in localStorage. */
export function saveOffice(office: Office): void {
  const set = (key: string, val: string | undefined) => {
    if (val) localStorage.setItem(key, val)
    else localStorage.removeItem(key)
  }
  set(OFFICE_KEYS.corporation, office.corporation)
  set(OFFICE_KEYS.zone, office.zone)
  set(OFFICE_KEYS.circle, office.circle)
  set(OFFICE_KEYS.circleNumber, office.circleNumber)
}

/**
 * Normalise a partial office selection into a complete Office: fills the circle
 * number from the directory for the chosen corporation/zone/circle, and drops a
 * stale circle/number when the selection above it changed.
 */
export function normalizeOffice(next: Office): Office {
  const circleNumber = cnoOf(next.corporation, next.zone, next.circle)
  return { ...next, circleNumber }
}

/**
 * A stable key identifying an office ("Corporation|Zone|Circle"), used to
 * remember that office's own Works List link. Undefined until at least a
 * corporation + zone are chosen (a zonal office has an empty circle segment).
 */
export function officeKey(office: Office): string | undefined {
  if (!office.corporation || !office.zone) return undefined
  return [office.corporation, office.zone, office.circle ?? ''].join('|')
}

/**
 * A localStorage key scoped to this office, so a value entered for one office
 * (e.g. a contact phone number, remembered "per machine" rather than synced)
 * doesn't leak into another office switched to later. Falls back to the bare
 * key when no office is chosen yet.
 */
export function officeScopedKey(base: string, office?: Office): string {
  const key = office ? officeKey(office) : undefined
  return key ? `${base}:${key}` : base
}

/** A Zone chosen with no Circle of its own — a zone-level (SE) office, as opposed to a Circle (EE) office. Same test used throughout the app (WorkOrderAgreementTab's `seMode`, GiveIntimationTab, etc). */
export function isZoneOnlyOffice(office?: Office): boolean {
  return !!office?.zone?.trim() && !office?.circle?.trim()
}

// The Issue Tender Notice contact details (e-mail + the two mobile numbers)
// are properties of the issuing office, not of a single notice, so they're
// remembered per machine — scoped per office via officeScopedKey, so
// switching offices doesn't leak one office's numbers into another — and
// pre-filled next time rather than re-typed for every notice. Shared between
// TenderNoticeButton's dialog and Dashboard's always-visible fields (next to
// the 3-day/7-day tender toggle), so entering it in either place fills the other.
export const CONTACT_KEYS = {
  email: 'hda-tn-email',
  eePhone: 'hda-tn-ee-phone',
  hdPhone: 'hda-tn-hd-phone'
} as const

// Which bundled document-template variant an office has chosen (see
// core/workOrderTemplateVariants.ts, and Settings' Document Templates
// section) — scoped per office the same way as CONTACT_KEYS above, so one
// circle's chosen Work Order style doesn't leak into another's. Stores just
// the variant id (a short string), not a file — every variant already ships
// bundled with the app.
export const TEMPLATE_KEYS = {
  workOrder: 'hda-template-work-order',
  agreement: 'hda-template-agreement',
  intimation: 'hda-template-intimation',
  fileBacker: 'hda-template-file-backer',
  civilTender: 'hda-template-civil-tender'
} as const

/**
 * Whether an office has been chosen. A Corporation + Zone is enough — that's a
 * zonal office (a zone-level Head Draughtsman, spanning every circle in the zone);
 * picking a Circle as well narrows it to a single circle office.
 */
export function isOfficeReady(office: Office): boolean {
  return !!office.corporation && !!office.zone
}
