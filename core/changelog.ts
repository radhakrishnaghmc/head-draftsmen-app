/**
 * The user-facing "What's New" changelog. On the first launch after an update,
 * the app shows the changes for every version newer than the one last seen on
 * this machine (see src/components/WhatsNew.tsx and App.tsx). Keep entries short,
 * plain-language, and written for a Head Draughtsman — not a developer.
 *
 * Add a new entry at the TOP (newest first) whenever a release ships, matching
 * the version in package.json.
 */
export interface ChangelogEntry {
  version: string
  /** One short, plain-language line per user-visible change. */
  changes: string[]
}

// Newest first. The top entry's version should match package.json.
export const CHANGELOG: ChangelogEntry[] = [
  {
    version: '1.20.0',
    changes: [
      'Agreement and Work Order: "Download all documents" is now available for SE (Zone) offices too, in the page header, with the same Word/PDF toggle as EE offices — and now shows live progress ("Preparing 3 of 8…", "Saving 5 of 8…") instead of a plain "Preparing…" that could look frozen.',
      'Fixed: Bid Evaluation and Agency Approval notes failed to open with an "Invalid XML name" error — now generate correctly.',
      'Fixed: SE Work Order date was never filling in ("DT." stayed blank).',
      'Fixed: Memo to EE showed a blank circle ("Executive Engineer, , Zone") when no Works List row matched — now fills in from the NIT/work name.',
      'Fixed: the Item No. tag was printing inside the Name of Work text on every document — now shown only in its own Item No. field.',
      'Fixed: SE Agreement Bond page layout (blank top third for the letterhead) and the Agmt. No./Date line alignment.',
      'Fixed: SE Agreement Put-up Note — Sub/Ref block alignment and a missing Item No. in the reference line.',
      'Fixed: Memo Concluding Agreement and Contract Deed — the agency address block no longer collapses to the left margin partway through.',
      'Fixed: TS Note — a stray blank space in the "Lr no." reference line.',
      'Works List: removed the duplicate "Change your Office" link (use the office picker in the sidebar).'
    ]
  },
  {
    version: '1.19.0',
    changes: [
      'New: Bid Document for SE (Zone) offices — its own format with Item No., Technical Sanction and Administrative Sanction fields, replacing the Circle-office format automatically for a Zone office.',
      'New: an "Issue Bid Document" tile below the Calendar for SE (Zone) offices — pick a work from the Works List and generate its Bid Document directly, without issuing a tender notice first.',
      'Give Intimation (SE/Zone offices): document previews (Intimation, TS Note, Eligibility Criteria, Bid Evaluation, Agency Approval) now show their fill-in fields beside the preview instead of above it, so the preview is easier to see.',
      'Every page now keeps its title bar in place while you scroll its contents, instead of the title scrolling out of view.',
      'Fixed: the Forwarding Slip’s Technical Sanction No & Date was printing in tiny, misplaced text.',
      'Sidebar: added a background image and a button to reopen this "What’s New" list at any time.'
    ]
  },
  {
    version: '1.18.0',
    changes: [
      'Give Intimation (SE/Zone offices): 4 new companion documents — TS Note, Eligibility Criteria, Bid Evaluation Note, Agency Approval Note — with Cement & Steel rate circular upload.',
      'Agreement and Work Order (SE/Zone offices): 3 new documents — Agreement Bond, Agreement Put-up Note, Contract Deed — plus an SE-formatted Schedule A, and Balance EMD receipt upload.',
      'Schedule A now shows automatically alongside the other documents, filled in blank — no need to upload the estimate first just to see it.',
      'Give Intimation page redesigned to match the Agreement page: documents now show as a tile grid with a preview popup, plus a "Download all documents" button.',
      'Photos to PDF: added a crop, rotate, and filter (black & white / contrast / auto-enhance) step before converting photos.',
      'Fixed: SE Agreement Bond date wasn’t filling in; several documents’ signature blocks (SE Agreement Bond, Agreement Note, Contract Deed, Forwarding Slip, TS Note) now line up the title and office/zone line properly.',
      'Fixed: Balance EMD upload button no longer appears for EE (Circle) offices.',
      'Small fix: clicking outside a popup to close it no longer misfires while selecting text inside it.'
    ]
  },
  {
    version: '1.17.1',
    changes: [
      'Intimation & Note Submitted: for a reserved-category work quoted more than 25% below, the ASD amount now shows (only the EMD is exempted for reserved works, not the ASD).',
      'Note Submitted: the non-responsive sheet now counts every disqualified agency, not just the first one.',
      'Note Submitted: each paragraph now opens with a first-line indent for a cleaner, letter-style layout.',
      'Work Order: the Executive Engineer signature is now a neat centered block on the right — Executive Engineer, circle, and corporation aligned under one another.'
    ]
  },
  {
    version: '1.17.0',
    changes: [
      'To Do List and MB Scrutiny list are now kept separately for each office — what you add under one circle no longer shows under another.',
      'Completion Report: a blank Estimate Amount now stays blank (no more "Rs 0/-"), and the stray "dt." is removed when there is no Technical Sanction date.',
      'After an update, the app now shows a short summary of what changed (this window).'
    ]
  },
  {
    version: '1.16.0',
    changes: [
      'New tool: Photos / PDF → Word or Excel — read photos or a scanned PDF into editable text and save as a Word (.docx) or Excel (.xlsx) file, keeping tables and layout.',
      'GPS Photos: better at reading faint or stamped Latitude–Longitude overlays.',
      'Tender Notice now fills the correct circle name, circle number and contact numbers for the office issuing it.',
      'Smaller fixes to the Work Order / Agreement (ECV) and the Bid Document.'
    ]
  }
]

/**
 * Compare two dotted version strings ("1.16.0"). Returns -1 if a < b, 1 if
 * a > b, 0 if equal. Missing segments count as 0, and any non-numeric segment
 * as 0, so it never throws on an odd version string — it just orders it low.
 */
export function compareVersions(a: string, b: string): number {
  const pa = a.split('.')
  const pb = b.split('.')
  const len = Math.max(pa.length, pb.length)
  for (let i = 0; i < len; i++) {
    const na = Number(pa[i] ?? 0) || 0
    const nb = Number(pb[i] ?? 0) || 0
    if (na < nb) return -1
    if (na > nb) return 1
  }
  return 0
}

/**
 * The changelog entries to show on this launch: every version strictly newer
 * than `seen` and no newer than `current`, newest first. A fresh install
 * (`seen` null/blank) shows nothing — there's no prior version to announce an
 * update from — and so does a launch where nothing has changed since last seen.
 */
export function changesSince(seen: string | null | undefined, current: string): ChangelogEntry[] {
  if (!seen) return []
  return CHANGELOG.filter(
    (e) => compareVersions(e.version, seen) > 0 && compareVersions(e.version, current) <= 0
  )
}
