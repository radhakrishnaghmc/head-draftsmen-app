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
    version: '1.29.2',
    changes: [
      'Dashboard is now the screen you land on when you open the app.',
      'Dashboard "Upcoming" card now also shows NIT/tender bid-closing reminders, not just To Do and MB Scrutiny items.',
      'Tender reminders are now kept separately per office, like the Works List and To Do — a reminder added under one office no longer shows up under another.',
      'Fixed the Monitoring Format import not finding the right sheet for some zone-level offices.',
      'Fixed the Dashboard scrollbar overlapping the profile icon in the top-right corner.'
    ]
  },
  {
    version: '1.29.1',
    changes: [
      'New tool: Cement & Steel Rates, in the sidebar — shows the latest official circular rates from the Telangana Public Health department, with an alert when a new circular is published.',
      'Work Order, Intimation, File Backer and Tender Document templates each have a new logo-header style option (Telangana emblem + corporation title + corporation logo) alongside the existing plain-text header.',
      'Evaluation Sheet now accepts several "View Bidders" PDF uploads (or phone photos of the same page) at once and issues one Bid Capacity Evaluation Sheet per tender, instead of one at a time.',
      'MB Measurement tool: fixed misread depth values written as a handwritten average fraction (e.g. two readings over a line, divided by 2), and dropped stray OCR noise in the No./L/B/D columns instead of showing it as fabricated data.',
      'Fixed the app icon showing up invisible/blank on dark Windows taskbars and desktops.'
    ]
  },
  {
    version: '1.29.0',
    changes: [
      'Give Intimation, Agreement/Work Order, and Issue Documents now have a "Verify" button that checks every generated document for blank fields, leftover {{placeholders}}, or wrong EMD/ASD amounts before you send it out.',
      'Fixed: EMD and ASD figures in Intimation letters could land ₹1 off (e.g. 6204 instead of 6205) due to a rounding mismatch.',
      'Zone-level (Superintending Engineer) offices in MMC now get their own Letter of Acceptance wording, separate from the standard one.',
      'New tool (Settings → MB Measurement): upload photos of an MB (Measurement Book) sheet and get back an Excel-ready table, read automatically from the form.',
      'Uploading a multi-sheet estimate workbook now correctly finds the real item-wise estimate sheet instead of misreading a General Abstract or rate-schedule sheet.',
      'Fixed: a login failure from no internet connection now shows a clear message instead of a technical error.'
    ]
  },
  {
    version: '1.28.0',
    changes: [
      'Documents can now be issued for MMC (Malkajgiri Municipal Corporation) — its Malkajgiri, Uppal and LB Nagar zones/circles are now available in the Corporation → Zone → Circle picker.',
      'All bundled document templates (Tender Notice, Bid Document, Work Order, Agreement, File Backer, Issue Notices, Print Documents, and others) now print the actual issuing corporation\'s name instead of always printing "Cyberabad Municipal Corporation" / "CMC".',
      'Fixed: scanning a folder for tender documents could hang forever on a stalled cloud-sync file (OneDrive/Google Drive/Dropbox placeholder) or a dropped network drive — it now skips that file after a short wait instead of freezing the scan.',
      'Settings: template sections (Bid Document, Work Order, Agreement, etc.) are now collapsible, show the currently selected style at a glance, and the "preview" affordance now matches the rest of the app\'s "Click to preview" style instead of a small icon button.'
    ]
  },
  {
    version: '1.27.0',
    changes: [
      'Settings: added "Themes" — choose how Issue Documents tiles look on this device (Default, Colured, Windows theme, or Dark mode).',
      'Fixed: uploading a scanned or complex PDF/image could get stuck on "Reading…" forever with no error — it now gives up after a wait and shows a clear message instead.',
      'PDF conversion (Print/Preview) now shows a clear message if LibreOffice isn\'t installed, and converts multi-page documents noticeably faster.',
      'Issue Documents: batch document generation now shows a live progress percentage instead of appearing frozen.',
      'Work Order/Agreement: fixed Schedule A saving with the wrong (Executive Engineer) format for a Zone-level (SE) office when saved together with the bundle.',
      'Estimate BOQ files are now named using the work\'s actual name instead of the estimate sheet\'s tab name.',
      'Give Intimation (SE/Zone offices): fixed reference-list numbering alignment in the TS Note, Eligibility, Bid Evaluation, and Agency Approval notes.',
      'Faster document previews — the first Issue Documents preview no longer waits on the matching model to load, and repeated previews skip re-matching columns that haven\'t changed.'
    ]
  },
  {
    version: '1.26.0',
    changes: [
      'Login: if you\'ve successfully signed in on this computer before, you can now sign in again with no internet connection — only for that same login, not anyone else\'s.',
      'Fixed: picking a large folder for "Update from L1/LOA" or "Address from Intimation" could freeze the whole app while it scanned — it no longer blocks the app, and the button now fills in with a live percentage while it reads.'
    ]
  },
  {
    version: '1.25.0',
    changes: [
      'Works List: added an Errors button that checks your database for a duplicate Wincode/Tender ID, a work entered under two different Wincodes, an ECV greater than the Amount of estimate, or an EMD 1.5% that isn\'t greater than EMD 1% — it blinks when it finds a problem, tap it to see exactly which rows.',
      'Works List: "Update from L1/LOA" and "Address from Intimation" now accept a whole folder (or several files/folders at once), not just one file at a time — pick your office\'s Tender Evaluations folder and every L1/Intimation sheet inside it (including subfolders) gets scanned in one go.',
      'Works List: fixed the table not scrolling sideways to show every column — it now scrolls properly and the scrollbar is visible.',
      'Fixed: your office could get silently forgotten and asked again after logging out and back in.',
      'The "Log" link at the bottom of the sidebar now shows the full update history, not just the latest version.',
      'Note Submitted: corrected the wording used for a reserved-category work\'s EMD exemption note.'
    ]
  },
  {
    version: '1.24.0',
    changes: [
      'Login: if all 5 devices are already signed in, you can now force-log-out every other device right from the login screen instead of having to walk to another machine.',
      'Settings: added an "Active Devices" list showing every device currently signed in with your login, each with its own Log Out button — so you can free up a slot or sign out a device you no longer use.'
    ]
  },
  {
    version: '1.23.3',
    changes: [
      'Windows: reverted the bundled-LibreOffice change from the last two updates — it made previews slower and broke saving a document as PDF for some users. Document preview, print, and Save as PDF now work the same way they did before 1.23.1; if you use those features, please make sure LibreOffice is installed separately (libreoffice.org) as before.',
      'Work Order/Agreement document tiles go back to their fast preview — the exact-match preview from the last update wasn\'t worth the slow-down it caused.'
    ]
  },
  {
    version: '1.23.2',
    changes: [
      'Fixed: the Work Order/Agreement document tiles could take a long time to load their previews, especially on Windows — all the tiles on a page now render together instead of one at a time.'
    ]
  },
  {
    version: '1.23.1',
    changes: [
      'Windows: LibreOffice is now bundled with the app, same as Mac — no separate LibreOffice install is needed for accurate document preview, print, or PDF export.'
    ]
  },
  {
    version: '1.23.0',
    changes: [
      'Added a Kompally-style Agreement Bond template — pick it in Settings alongside the Kompally Work Order template.',
      'Fixed: the SE office Contract Deed did not match the real 5-page document — rebuilt to match exactly.',
      'File Backer (Agreement) now also prints the Tender ID and NIT No & Date above the Name of the Work.',
      'Issue Notice now opens as a tile with an expandable full preview, matching Intimation and Agreement.',
      'Agreement and Work Order tile previews are sharper and no longer look distorted.',
      'Removed the unused "CE Sanction Letter No & Date" field and the separate EMD Details box from the Agreement workspace (EMD is already picked up automatically from the Balance EMD/Bank Guarantee upload).',
      'Fixed: several full-page preview windows (document previews, photo tools, print) could open with their close button unreachable.',
      'Fixed: the search box on the Issue Tender Notice page was oversized and inconsistent with the rest of the app.',
      'Empty text boxes across the app now show a plain white background for readability.',
      'Preview and scroll areas now use a slimmer, rounded scrollbar consistent with the rest of the app.',
      'Document preview and PDF conversion are noticeably faster, especially for multi-page documents.'
    ]
  },
  {
    version: '1.22.0',
    changes: [
      'Agreement/Work Order: the EMD Details line can now be filled automatically by uploading either the CURE portal\'s Balance EMD payment receipt or a Bank Guarantee certificate (for offices that pay EMD/ASD via BG instead) — the BG\'s amounts are matched to EMD vs. ASD automatically from the work\'s ECV.',
      'Works List import: works whose Circle differs from your office\'s own (e.g. left over from a circle reorganisation, like Moosapet\'s wards merging into Kukatpally) are now accepted instead of rejecting the whole import. If more than a fifth of the works belong to another circle, you\'ll be asked to confirm before importing.',
      'Added the Serilingampally zone abbreviation (SLP) alongside Quthbullapur (QBZ) and Kukatpally (KPZ).',
      'Fixed: document previews (opening a single document, or Print) could show garbled text for some fonts — now renders accurately.',
      'Fixed: the app could feel very slow while previews were showing — live thumbnails are fast again; full-accuracy rendering is now used only when you open a single document or print.'
    ]
  },
  {
    version: '1.21.2',
    changes: [
      'Fixed: several SE documents (Agreement Bond, Bid Document, Agreement Put-up Note, Contract Deed, Eligibility Criteria, TS Note, and others) could show Word\'s "unreadable content, recover?" prompt — on Windows Word AND Office 365 — when downloaded. Several distinct causes are now all fixed: invalid formatting tags in 5 templates, documents whose layout elements were out of order, and some documents saved without Word-compatibility cleanup applied at all.',
      'Fixed: a stray "Number" word from the L-1 sheet could leak into the printed Name of Work on Agreement/Work Order documents.'
    ]
  },
  {
    version: '1.21.1',
    changes: [
      'Fixed: on Windows, uploading a photo or PDF anywhere in Tools (estimate from photos, photos to a Word/Excel document) failed with a "sharp module" error — every OCR-based tool now works correctly again.'
    ]
  },
  {
    version: '1.21.0',
    changes: [
      'Fixed: Note Submitted could show a different work\'s name, agency, ECV, tender %, NIT No, and contract amount than the rest of the Agreement/Work Order documents — every field now comes from the uploaded L-1 sheet / Online Intimation, same as every other document.',
      'Fixed: an unrelated work could occasionally get matched to the uploaded L-1 sheet in the Works List, Schedule A, and Note Submitted — matching now requires a much closer name match.',
      'New: "Clear" button on the Intimation and Agreement/Work Order pages to start a fresh work without switching tabs.',
      'Fixed: office (Corporation/Zone/Circle) was remembered per computer — logging in as a different person on the same machine skipped the office picker. Now remembered per login.',
      'Fixed: the Works List search box could hide the scrollbar, making rows below the visible area unreachable while searching.',
      'Fixed: an "ITEM N Dated:" tag from the L-1 sheet could leak into the printed Name of Work.',
      'Fixed: a Tender Notice No. with no separator before "Circle" (e.g. "NizampetCircle58") wasn\'t recognized — Circle came back blank.',
      'Fixed: Amount of Estimate could show blank on the Intimation and Agreement/Work Order pages even when the L-1 sheet\'s own ECV was available.',
      'Fixed: the Agreement/Work Order page could feel slow to respond to clicks after uploading the Online Intimation and L-1 sheet.',
      'SE Agreement Bond: further signature-block position adjustment.'
    ]
  },
  {
    version: '1.20.1',
    changes: [
      'Fixed: Word showed "unreadable content, recover?" when opening a downloaded Agreement Bond, Agreement Put-up Note, Contract Deed, Eligibility Criteria, or TS Note — a leftover invalid reference inside the file is now removed.',
      'Fixed: SE Agreement Bond layout — more space above "AGREEMENT", tighter spacing around Name of Work, and the body paragraphs are flush left instead of oddly indented.'
    ]
  },
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
