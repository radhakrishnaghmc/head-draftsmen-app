import { useEffect, useMemo, useRef, useState } from 'react'

/** One element of a feature's illustrated mock-up. `step` ties it to a numbered step in the text. */
type MockEl =
  | { kind: 'button'; label: string; step?: number }
  | { kind: 'upload'; label: string; step?: number }
  | { kind: 'tiles'; items: string[]; step?: number }
  | { kind: 'table'; cols: string[]; rows?: number; step?: number }
  | { kind: 'fields'; items: string[]; step?: number }
  | { kind: 'preview'; label: string; step?: number }
  | { kind: 'tabs'; items: string[]; step?: number }
  | { kind: 'list'; items: string[]; step?: number }

interface Feature {
  id: string
  /** Matches the sidebar label, which is highlighted in the mock-up. */
  nav: string
  title: string
  summary: string
  steps: string[]
  tips?: string[]
  mock: MockEl[]
}

interface Group {
  title: string
  features: Feature[]
}


const GROUPS: Group[] = [
  {
    title: 'Getting started',
    features: [
      {
        id: 'office',
        nav: 'Dashboard',
        title: 'Choose your office',
        summary:
          'Everything the app prepares — letterheads, document wording, the Works List, To Do list and MB Scrutiny list — is specific to your office. You pick it once; it is remembered on your account.',
        steps: [
          'Click “Change office” under the app name at the top of the sidebar (it opens by itself on first use).',
          'Pick the Corporation, then the Zone, then the Circle (a zone-level office needs only the Zone).',
          'Confirm. The office name now shows under the app name and every page uses it.'
        ],
        tips: ['The office is not tied to your login — change it any time from the same button.'],
        mock: [
          { kind: 'button', label: 'Change office', step: 1 },
          { kind: 'fields', items: ['Corporation', 'Zone', 'Circle'], step: 2 },
          { kind: 'button', label: 'Confirm', step: 3 }
        ]
      },
      {
        id: 'navigation',
        nav: 'Works List',
        title: 'Moving around the app',
        summary: 'The sidebar is split into Overview, Workflow, Utilities and System. Your work is never lost when you switch pages.',
        steps: [
          'Click any sidebar item to open that page. The page you leave keeps running in the background with everything you typed or uploaded.',
          'Double-click a sidebar item to close every open workspace and start that page fresh.',
          'Use the profile menu (top right) for your account, logout and “What’s New”.'
        ],
        mock: [
          { kind: 'list', items: ['Overview', 'Workflow', 'Utilities', 'System'], step: 1 },
          { kind: 'button', label: 'Double-click = fresh start', step: 2 }
        ]
      }
    ]
  },
  {
    title: 'Overview',
    features: [
      {
        id: 'dashboard',
        nav: 'Dashboard',
        title: 'Dashboard',
        summary: 'A one-screen picture of your office: works in progress, tender reminders, weather and the monitoring format.',
        steps: [
          'Open Dashboard — it is the first page after login.',
          'Read the summary cards for works, pending items and reminders.',
          'In the Monitoring Format card, import your office’s monitoring workbook; you can preview it as PDF or print it.'
        ],
        mock: [
          { kind: 'tiles', items: ['Works', 'Reminders', 'Weather'], step: 2 },
          { kind: 'upload', label: 'Import Monitoring Format (.xlsx)', step: 3 }
        ]
      },
      {
        id: 'works',
        nav: 'Works List',
        title: 'Works List',
        summary: 'The master register of your office’s works. Other pages (Intimation, Work Order, Agreement, Estimate) read from it.',
        steps: [
          'Open Works List. Works for the selected office appear in the table.',
          'Use the search bar to find a work by name or number.',
          'Add works by uploading your list (Excel/PDF) or typing rows directly.',
          'Update the L-1 (lowest bidder) details when a tender is evaluated — the later documents pick it up automatically.',
          'Works that belong to another circle are flagged so you do not mix them in.'
        ],
        mock: [
          { kind: 'upload', label: 'Upload works list', step: 3 },
          { kind: 'fields', items: ['Search works…'], step: 2 },
          { kind: 'table', cols: ['Work', 'Estimate', 'L-1', 'Status'], rows: 4, step: 1 }
        ]
      },
      {
        id: 'calendar',
        nav: 'Calendar',
        title: 'Calendar',
        summary: 'The Telangana Government holiday calendar for the year, to plan dates for notices and completion.',
        steps: ['Open Calendar.', 'Browse months to see holidays highlighted.', 'Use it when choosing a due date for a notice or a To Do task.'],
        mock: [{ kind: 'tabs', items: ['Jan', 'Feb', 'Mar', 'Apr'], step: 2 }, { kind: 'table', cols: ['Mon', 'Tue', 'Wed', 'Thu', 'Fri'], rows: 4, step: 3 }]
      }
    ]
  },
  {
    title: 'Workflow',
    features: [
      {
        id: 'intimation',
        nav: 'Intimation',
        title: 'Intimation',
        summary:
          'Turns the tender’s L-1 documents into the intimation letter (Letter of Acceptance), the Note Submitted, evaluation notes and the Bid Capacity Evaluation Sheet.',
        steps: [
          'Open Intimation and upload the L-1 PDF (or photos of it) plus, if you have one, your existing Intimation format.',
          'The app reads the document and fills the fields. Check each field on the left; edit anything that is wrong.',
          'The live page on the right updates as you type — this is exactly what will print.',
          'Download the Word document, or use the “Evaluation Sheet” sub-tab: upload the portal’s “View Bidders” PDF to get the Bid Capacity Evaluation Excel (one column per bidder).'
        ],
        tips: ['Zone-level offices get the SE Letter of Acceptance and the extra SE evaluation notes.'],
        mock: [
          { kind: 'tabs', items: ['Intimation', 'Evaluation Sheet'], step: 4 },
          { kind: 'upload', label: 'Upload L-1 PDF / photos', step: 1 },
          { kind: 'fields', items: ['Work name', 'Agency', 'Amount', 'Date'], step: 2 },
          { kind: 'preview', label: 'Live letter preview', step: 3 }
        ]
      },
      {
        id: 'agreement',
        nav: 'Agreement',
        title: 'Agreement and Work Order',
        summary: 'Prepares the Work Order, Agreement Bond, File Backer, Forwarding Slip, QCC Intimation and Schedule A / BOQ for a work.',
        steps: [
          'Open Agreement and Work order and upload the L-1 and the Intimation (or pick the work from the Works List).',
          'Review the auto-filled fields in the editor on the left.',
          'Choose the document tile you need — Work Order, Agreement Bond, File Backer, etc.',
          'Check the live page on the right, then download or print.',
          'Use “Clear” in the page header to wipe the current upload before starting the next work.'
        ],
        mock: [
          { kind: 'upload', label: 'Upload L-1 + Intimation', step: 1 },
          { kind: 'tiles', items: ['Work Order', 'Agreement', 'File Backer', 'Forwarding Slip'], step: 3 },
          { kind: 'fields', items: ['Agency', 'Agreement no.', 'Date'], step: 2 },
          { kind: 'preview', label: 'Live page', step: 4 }
        ]
      },
      {
        id: 'notices',
        nav: 'Issue Notices',
        title: 'Issue Notices',
        summary: 'Create and track notices to agencies. The sidebar item blinks red when any active notice is past its due date.',
        steps: [
          'Open Issue Notices and click “New notice”.',
          'Choose the kind of notice and fill the work and agency details.',
          'Set the due date and issue the notice (download/print).',
          'Track it under Active, Overdue or Resolved; mark it Resolved when handled.'
        ],
        mock: [
          { kind: 'button', label: '+ New notice', step: 1 },
          { kind: 'tabs', items: ['Active', 'Overdue', 'Resolved'], step: 4 },
          { kind: 'list', items: ['Notice 1 — due 12 Nov', 'Notice 2 — overdue'], step: 3 }
        ]
      },
      {
        id: 'documents',
        nav: 'Issue Documents',
        title: 'Issue Documents',
        summary: 'A library of ready-made office documents (tender document, bid document, completion report, EOT, deviation and many more) as live-preview tiles.',
        steps: [
          'Open Issue Documents and click a tile (each tile is a live preview of the document).',
          'Fill or confirm the fields; the page updates as you type.',
          'Download as Word/PDF or print. Created documents are counted on the sidebar badge.',
          'Pick which style is the default for each document in Settings → Document Templates.'
        ],
        mock: [
          { kind: 'tiles', items: ['Tender Doc', 'Bid Doc', 'Completion', 'EOT'], step: 1 },
          { kind: 'fields', items: ['Work', 'Agency', 'Date'], step: 2 },
          { kind: 'button', label: 'Download', step: 3 }
        ]
      },
      {
        id: 'estimate',
        nav: 'BOQ / Schedule A',
        title: 'BOQ, Schedule A, Deviation & Material Quantity',
        summary: 'Upload one estimate and get the BOQ, Schedule A, Deviation Statement and Material Quantity workbooks.',
        steps: [
          'Open the page and upload the estimate (Excel, PDF or photos).',
          'Optionally upload the Data Sheet so rates are filled from it.',
          'Review the extracted items and rates.',
          'Download the BOQ, Schedule A, Deviation Statement and Material Quantity.'
        ],
        mock: [
          { kind: 'upload', label: 'Upload estimate', step: 1 },
          { kind: 'upload', label: 'Upload Data Sheet (optional)', step: 2 },
          { kind: 'table', cols: ['Item', 'Qty', 'Rate', 'Amount'], rows: 3, step: 3 },
          { kind: 'tiles', items: ['BOQ', 'Schedule A', 'Deviation', 'Material'], step: 4 }
        ]
      },
      {
        id: 'todo',
        nav: 'To Do List',
        title: 'To Do List',
        summary: 'Tasks with a target date, kept separately for each office.',
        steps: [
          'Open To Do List and type a task with its target completion date.',
          'Tick a task when it is done — it stays visible for the rest of the day, then rolls off.',
          'Tasks sync across your devices.'
        ],
        mock: [{ kind: 'fields', items: ['New task…', 'Target date'], step: 1 }, { kind: 'list', items: ['☐ Send LOA', '☑ Issue notice'], step: 2 }]
      },
      {
        id: 'mb',
        nav: 'MB Scrutiny',
        title: 'MB Scrutiny list',
        summary: 'A register of Measurement Books received for scrutiny, tracked through to completion.',
        steps: [
          'Open MB Scrutiny list and add each Measurement Book when received.',
          'Update its stage as scrutiny progresses.',
          'Mark it complete when finished. The list is specific to the selected office.',
          'To read measurements from a scanned MB page, use the MB Measurement scan upload — it auto-orients pages and finds faint grids.'
        ],
        mock: [{ kind: 'button', label: '+ Add MB', step: 1 }, { kind: 'table', cols: ['MB no.', 'Work', 'Received', 'Stage'], rows: 3, step: 2 }]
      },
      {
        id: 'search',
        nav: 'Search Tender',
        title: 'Search Tender',
        summary: 'Search live tenders on the Telangana e-procurement portal.',
        steps: ['Open Search Tender.', 'Enter a tender ID, NIT number or part of the work name.', 'Open a result for its details; set a reminder if needed.'],
        mock: [{ kind: 'fields', items: ['Tender ID / NIT / work name'], step: 2 }, { kind: 'table', cols: ['Tender', 'Work', 'Closing'], rows: 3, step: 3 }]
      },
      {
        id: 'sanction',
        nav: 'Issue Documents',
        title: 'Give Technical Sanction (Experimental)',
        summary: 'Drafts a technical sanction using rates retrieved from your Data Sheet and the Schedule of Rates PDF together.',
        steps: [
          'Open Give Technical Sanction.',
          'Upload the Data Sheet / Schedule of Rates if asked, and the estimate.',
          'Review every suggested rate before accepting — this feature is experimental.'
        ],
        tips: ['Always verify rates against the SoR before issuing.'],
        mock: [{ kind: 'upload', label: 'Upload estimate', step: 2 }, { kind: 'table', cols: ['Item', 'Suggested rate', 'Source'], rows: 3, step: 3 }]
      }
    ]
  },
  {
    title: 'Agreement and Work Order — every document',
    features: [
      {
        id: "wo-overview",
        nav: 'Agreement',
        title: "How all documents are generated here",
        summary:
          "The page produces up to ten documents from the same uploads. Fill the details once; every tile below then carries the same values. Which tiles show depends on your office (a zone-level office gets its own set) and on which templates are available.",
        steps: [
          "Open Agreement and Work order. Upload the L1 selection form (and the Online Intimation, Supplier Details, Technical sanctioned Estimate if you have them) — or click “Fill details manually”.",
          "Check the green “read from…” lines and the Works List match; fix any red mismatch message (wrong circle or wrong work) before going on.",
          "Tiles appear for each document, each a live thumbnail filled with your data. Click a tile to open its full preview.",
          "In the preview, edit any field on the left; the page on the right updates. Choose Word, PDF or Print.",
          "Click “Download all documents” to save every ready document in one go (use Select to leave some out)."
        ],
        tips: ["The Tools tab has single-document versions (Work Order, Agreement Bond, Intimation, Schedule A) that work for any circle’s files."],
        mock: [
          { kind: 'upload', label: 'Upload L1 selection form + Intimation', step: 1 },
          { kind: 'fields', items: ['Work name', 'Agency', 'Amount'], step: 2 },
          { kind: 'tiles', items: ['File Backer', 'Note', 'Agreement', 'Work Order'], step: 3 },
          { kind: 'preview', label: 'Live preview', step: 4 },
          { kind: 'button', label: 'Download all documents', step: 5 }
        ]
      },
      {
        id: "wo-filebacker",
        nav: 'Agreement',
        title: "File Backer",
        summary:
          "The cover page of the work file, with the work name, agency, amounts and agreement details.",
        steps: [
          "Open Agreement and Work order. Upload the L1 selection form (and the Online Intimation, Supplier Details, Technical sanctioned Estimate if you have them) — or click “Fill details manually”.",
          "Click the “File Backer” tile.",
          "In the preview, check the fields on the left (work name, agency, estimate, contract amount, agreement no.).",
          "Click Word, PDF or Print."
        ],
        mock: [
          { kind: 'upload', label: 'Upload L1 selection form', step: 1 },
          { kind: 'tiles', items: ['File Backer'], step: 2 },
          { kind: 'fields', items: ['Work name', 'Agency', 'Amount'], step: 3 },
          { kind: 'button', label: 'Word / PDF / Print', step: 4 }
        ]
      },
      {
        id: "wo-note",
        nav: 'Agreement',
        title: "Note Submitted",
        summary:
          "The note recording the tender evaluation — bidders, L-1 percentage and the recommendation.",
        steps: [
          "Open Agreement and Work order. Upload the L1 selection form (and the Online Intimation, Supplier Details, Technical sanctioned Estimate if you have them) — or click “Fill details manually”.",
          "Click the “Note Submitted” tile.",
          "Also upload the Balance EMD Payment (PDF or photo) so the receipt number and date fill, and the Non-responsive form if some bidders were non-responsive.",
          "Review the note in the editor; fix anything that doesn’t match the other documents (the page warns if the work name or estimate differs).",
          "Click Word, PDF or Print."
        ],
        mock: [
          { kind: 'upload', label: 'Upload L1 selection form', step: 1 },
          { kind: 'tiles', items: ['Note Submitted'], step: 2 },
          { kind: 'upload', label: 'Upload Balance EMD receipt', step: 3 },
          { kind: 'preview', label: 'Note preview', step: 4 },
          { kind: 'button', label: 'Word / PDF / Print', step: 5 }
        ]
      },
      {
        id: "wo-dbnote",
        nav: 'Agreement',
        title: "DB Note",
        summary:
          "The bill note prepared for the work, with 28 numbered points (agreement, recording and check dates, M.B. details, penalties, recommended amount…).",
        steps: [
          "Open Agreement and Work order. Upload the L1 selection form (and the Online Intimation, Supplier Details, Technical sanctioned Estimate if you have them) — or click “Fill details manually”.",
          "Click the “DB Note” tile.",
          "Fill the numbered points on the left — the ones the app knows (work, agency, estimate, contract amount, agreement no.) are already filled; type the rest (recording dates, M.B. details, checks, recommended amount).",
          "Check the live page, then Word, PDF or Print."
        ],
        tips: ["This is the only document whose later points (bill details) you must type yourself."],
        mock: [
          { kind: 'tiles', items: ['DB Note'], step: 2 },
          { kind: 'fields', items: ['Work name', 'Agency', 'Amount', 'Date'], step: 3 },
          { kind: 'button', label: 'Word / PDF / Print', step: 4 }
        ]
      },
      {
        id: "wo-aeletters",
        nav: 'Agreement',
        title: "AE/Dy.EE Forwarding Letters",
        summary:
          "The forwarding letters sent to the AE and Dy.EE for the work, printed on the office letterhead.",
        steps: [
          "Open Agreement and Work order. Upload the L1 selection form (and the Online Intimation, Supplier Details, Technical sanctioned Estimate if you have them) — or click “Fill details manually”.",
          "Click the “AE/Dy.EE Forwarding Letters” tile.",
          "Upload the Online Intimation (for the address) and Supplier Details (for the phone number) so the agency contact details fill.",
          "Review and edit the letter fields.",
          "Word, PDF or Print."
        ],
        mock: [
          { kind: 'upload', label: 'Upload L1 selection form', step: 1 },
          { kind: 'tiles', items: ['AE/Dy.EE Letters'], step: 2 },
          { kind: 'upload', label: 'Upload Intimation + Supplier Details', step: 3 },
          { kind: 'fields', items: ['Agency', 'Phone', 'Address'], step: 4 },
          { kind: 'button', label: 'Word / PDF / Print', step: 5 }
        ]
      },
      {
        id: "wo-slip",
        nav: 'Agreement',
        title: "Forwarding Slip",
        summary:
          "The slip that forwards the agreement file for the work.",
        steps: [
          "Open Agreement and Work order. Upload the L1 selection form (and the Online Intimation, Supplier Details, Technical sanctioned Estimate if you have them) — or click “Fill details manually”.",
          "Click the “Forwarding Slip” tile.",
          "Upload the Online Intimation / Supplier Details so the agency address and phone fill.",
          "Review the fields, then Word, PDF or Print."
        ],
        mock: [
          { kind: 'upload', label: 'Upload L1 selection form', step: 1 },
          { kind: 'tiles', items: ['Forwarding Slip'], step: 2 },
          { kind: 'upload', label: 'Upload Intimation + Supplier Details', step: 3 },
          { kind: 'fields', items: ['Agency', 'Address', 'Amount'], step: 4 },
          { kind: 'button', label: 'Word / PDF / Print', step: 4 }
        ]
      },
      {
        id: "wo-agreement",
        nav: 'Agreement',
        title: "Agreement Bond",
        summary:
          "The agreement bond with the A.B. number, date and the “…day of…” wording.",
        steps: [
          "Open Agreement and Work order. Upload the L1 selection form (and the Online Intimation, Supplier Details, Technical sanctioned Estimate if you have them) — or click “Fill details manually”.",
          "Click the “Agreement Bond” tile.",
          "A box asks for the agreement date. Pick it — it fills the A.B. No line (dd.mm.yyyy) and the “…day of…” wording in words, and the Work Order date. If the L-1 doesn’t name a Circle/Zone, you enter those here too. Choose “Continue without date” to leave it blank for hand-dating.",
          "Enter the Agreement number, Technical Sanction no. & date, period of completion, number of pages and items in Schedule ‘A’ if shown.",
          "Check the page, then Word, PDF or Print."
        ],
        mock: [
          { kind: 'upload', label: 'Upload L1 selection form', step: 1 },
          { kind: 'tiles', items: ['Agreement Bond'], step: 2 },
          { kind: 'fields', items: ['Date'], step: 3 },
          { kind: 'fields', items: ['Agreement no.', 'Agency'], step: 4 },
          { kind: 'button', label: 'Word / PDF / Print', step: 5 }
        ]
      },
      {
        id: "wo-schedulea",
        nav: 'Agreement',
        title: "Schedule A",
        summary:
          "The priced schedule of items, built from the technical sanctioned estimate.",
        steps: [
          "Open Agreement and Work order. Upload the L1 selection form (and the Online Intimation, Supplier Details, Technical sanctioned Estimate if you have them) — or click “Fill details manually”.",
          "Click the “Schedule A” tile.",
          "Upload the Technical sanctioned Estimate (if not already uploaded).",
          "The Schedule A tile opens a preview of the items with rates and amounts.",
          "Download it as Excel or PDF, or Print."
        ],
        mock: [
          { kind: 'upload', label: 'Upload L1 selection form', step: 1 },
          { kind: 'tiles', items: ['Schedule A'], step: 2 },
          { kind: 'upload', label: 'Upload Technical sanctioned Estimate', step: 3 },
          { kind: 'table', cols: ['Item', 'Qty', 'Rate', 'Amount'], rows: 3, step: 4 },
          { kind: 'button', label: 'Excel / PDF / Print', step: 5 }
        ]
      },
      {
        id: "wo-workorder",
        nav: 'Agreement',
        title: "Work Order",
        summary:
          "The work order issued to the agency.",
        steps: [
          "Open Agreement and Work order. Upload the L1 selection form (and the Online Intimation, Supplier Details, Technical sanctioned Estimate if you have them) — or click “Fill details manually”.",
          "Click the “Work Order” tile.",
          "If the agreement date wasn’t asked yet, pick it in the box that appears (the same date as the Agreement Bond).",
          "Enter the Work Order number if it isn’t filled; check the AS file no., TS no., contract amount and tender percentage.",
          "Word, PDF or Print."
        ],
        mock: [
          { kind: 'upload', label: 'Upload L1 selection form', step: 1 },
          { kind: 'tiles', items: ['Work Order'], step: 2 },
          { kind: 'fields', items: ['Date'], step: 3 },
          { kind: 'fields', items: ['Agreement no.', 'Amount'], step: 4 },
          { kind: 'button', label: 'Word / PDF / Print', step: 5 }
        ]
      },
      {
        id: "wo-tender",
        nav: 'Agreement',
        title: "Tender Document",
        summary:
          "The civil tender document for the work, on the office letterhead.",
        steps: [
          "Open Agreement and Work order. Upload the L1 selection form (and the Online Intimation, Supplier Details, Technical sanctioned Estimate if you have them) — or click “Fill details manually”.",
          "Click the “Tender Document” tile.",
          "Upload the Online Intimation / Supplier Details if you want the agency address and phone filled.",
          "Review the NIT no., NIT date, estimate, ECV and tender percentage.",
          "Word, PDF or Print."
        ],
        mock: [
          { kind: 'upload', label: 'Upload L1 selection form', step: 1 },
          { kind: 'tiles', items: ['Tender Document'], step: 2 },
          { kind: 'upload', label: 'Upload Intimation + Supplier Details', step: 3 },
          { kind: 'fields', items: ['Work name', 'Amount'], step: 4 },
          { kind: 'button', label: 'Word / PDF / Print', step: 5 }
        ]
      },
      {
        id: "wo-qcc",
        nav: 'Agreement',
        title: "QC Intimation/Completion",
        summary:
          "The letters intimating the quality-control party about the work and its completion.",
        steps: [
          "Open Agreement and Work order. Upload the L1 selection form (and the Online Intimation, Supplier Details, Technical sanctioned Estimate if you have them) — or click “Fill details manually”.",
          "Click the “QC Intimation/Completion” tile.",
          "Check the QC party details and Work Order reference on the left.",
          "Review the live page, then Word, PDF or Print."
        ],
        mock: [
          { kind: 'upload', label: 'Upload L1 selection form', step: 1 },
          { kind: 'tiles', items: ['QC Intimation'], step: 2 },
          { kind: 'fields', items: ['Agency', 'Work Order ref', 'Date'], step: 3 },
          { kind: 'preview', label: 'QCC letter', step: 4 },
          { kind: 'button', label: 'Word / PDF / Print', step: 4 }
        ]
      }
    ]
  },
  {
    title: 'Utilities',
    features: [
      {
        id: 'tools',
        nav: 'Tools',
        title: 'Tools',
        summary: 'Stand-alone helpers, each opening as its own full-page workspace with a Back button.',
        steps: [
          'Open Tools and click a tile.',
          'Excel Sheet Separator — tick sheets and save selected or all as separate files.',
          'PDF / Word Merge-Separator — upload files, pick pages (numbers show the pick order), then merge into one or save each separately.',
          'Photos → PDF — combine photos or PDF pages into one PDF, borders auto-cropped.',
          'Photos / PDF → Word / Excel — OCR photos or a scanned PDF into editable text, export as .docx or .xlsx.',
          'GPS Photos — read latitude/longitude from each photo (EXIF or the stamped overlay) and export to Excel.',
          'Estimate from Photos / PDF, Work Order, Agreement Bond, Intimation, Schedule A, Electrical Estimate — the same generators as the workflow pages, usable on any circle’s files.'
        ],
        tips: ['Click “Back to Tools” at the top to return to the tiles.'],
        mock: [
          { kind: 'tiles', items: ['Excel Separator', 'PDF Merge', 'Word Merge', 'Photos → PDF', 'OCR → Word', 'GPS Photos'], step: 1 },
          { kind: 'upload', label: 'Upload files', step: 3 }
        ]
      },
      {
        id: 'cement',
        nav: 'Cement & Steel',
        title: 'Cement & Steel Rates',
        summary: 'Rate circulars fetched live from the Public Health department’s website.',
        steps: [
          'Open Cement & Steel Rates to browse the circulars.',
          'Download the one you need.',
          'The sidebar item blinks when a new circular appears; the app checks hourly and at each login.'
        ],
        mock: [{ kind: 'list', items: ['Circular — Oct 2026', 'Circular — Sep 2026'], step: 2 }, { kind: 'button', label: 'Download', step: 2 }]
      }
    ]
  },
  {
    title: 'System',
    features: [
      {
        id: 'settings',
        nav: 'Settings',
        title: 'Settings',
        summary: 'Themes, document templates, backup & restore and signed-in devices.',
        steps: [
          'Themes — choose Default or Aurora; it applies to the whole app.',
          'Document Templates — click a tile to make that style the default for the document; each tile is a live preview with sample data.',
          'Backup & Restore — save your whole workspace to a file, or restore it on this or another computer.',
          'Active Devices — see where you are signed in and log out any device.',
          'Check for updates from the bottom of the sidebar; “What’s New” shows the changes in each release.'
        ],
        mock: [
          { kind: 'tabs', items: ['Themes', 'Templates', 'Backup', 'Devices'], step: 1 },
          { kind: 'tiles', items: ['Style A', 'Style B', 'Style C'], step: 2 },
          { kind: 'button', label: 'Back up now', step: 3 }
        ]
      }
    ]
  }
]

// Real screenshots dropped into src/assets/manual/<feature-id>.png replace the illustration.
const SHOTS = import.meta.glob('../assets/manual/*.png', { eager: true, query: '?url', import: 'default' }) as Record<string, string>
function shotFor(id: string): string | undefined {
  const hit = Object.keys(SHOTS).find((k) => k.endsWith(`/${id}.png`))
  return hit ? SHOTS[hit] : undefined
}

function Badge({ n }: { n?: number }) {
  return n ? <span className="man-badge">{n}</span> : null
}

const SAMPLE: Record<string, string> = {
  'Work name': 'Improvement of CC road at Ward 12',
  Work: 'CC road, Ward 12',
  Agency: 'M/s Sri Sai Constructions',
  Amount: '₹ 24,50,000',
  Date: '27-10-2026',
  'Agreement no.': '58/2026-27',
  'Target date': '30-10-2026',
  Corporation: 'CMC',
  Zone: 'Quthbullapur',
  Circle: 'Circle 58'
}

/** pending = not reached yet, active = the step being demonstrated now, done = already demonstrated. */
type Phase = 'pending' | 'active' | 'done'

function MockBody({ el, phase }: { el: MockEl; phase: Phase }) {
  const on = phase !== 'pending'
  let body: JSX.Element
  switch (el.kind) {
    case 'button':
      body = <span className="man-mock-btn">{phase === 'done' ? `✓ ${el.label}` : el.label}</span>
      break
    case 'upload':
      body = (
        <span className="man-mock-upload">
          {phase === 'pending' ? `⬆ ${el.label}` : phase === 'active' ? '⏳ Reading document…' : '✓ File uploaded & read'}
          {on && <span className={`man-bar${phase === 'active' ? ' go' : ''}`} />}
        </span>
      )
      break
    case 'tiles':
      body = (
        <div className="man-mock-tiles">
          {el.items.map((t, i) => (
            <span key={t} className={`man-mock-tile${on && i === 0 ? ' sel' : ''}`}>{t}</span>
          ))}
        </div>
      )
      break
    case 'table':
      body = (
        <div className="man-mock-table" style={{ gridTemplateColumns: `repeat(${el.cols.length}, 1fr)` }}>
          {el.cols.map((c) => (
            <b key={c}>{c}</b>
          ))}
          {Array.from({ length: (el.rows ?? 3) * el.cols.length }).map((_, i) => (
            <i key={i} className={on ? 'full' : ''} style={{ transitionDelay: `${i * 60}ms` }} />
          ))}
        </div>
      )
      break
    case 'fields':
      body = (
        <div className="man-mock-fields">
          {el.items.map((f, i) => (
            <span key={f} className={`man-mock-field${on ? ' filled' : ''}`}>
              {on ? (
                <span className="man-type" style={{ animationDelay: `${i * 0.35}s` }}>{SAMPLE[f] ?? '✓ filled in'}</span>
              ) : (
                f
              )}
            </span>
          ))}
        </div>
      )
      break
    case 'preview':
      body = (
        <div className="man-mock-page">
          <span>{el.label}</span>
          {[0, 1, 2, 3, 4].map((i) => (
            <i key={i} className={on ? 'full' : ''} style={{ transitionDelay: `${i * 160}ms` }} />
          ))}
          {phase === 'done' && <em className="man-stamp">✓ ready</em>}
        </div>
      )
      break
    case 'tabs':
      body = (
        <div className="man-mock-tabs">
          {el.items.map((t, i) => (
            <span key={t} className={(on ? i === el.items.length - 1 : i === 0) ? 'on' : ''}>{t}</span>
          ))}
        </div>
      )
      break
    case 'list':
      body = (
        <div className="man-mock-list">
          {el.items.map((t, i) => (
            <span key={t} className={on ? 'hit' : ''} style={{ transitionDelay: `${i * 200}ms` }}>{t}</span>
          ))}
        </div>
      )
      break
  }
  return (
    <div className={`man-mock-row ph-${phase}`}>
      <Badge n={el.step} />
      {body}
    </div>
  )
}

/**
 * Plays a feature's steps in order: the pointer moves to each control, clicks it, and the control shows
 * its result (file read, fields filled, page previewed…). `step` is reported upward so the matching
 * numbered instruction lights up too.
 */
function Illustration({ f, onStep }: { f: Feature; onStep: (n: number) => void }) {
  const shot = shotFor(f.id)
  const ref = useRef<HTMLDivElement>(null)
  const cursorRef = useRef<SVGSVGElement>(null)
  const [visible, setVisible] = useState(false)
  const [idx, setIdx] = useState(-1)

  // Elements in the order the steps happen.
  const order = useMemo(
    () => f.mock.map((el, i) => ({ el, i })).sort((a, b) => (a.el.step ?? 99) - (b.el.step ?? 99)),
    [f]
  )
  const rank = (i: number) => order.findIndex((o) => o.i === i)

  useEffect(() => {
    const el = ref.current
    if (!el) return
    const io = new IntersectionObserver(([e]) => setVisible(e.isIntersecting), { threshold: 0.35 })
    io.observe(el)
    return () => io.disconnect()
  }, [])

  // Advance one step every 2.2s; hold on the finished state, then restart.
  useEffect(() => {
    if (!visible) {
      setIdx(-1)
      onStep(0)
      return
    }
    setIdx(0)
    let n = 0
    const t = setInterval(() => {
      n = n >= order.length + 1 ? 0 : n + 1
      setIdx(n >= order.length + 1 ? -1 : n)
    }, 2200)
    return () => clearInterval(t)
  }, [visible, order.length])

  useEffect(() => {
    onStep(idx >= 0 && idx < order.length ? order[idx].el.step ?? 0 : 0)
    const root = ref.current
    const cur = cursorRef.current
    if (!root || !cur) return
    const row = root.querySelectorAll<HTMLElement>('.man-mock-row')[order[idx]?.i ?? -1]
    if (!row || idx < 0 || idx >= order.length) {
      cur.style.opacity = '0'
      return
    }
    const r = row.getBoundingClientRect()
    const m = root.getBoundingClientRect()
    cur.style.opacity = '1'
    cur.style.transform = `translate(${r.left - m.left + Math.min(r.width * 0.6, 160)}px, ${r.top - m.top + r.height / 2}px)`
  }, [idx])

  if (shot) return <img className="man-shot" src={shot} alt={`${f.title} screenshot`} />
  const caption =
    idx < 0 ? 'Scroll here to watch the steps' : idx >= order.length ? 'Done ✓' : `Step ${order[idx].el.step}`
  return (
    <div ref={ref} className="man-mock" role="img" aria-label={`Animated demo of the ${f.title} page`}>
      <div className="man-mock-main">
        <div className="man-mock-crumb">
          <span>Sidebar</span> › <b>{f.nav}</b>
          <em className="man-caption">{caption}</em>
        </div>
        {f.mock.map((el, i) => {
          const r = rank(i)
          const phase: Phase = idx < 0 ? 'pending' : idx > r ? 'done' : idx === r ? 'active' : 'pending'
          return <MockBody key={i} el={el} phase={phase} />
        })}
        <svg ref={cursorRef} className="man-cursor" width="18" height="18" viewBox="0 0 24 24" aria-hidden="true">
          <path d="M5 3l14 8-6 2-2 6z" fill="#222" stroke="#fff" strokeWidth="1.5" strokeLinejoin="round" />
        </svg>
      </div>
    </div>
  )
}

function FeatureCard({ f }: { f: Feature }) {
  const [step, setStep] = useState(0)
  return (
    <article id={`man-${f.id}`} className="card man-feature">
      <h3>{f.title}</h3>
      <p className="man-summary">{f.summary}</p>
      <div className="man-body">
        <ol className="man-steps">
          {f.steps.map((s, i) => (
            <li key={i} className={step === i + 1 ? 'on' : ''}>
              <span className="man-badge">{i + 1}</span>
              <span>{s}</span>
            </li>
          ))}
        </ol>
        <Illustration f={f} onStep={setStep} />
      </div>
      {f.tips?.map((t) => (
        <p key={t} className="man-tip">💡 {t}</p>
      ))}
    </article>
  )
}

export default function ManualPage() {
  const [query, setQuery] = useState('')
  const q = query.trim().toLowerCase()
  const groups = useMemo(
    () =>
      GROUPS.map((g) => ({
        ...g,
        features: g.features.filter(
          (f) => !q || `${f.title} ${f.summary} ${f.steps.join(' ')}`.toLowerCase().includes(q)
        )
      })).filter((g) => g.features.length > 0),
    [q]
  )

  const jump = (id: string) => document.getElementById(`man-${id}`)?.scrollIntoView({ behavior: 'smooth', block: 'start' })

  return (
    <div className="man-layout">
      <nav className="card man-toc" aria-label="Guide contents">
        <input
          className="man-search"
          placeholder="Search the guide…"
          value={query}
          onChange={(e) => setQuery(e.target.value)}
        />
        {groups.map((g) => (
          <div key={g.title} className="man-toc-group">
            <div className="man-toc-title">{g.title}</div>
            {g.features.map((f) => (
              <button key={f.id} className="man-toc-link" onClick={() => jump(f.id)}>
                {f.title}
              </button>
            ))}
          </div>
        ))}
      </nav>

      <div className="man-content">
        <div className="card man-intro">
          <h2>What this app does</h2>
          <p>
            Head Draughtsman prepares the paperwork of a works office: it reads tender and estimate documents, fills the
            office’s letters, notices and agreements from them, and gives you a set of file tools — so each document is
            produced once, correctly, instead of being retyped. Below, every feature is explained step by step with a
            picture; the numbered circles in each picture match the numbered steps.
          </p>
        </div>
        {groups.length === 0 && <div className="card">No topic matches “{query}”.</div>}
        {groups.map((g) => (
          <section key={g.title}>
            <h2 className="man-group">{g.title}</h2>
            {g.features.map((f) => (
              <FeatureCard key={f.id} f={f} />
            ))}
          </section>
        ))}
      </div>
    </div>
  )
}
