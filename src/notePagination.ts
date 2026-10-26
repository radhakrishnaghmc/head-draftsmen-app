// Note Submitted is plain HTML, so nothing paginates it for us. The on-screen
// editor and the print window both lay it out with THIS file: A4 sheets at the
// real print geometry (794 x 1122 px = 210 x 297 mm at 96 dpi; margins 6.35 mm
// top/bottom, 31.75 mm left/right) with whole notes assigned to pages. Because
// both draw the same sheets, the preview shows exactly what prints.

export const NOTE_SHEET = { w: 794, h: 1122 }

/** Page margins in millimetres. The defaults match the exported Word document. */
export interface NoteMargins {
  top: number
  bottom: number
  left: number
  right: number
}
export const DEFAULT_NOTE_MARGINS: NoteMargins = { top: 12, bottom: 8, left: 16, right: 16 }
/** The defaults before the sides/top were tightened — a saved copy of exactly these was never hand-set, so it upgrades. */
export const PREVIOUS_NOTE_MARGINS: NoteMargins = { top: 6.35, bottom: 6.35, left: 31.75, right: 31.75 }
const MM_PX = 96 / 25.4
const px = (mm: number) => Math.round(mm * MM_PX)
export const NOTE_FONT = "font-family:'Bookman Old Style',Georgia,serif;font-size:12pt;line-height:1.5;color:#000"

/** The part of one note that sits on one page (a whole note, or a run of its paragraphs). */
export interface NoteSeg {
  /** 0-based note index */
  n: number
  /** the note's HTML for this page, without padding */
  html: string
  /** this piece starts the note */
  first: boolean
  /** this piece ends the note (carries its sign-off) */
  last: boolean
}

/**
 * Cuts a note into pieces that may move to the next page on their own: its
 * paragraphs and tables. Two things stay glued so a page never ends badly — the
 * "Submitted: -" opener goes with the paragraph after it, and the HD/EE sign-off
 * (with its signing space) goes with the paragraph before it.
 */
interface NoteUnit {
  html: string
}

function noteUnits(html: string): NoteUnit[] {
  const box = document.createElement('div')
  box.innerHTML = html
  const kids = Array.from(box.children) as HTMLElement[]
  const units: NoteUnit[] = []
  let i = 0
  // sign-off = trailing borderless table holding HD/EE, plus the blank lines above it
  let signFrom = kids.length
  const lastKid = kids[kids.length - 1]
  if (lastKid && lastKid.tagName === 'TABLE' && /HD/.test(lastKid.textContent ?? '')) {
    signFrom = kids.length - 1
    while (signFrom > 0 && kids[signFrom - 1].tagName === 'P' && !(kids[signFrom - 1].textContent ?? '').replace(/\u00a0/g, '').trim()) signFrom--
  }
  const body = kids.slice(0, signFrom)
  const sign = kids.slice(signFrom).map((k) => k.outerHTML).join('')
  const rowTable = (t: HTMLTableElement, tr: HTMLTableRowElement) =>
    `${t.outerHTML.slice(0, t.outerHTML.indexOf('>') + 1)}<tbody>${tr.outerHTML}</tbody></table>`
  while (i < body.length) {
    const el = body[i]
    if (el.tagName === 'TABLE' && (el as HTMLTableElement).rows.length > 1) {
      // a table splits between its rows; the header is not repeated on the next page
      const t = el as HTMLTableElement
      const rows = Array.from(t.rows)
      rows.forEach((tr) => units.push({ html: rowTable(t, tr) }))
      i++
      continue
    }
    let chunk = el.outerHTML
    if (/^\s*Submitted:/.test(el.textContent ?? '') && i + 1 < body.length) {
      chunk += body[i + 1].outerHTML
      i++
    }
    units.push({ html: chunk })
    i++
  }
  if (units.length === 0) units.push({ html: '' })
  units[units.length - 1].html += sign
  return units
}

/** Joins consecutive one-row table pieces back into a single table. */
const mergeRows = (html: string) => html.replace(/<\/tbody><\/table>\s*<table[^>]*><tbody>/g, '')

/**
 * Lays the notes onto A4 pages, returning for each page the pieces of notes on it.
 * A note that doesn't fit in what is left of a page carries on onto the next one
 * between its paragraphs, instead of leaving the gap and jumping to a new page.
 *  */
export function paginateNotes(
  parts: string[],
  margins: NoteMargins = DEFAULT_NOTE_MARGINS
): NoteSeg[][] {
  const { w, h } = NOTE_SHEET
  const room = Math.max(100, h - px(margins.top) - px(margins.bottom))
  const probe = document.createElement('div')
  probe.style.cssText = `position:absolute;left:-10000px;top:0;visibility:hidden;${NOTE_FONT}`
  document.body.appendChild(probe)
  const measured = parts.map((html, n) => {
    const units = noteUnits(html)
    const box = document.createElement('div')
    box.style.width = `${Math.max(100, w - px(margins.left) - px(margins.right))}px`
    box.innerHTML = units.map((u) => `<div style="display:flow-root">${u.html}</div>`).join('')
    probe.appendChild(box)
    const heights = Array.from(box.children).map((c) => (c as HTMLElement).getBoundingClientRect().height)
    return { units, heights }
  })
  probe.remove()

  const pages: NoteSeg[][] = [[]]
  let used = 0
  measured.forEach(({ units, heights }, n) => {
    let run: string[] = []
    let runStart = 0
    const flush = (end: number) => {
      if (run.length) {
        pages[pages.length - 1].push({ n, html: mergeRows(run.join('')), first: runStart === 0, last: end === units.length })
      }
      run = []
    }
    units.forEach((u, k) => {
      if (used > 0 && used + heights[k] > room) {
        flush(k)
        pages.push([])
        used = 0
      }
      if (run.length === 0) runStart = k
      run.push(u.html)
      used += heights[k]
    })
    flush(units.length)
  })
  return pages.filter((pg, i) => pg.length > 0 || i === 0)
}

/** The sheets as self-contained HTML for the print window (inline styles only). */
export function noteSheetsHtml(pages: NoteSeg[][], margins: NoteMargins = DEFAULT_NOTE_MARGINS): string {
  const { w, h } = NOTE_SHEET
  const sheets = pages
    .map((segs, i) => {
      const last = i === pages.length - 1
      return (
        `<div class="note-a4" style="width:${w}px;min-height:${h}px;box-sizing:border-box;` +
        `padding:${px(margins.top)}px ${px(margins.right)}px ${px(margins.bottom)}px ${px(margins.left)}px;margin:0 auto 20px;background:#fff;` +
        `${last ? '' : 'break-after:page;'}">` +
        segs.map((sg) => `<div style="display:flow-root">${sg.html}</div>`).join('') +
        `</div>`
      )
    })
    .join('')
  return `<div class="note-a4s" style="${NOTE_FONT}">${sheets}</div>`
}

/** Page margins in twips (1/1440 in) for the .docx converter. */
export function noteMarginsTwips(m: NoteMargins): { top: number; bottom: number; left: number; right: number } {
  const tw = (mm: number) => Math.round((mm / 25.4) * 1440)
  return { top: tw(m.top), bottom: tw(m.bottom), left: tw(m.left), right: tw(m.right) }
}

/**
 * The notes as one HTML body for the Word/PDF export. No hard page breaks:
 * Word sets table rows and cell padding a little taller than the preview, so a
 * break forced where the preview turns the page can strand a near-empty page.
 * Left to flow, the export fills each page and keeps the preview's page count.
 */
export function noteDocxHtml(pages: NoteSeg[][]): string {
  // The export has no hard page breaks, so a table the preview split over two pages must be one table again
  // here — otherwise its two halves sit apart with a gap and their own column widths.
  const body = mergeRows(pages.map((segs) => segs.map((sg) => sg.html).join('')).join(''))
  return `<div style="${NOTE_FONT}">${body}</div>`
}
