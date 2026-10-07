import { useEffect, useMemo, useRef, useState, type ReactNode } from 'react'
import { noteSubmittedSections, type NoteSubmittedData, type NoteBidder } from '@core/noteSubmitted'
import { IconPlus } from './Icons'

interface Props {
  data: NoteSubmittedData
  onChange: (data: NoteSubmittedData) => void
}

/** Hand-writable blanks a note still prints — the ruled "______" lines the office fills in by pen. */
function blanksIn(html: string): number {
  return (html.match(/_{4,}/g) ?? []).length
}

/**
 * Controlled editor for a Note Submitted document.
 *
 * The sheet is six numbered notes, appended in the order the file moves
 * through the office, and note 4 refers back to the ones above it — so the
 * editor is built as that same sequence: an index of the six notes on the
 * left, one opening at a time to show only the fields that note prints, and
 * the sheet itself on the right with the note numbers in its margin the way
 * the paper file carries them. Picking a note on either side moves the other.
 *
 * The numbers and the index are the app's own chrome: the exported .docx is
 * built from the same section HTML joined back together, unchanged.
 */
export default function NoteSubmittedEditor({ data, onChange }: Props) {
  const sections = useMemo(() => noteSubmittedSections(data), [data])
  const blanks = useMemo(() => sections.map((s) => blanksIn(s.html)), [sections])
  const totalBlanks = blanks.reduce((a, b) => a + b, 0)
  // Open the first note that still prints a blank — that's the one needing
  // attention — otherwise start at the top of the sheet.
  const [openNote, setOpenNote] = useState(() => {
    const first = sections.findIndex((s) => blanksIn(s.html) > 0)
    return first < 0 ? 1 : first + 1
  })
  const paperRef = useRef<HTMLDivElement>(null)
  const pageRef = useRef<HTMLDivElement>(null)
  const noteRefs = useRef<Record<number, HTMLElement | null>>({})

  // A4 printable height in px (297mm minus the 6.35mm/1-line top+bottom
  // margins that match the Word doc, at 96dpi) — same geometry
  // printCreatedDocument prints this sheet at, so the count here matches
  // the real page count on paper.
  const A4_PAGE_PX = 1075
  const [pageCount, setPageCount] = useState(1)
  useEffect(() => {
    const measure = () => {
      const h = pageRef.current?.scrollHeight ?? 0
      setPageCount(Math.max(1, Math.ceil(h / A4_PAGE_PX)))
    }
    measure()
    window.addEventListener('resize', measure)
    return () => window.removeEventListener('resize', measure)
  }, [sections])

  function openAndScroll(n: number) {
    setOpenNote(n)
    const paper = paperRef.current
    const el = noteRefs.current[n]
    if (!paper || !el) return
    const smooth = !window.matchMedia('(prefers-reduced-motion: reduce)').matches
    paper.scrollTo({ top: Math.max(0, el.offsetTop - 16), behavior: smooth ? 'smooth' : 'auto' })
  }

  function setField<K extends keyof NoteSubmittedData>(key: K, value: NoteSubmittedData[K]) {
    onChange({ ...data, [key]: value })
  }
  function setBidder(i: number, key: keyof NoteBidder, value: string) {
    onChange({ ...data, bidders: data.bidders.map((b, j) => (j === i ? { ...b, [key]: value } : b)) })
  }
  function addBidder() {
    const n = data.bidders.length + 1
    onChange({ ...data, bidders: [...data.bidders, { sno: `${n}.`, name: '', ecv: '', pct: '', tcv: '', rank: `L-${n}` }] })
  }
  function removeBidder(i: number) {
    onChange({
      ...data,
      bidders: data.bidders.filter((_, j) => j !== i).map((b, j) => ({ ...b, sno: `${j + 1}.` }))
    })
  }

  const field = (label: string, key: keyof NoteSubmittedData) => (
    <label className="field-label ns-field" key={label}>
      {label}
      <input
        value={(data[key] as string) ?? ''}
        onChange={(e) => setField(key, e.target.value as NoteSubmittedData[typeof key])}
      />
    </label>
  )

  /** Points at the note that actually owns a value printed in more than one note. */
  const carriedFrom = (n: number, what: string) => (
    <p className="note-carried">
      {what}{' '}
      <button className="note-jump" onClick={() => openAndScroll(n)}>
        change it in note {n}
      </button>
      .
    </p>
  )

  const noteFields: Record<number, ReactNode> = {
    1: (
      <>
        <label className="field-label ns-field">
          Name of the work
          <input value={data.workName} onChange={(e) => setField('workName', e.target.value)} />
        </label>
        <div className="ns-grid">
          {field('Municipal body', 'body')}
          {field('Estimate (Lakhs)', 'estimateLakhs')}
          {field('Sanction date', 'asDate')}
          {field('Financial year', 'financialYear')}
        </div>
      </>
    ),
    2: (
      <>
        {/* The year of the data the estimate was scrutinised against — its own
            value, not note 1's budget year, since the two often differ. */}
        <div className="ns-grid">{field('Data year', 'dataYear')}</div>
        <p className="note-carried">The rest of this note&apos;s wording is fixed.</p>
      </>
    ),
    3: (
      <div className="ns-grid">
        {/* One reference fills both note 3's "tender Notice No" and note 4's
            "NIT No", so they can never diverge. */}
        <label className="field-label ns-field">
          Tender notice / NIT No
          <input
            value={data.tenderNoticeNo}
            onChange={(e) => onChange({ ...data, tenderNoticeNo: e.target.value, nitNo: e.target.value })}
          />
        </label>
        <label className="field-label ns-field">
          Tender notice / NIT date
          <input
            value={data.tenderNoticeDate}
            onChange={(e) => onChange({ ...data, tenderNoticeDate: e.target.value, nitDate: e.target.value })}
          />
        </label>
      </div>
    ),
    4: (
      <>
        {field('Newspapers the notice appeared in', 'newspapers')}
        <div className="ns-grid">
          <label className="field-label ns-field">
            Bidders rejected as non-responsive
            <input
              type="number"
              min={0}
              value={String(data.rejectedCount ?? 0)}
              onChange={(e) => setField('rejectedCount', Math.max(0, Math.trunc(Number(e.target.value) || 0)))}
            />
          </label>
          {field('Why they were rejected', 'qualificationNote')}
        </div>
        <div className="note-bidders-head">
          Price bid comparison
          <button className="ns-add" onClick={addBidder} title="Add a bidder row">
            <IconPlus /> Add bidder
          </button>
        </div>
        <div className="ns-bidders">
          <div className="ns-bidder-head">
            <span>S.No</span>
            <span>Name</span>
            <span>ECV</span>
            <span>% Quoted</span>
            <span>TCV</span>
            <span>Rank</span>
            <span />
          </div>
          {data.bidders.map((b, i) => (
            <div className="ns-bidder-row" key={i}>
              <input value={b.sno} onChange={(e) => setBidder(i, 'sno', e.target.value)} />
              <input value={b.name} onChange={(e) => setBidder(i, 'name', e.target.value)} />
              <input value={b.ecv} onChange={(e) => setBidder(i, 'ecv', e.target.value)} />
              <input value={b.pct} onChange={(e) => setBidder(i, 'pct', e.target.value)} />
              <input value={b.tcv} onChange={(e) => setBidder(i, 'tcv', e.target.value)} />
              <input value={b.rank} onChange={(e) => setBidder(i, 'rank', e.target.value)} />
              <button className="ns-del" onClick={() => removeBidder(i)} title="Remove this bidder">
                ✕
              </button>
            </div>
          ))}
          {data.bidders.length === 0 && (
            <p className="note-empty">No bidders yet. Add the rows the evaluation sheet shows, lowest first.</p>
          )}
        </div>
        <div className="ns-grid">
          {field('L-1 agency', 'l1Name')}
          {field('L-1 % quoted', 'l1PctText')}
          {field('L-1 TCV', 'l1Tcv')}
        </div>
      </>
    ),
    5: (
      <>
        <div className="ns-grid">
          {field('Circle the EE approves for', 'circle')}
          {field('Reservation (SC / ST / blank)', 'reservation')}
        </div>
        {carriedFrom(4, 'The letter goes to the L-1 agency —')}
      </>
    ),
    6: (
      <div className="ns-grid">
        {field('Intimation date', 'intimationDate')}
        {field('Online receipt No', 'receiptNo')}
        {field('Receipt date', 'receiptDate')}
      </div>
    )
  }

  return (
    <div className="note-sheet">
      <div className="note-index">
        <div className="note-index-head">
          <h3>Note sheet</h3>
          {totalBlanks > 0 && (
            <span className="note-blank-total">
              {totalBlanks} blank{totalBlanks === 1 ? '' : 's'} to fill by hand
            </span>
          )}
        </div>
        <ol className="note-list">
          {sections.map((s, i) => {
            const open = openNote === s.n
            return (
              <li key={s.n} className={open ? 'open' : ''}>
                <button className="note-row" onClick={() => openAndScroll(s.n)} aria-expanded={open}>
                  <span className="note-row-n">{s.n}</span>
                  <span className="note-row-title">{s.title}</span>
                  {s.n === 4 && data.bidders.length > 0 && (
                    <span className="note-row-meta">
                      {data.bidders.length} bidder{data.bidders.length === 1 ? '' : 's'}
                    </span>
                  )}
                  {blanks[i] > 0 && <span className="note-row-blank">{blanks[i]}</span>}
                </button>
                {open && <div className="note-fields">{noteFields[s.n]}</div>}
              </li>
            )
          })}
        </ol>
      </div>

      <div className="note-paper" ref={paperRef}>
        <div className="note-page-badge">
          A4 · {pageCount} page{pageCount === 1 ? '' : 's'}
        </div>
        <div className="note-page" ref={pageRef}>
          {sections.map((s) => (
            <section
              key={s.n}
              ref={(el) => {
                noteRefs.current[s.n] = el
              }}
              className={`note-block ${openNote === s.n ? 'active' : ''}`}
              onClick={() => setOpenNote(s.n)}
            >
              <span className="note-block-n" aria-hidden="true">
                {s.n}
              </span>
              <div className="note-block-body" dangerouslySetInnerHTML={{ __html: s.html }} />
            </section>
          ))}
        </div>
      </div>
    </div>
  )
}
