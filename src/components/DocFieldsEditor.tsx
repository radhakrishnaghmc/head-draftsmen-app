import { useMemo, useState } from 'react'

export interface DocField {
  /** The template's {{Placeholder}} name. */
  label: string
  /** What the editor panel calls it. */
  title: string
  multiline?: boolean
  /** Sits on its own full-width row instead of sharing one. Multiline fields always do. */
  wide?: boolean
  /** A number + Months/Days toggle instead of a plain text box (e.g. "Period of Completion"), composed into "<n> Months" / "<n> Days". Defaults to Months. */
  periodUnit?: boolean
  /** A dropdown of fixed choices instead of a text box. */
  options?: { value: string; label: string }[]
  /** Hint shown in an empty text box (defaults to "Type to fill"). */
  placeholder?: string
}

const PERIOD_UNITS = ['Months', 'Days'] as const

/** Splits a "<n> Months"/"<n> Days" (or bare "<n>") value into its number and unit, defaulting to Months when the value carries no unit yet. */
function parsePeriod(value: string): { amount: string; unit: (typeof PERIOD_UNITS)[number] } {
  const m = value.trim().match(/^(\d+)\s*(Months?|Days?)?$/i)
  if (!m) return { amount: value.trim(), unit: 'Months' }
  const unit = /^day/i.test(m[2] ?? '') ? 'Days' : 'Months'
  return { amount: m[1], unit }
}

function composePeriod(amount: string, unit: (typeof PERIOD_UNITS)[number]): string {
  return amount.trim() ? `${amount.trim()} ${unit}` : ''
}

export interface DocFieldGroup {
  title: string
  fields: DocField[]
}

interface Props {
  groups: DocFieldGroup[]
  /** Values derived from the uploads / Works List row — what the panel shows when nothing is typed. */
  derived: Record<string, string>
  /** Hand-typed overrides, keyed by placeholder label. */
  edits: Record<string, string>
  onChange: (label: string, value: string) => void
  onRevert: (label: string) => void
  onResetAll: () => void
}

/** Turns a raw token ("NIT_No", "agencyName", "Corp Full Caps") into plain words for the panel. */
export function humanizeLabel(label: string): string {
  const spaced = label
    .replace(/[_\-.]+/g, ' ')
    .replace(/([a-z])([A-Z])/g, '$1 $2')
    .replace(/\s+/g, ' ')
    .trim()
  if (!spaced) return label
  return spaced === spaced.toUpperCase() && spaced.length > 4
    ? spaced.charAt(0) + spaced.slice(1).toLowerCase()
    : spaced.charAt(0).toUpperCase() + spaced.slice(1)
}

/** Clears earlier highlights and marks the first place the value is printed on the live page beside the panel. */
function highlightInPreview(value: string) {
  const root = document.querySelector('.ns-preview-scroll .intimation-docx-preview')
  if (!root) return
  root.querySelectorAll('mark.doc-hl').forEach((m) => {
    const parent = m.parentNode
    if (!parent) return
    parent.replaceChild(document.createTextNode(m.textContent ?? ''), m)
    parent.normalize()
  })
  const needle = value.trim()
  if (needle.length < 2) return
  const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT)
  let node: Node | null
  while ((node = walker.nextNode())) {
    const text = node.nodeValue ?? ''
    const at = text.indexOf(needle)
    if (at < 0) continue
    const range = document.createRange()
    range.setStart(node, at)
    range.setEnd(node, at + needle.length)
    const mark = document.createElement('mark')
    mark.className = 'doc-hl'
    range.surroundContents(mark)
    mark.scrollIntoView({ block: 'center', behavior: 'smooth' })
    return
  }
}

/** Groups a template's raw placeholder labels when a document has no hand-written field list. */
export function groupLabels(labels: string[]): DocFieldGroup[] {
  const bucket = (label: string): string => {
    if (/corp|circle|cno|zone|financial ?year|division|office/i.test(label)) return 'Office'
    if (/agency|contractor|address|phone|firm/i.test(label)) return 'Agency'
    if (/amount|ecv|value|percent|estimate|emd|asd|rupees|words|tp\b/i.test(label)) return 'Amounts'
    if (/date|no\.?$|nit|tender id|lr|ref|receipt/i.test(label)) return 'References and dates'
    return 'The work'
  }
  const order = ['Office', 'The work', 'Amounts', 'Agency', 'References and dates']
  const byTitle = new Map<string, DocField[]>()
  for (const label of labels) {
    const title = bucket(label)
    const fields = byTitle.get(title) ?? []
    fields.push({ label, title: humanizeLabel(label), multiline: /address|words|name of the work/i.test(label) })
    byTitle.set(title, fields)
  }
  return order
    .filter((t) => byTitle.has(t))
    .map((t) => ({ title: t, fields: byTitle.get(t) as DocField[] }))
}

/**
 * The fields pane of an editable document preview: every value the template
 * prints, grouped in the order the page prints them, beside the live page.
 *
 * A field starts on the value read from the uploads and stays tied to it until
 * it's typed in; from then on it's marked as a hand edit and can be put back on
 * its own. That distinction is the point of the panel — an office needs to see
 * at a glance which lines it has overruled before sending a document out.
 */
export default function DocFieldsEditor({ groups, derived, edits, onChange, onRevert, onResetAll }: Props) {
  const [onlyEmpty, setOnlyEmpty] = useState(false)
  const stats = useMemo(() => {
    let edited = 0
    let auto = 0
    let empty = 0
    for (const g of groups)
      for (const f of g.fields) {
        if (edits[f.label] !== undefined) edited++
        else if ((derived[f.label] ?? '').trim()) auto++
        if (!(edits[f.label] ?? derived[f.label] ?? '').trim()) empty++
      }
    return { edited, auto, empty }
  }, [groups, derived, edits])
  const editedCount = stats.edited

  return (
    <div className="ns-form doc-edit">
      <div className="doc-edit-head">
        <div className="doc-edit-title">Fill in the document</div>
        <span className="doc-edit-hint">
          Values are read from your uploads. Click a field to see where it prints on the page; type to change it —
          Word, PDF and Print carry your edits.
        </span>
        <div className="doc-edit-legend">
          <span className="doc-edit-chip auto">{stats.auto} auto-filled</span>
          <span className={`doc-edit-chip need ${stats.empty ? '' : 'zero'}`}>{stats.empty} need input</span>
          <span className={`doc-edit-chip edited ${stats.edited ? '' : 'zero'}`}>{stats.edited} edited</span>
        </div>
        <div className="doc-edit-tools">
          <label className="doc-edit-filter">
            <input type="checkbox" checked={onlyEmpty} onChange={(e) => setOnlyEmpty(e.target.checked)} />
            Show only fields that need input
          </label>
          {editedCount > 0 && (
            <button className="doc-edit-reset" onClick={onResetAll} title="Put every field back on the uploaded values">
              Reset {editedCount} edit{editedCount === 1 ? '' : 's'}
            </button>
          )}
        </div>
      </div>
      {groups.map((g) => {
        const edited = g.fields.filter((f) => edits[f.label] !== undefined).length
        const visible = onlyEmpty ? g.fields.filter((f) => !(edits[f.label] ?? derived[f.label] ?? '').trim()) : g.fields
        if (visible.length === 0) return null
        return (
          <section className="doc-edit-group" key={g.title}>
            <div className="ns-group-title">
              {g.title}
              {edited > 0 && <span className="doc-edit-count">{edited} edited</span>}
            </div>
            {visible.map((f) => {
              const isEdited = edits[f.label] !== undefined
              const value = edits[f.label] ?? derived[f.label] ?? ''
              const state = isEdited ? 'edited' : value.trim() ? 'auto' : 'need'
              return (
                <label
                  className={`field-label ns-field doc-edit-field ${f.multiline || f.wide ? 'wide' : ''} ${
                    isEdited ? 'edited' : ''
                  } state-${state}`}
                  key={f.label}
                  onFocus={() => highlightInPreview(value)}
                >
                  <span className="doc-edit-label">
                    <span className="doc-edit-name">
                      <i className={`doc-edit-dot ${state}`} aria-hidden="true" />
                      {f.title}
                      {state === 'need' && <em className="doc-edit-badge">Needs input</em>}
                    </span>
                    {isEdited && (
                      <button
                        className="doc-edit-revert"
                        onClick={(e) => {
                          e.preventDefault()
                          onRevert(f.label)
                        }}
                        title={`Put back "${derived[f.label] ?? ''}"`}
                      >
                        Revert
                      </button>
                    )}
                  </span>
                  {f.periodUnit ? (
                    (() => {
                      const { amount, unit } = parsePeriod(value)
                      return (
                        <span className="doc-edit-period">
                          <input
                            value={amount}
                            onChange={(e) => onChange(f.label, composePeriod(e.target.value, unit))}
                          />
                          <span className="doc-edit-period-toggle">
                            {PERIOD_UNITS.map((u) => (
                              <button
                                key={u}
                                type="button"
                                className={u === unit ? 'active' : ''}
                                onClick={() => onChange(f.label, composePeriod(amount, u))}
                              >
                                {u}
                              </button>
                            ))}
                          </span>
                        </span>
                      )
                    })()
                  ) : f.options ? (
                    <select value={value} onChange={(e) => onChange(f.label, e.target.value)}>
                      {f.options.map((o) => (
                        <option key={o.value} value={o.value}>
                          {o.label}
                        </option>
                      ))}
                    </select>
                  ) : f.multiline ? (
                    <textarea rows={2} value={value} onChange={(e) => onChange(f.label, e.target.value)} />
                  ) : (
                    <input value={value} placeholder={f.placeholder ?? 'Type to fill'} onChange={(e) => onChange(f.label, e.target.value)} />
                  )}
                </label>
              )
            })}
          </section>
        )
      })}
    </div>
  )
}
