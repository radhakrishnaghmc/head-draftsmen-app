import { useMemo } from 'react'

export interface DocField {
  /** The template's {{Placeholder}} name. */
  label: string
  /** What the editor panel calls it. */
  title: string
  multiline?: boolean
  /** Sits on its own full-width row instead of sharing one. Multiline fields always do. */
  wide?: boolean
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
    fields.push({ label, title: label, multiline: /address|words|name of the work/i.test(label) })
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
  const editedCount = useMemo(
    () => groups.reduce((n, g) => n + g.fields.filter((f) => edits[f.label] !== undefined).length, 0),
    [groups, edits]
  )

  return (
    <div className="ns-form doc-edit">
      <div className="doc-edit-head">
        <span className="doc-edit-hint">Type over any line — Word, PDF and Print carry your edits.</span>
        {editedCount > 0 ? (
          <button className="doc-edit-reset" onClick={onResetAll} title="Put every field back on the uploaded values">
            Reset {editedCount} edit{editedCount === 1 ? '' : 's'}
          </button>
        ) : (
          <span className="doc-edit-source">From your uploads</span>
        )}
      </div>
      {groups.map((g) => {
        const edited = g.fields.filter((f) => edits[f.label] !== undefined).length
        return (
          <section className="doc-edit-group" key={g.title}>
            <div className="ns-group-title">
              {g.title}
              {edited > 0 && <span className="doc-edit-count">{edited} edited</span>}
            </div>
            {g.fields.map((f) => {
              const isEdited = edits[f.label] !== undefined
              const value = edits[f.label] ?? derived[f.label] ?? ''
              return (
                <label
                  className={`field-label ns-field doc-edit-field ${f.multiline || f.wide ? 'wide' : ''} ${
                    isEdited ? 'edited' : ''
                  }`}
                  key={f.label}
                >
                  <span className="doc-edit-label">
                    {f.title}
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
                  {f.multiline ? (
                    <textarea rows={2} value={value} onChange={(e) => onChange(f.label, e.target.value)} />
                  ) : (
                    <input value={value} onChange={(e) => onChange(f.label, e.target.value)} />
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
