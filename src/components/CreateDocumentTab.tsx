import { useEffect, useRef, useState } from 'react'
import type { MouseEvent as ReactMouseEvent } from 'react'
import { findPlaceholders } from '@core/createDocument'
import { cleanPastedHtml } from '@core/pasteClean'
import type { CreatedDocument } from '@core/types'
import { IconDoc, IconPlus, IconEye } from './Icons'
import { pageShellStyle, usePageScrollTracker, PAGE_WIDTH } from './docPage'

// The base document written into the editable iframe. No stylesheet of
// ours loaded beyond pageShellStyle (which only sets up the page-on-a-desk
// look) — pasted content renders using the browser's own defaults plus
// whatever styling it brought with it, so our app's CSS can't clobber a
// pasted paragraph's margin/line-height/fonts the way it would if the same
// markup sat inside a div on this page.
const EDITOR_SHELL = `<!DOCTYPE html><html><head><meta charset="utf-8"><style>${pageShellStyle()}</style></head><body></body></html>`

const FONT_FAMILIES = [
  'Times New Roman',
  'Arial',
  'Calibri',
  'Georgia',
  'Verdana',
  'Courier New'
]
const FONT_SIZES = ['8', '9', '10', '11', '12', '14', '16', '18', '20', '24', '28', '32']

interface Props {
  documents: CreatedDocument[]
  onChange: (docs: CreatedDocument[]) => void
  /** Set when a document was sent here from Issue Document for edits — loads
   *  its content into the canvas, and saving updates it in place instead of
   *  adding a new entry. */
  editingDoc?: CreatedDocument | null
  /** Called once the incoming editingDoc has been loaded (or the edit was
   *  cancelled/saved), so the parent can clear it and not re-trigger the load. */
  onDoneEditing?: () => void
}

function today(): string {
  const d = new Date()
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`
}

export default function CreateDocumentTab({ documents, onChange, editingDoc, onDoneEditing }: Props) {
  const frameRef = useRef<HTMLIFrameElement>(null)
  // The editor's selection at the moment a toolbar control was last
  // interacted with — see captureSelection/focusAndRestoreSelection below.
  const savedRangeRef = useRef<Range | null>(null)
  const [draftName, setDraftName] = useState('')
  const [detected, setDetected] = useState<string[]>([])
  const [savedNotice, setSavedNotice] = useState<string | null>(null)
  const [addError, setAddError] = useState<string | null>(null)
  const [shellReady, setShellReady] = useState(false)
  const [previewHtml, setPreviewHtml] = useState<string | null>(null)
  const [previewError, setPreviewError] = useState<string | null>(null)
  const previewFrameRef = useRef<HTMLIFrameElement>(null)
  const [editingId, setEditingId] = useState<string | null>(null)

  // Only attach once the shell below has actually written the iframe's
  // document — attaching earlier would observe/listen on the throwaway
  // about:blank document that doc.write() then replaces.
  const { current: currentPage, total: totalPages } = usePageScrollTracker(frameRef, shellReady)
  const { current: previewPage, total: previewPages } = usePageScrollTracker(previewFrameRef, previewHtml)

  function scanCanvas() {
    const text = frameRef.current?.contentDocument?.body.innerText ?? ''
    setDetected(findPlaceholders(text))
  }

  // Set up the editable surface once: a real, isolated HTML document (not a
  // div on this page) with editing turned on. Pasting into it uses the
  // browser's normal rich-paste handling, but the result renders with only
  // the browser's own defaults — nothing here can override a pasted
  // paragraph's font, size, spacing, or table borders.
  useEffect(() => {
    const frame = frameRef.current
    const doc = frame?.contentDocument
    if (!frame || !doc) return
    doc.open()
    doc.write(EDITOR_SHELL)
    doc.close()
    doc.designMode = 'on'
    const onEdit = () => scanCanvas()

    // Intercept paste ourselves instead of letting the browser insert the
    // clipboard HTML unfiltered — Word's clipboard fragment carries forced
    // page breaks and runs of empty spacer paragraphs that, dropped into a
    // continuous (non-paginated) document, show up as extra blank lines and
    // content that visually lands on the "wrong" page.
    const onPaste = (e: ClipboardEvent) => {
      const html = e.clipboardData?.getData('text/html')
      const text = e.clipboardData?.getData('text/plain')
      e.preventDefault()
      if (html) {
        doc.execCommand('insertHTML', false, cleanPastedHtml(html))
      } else if (text) {
        doc.execCommand('insertText', false, text)
      }
      onEdit()
    }

    doc.body.addEventListener('input', onEdit)
    doc.body.addEventListener('paste', onPaste)
    setShellReady(true)
    return () => {
      doc.body.removeEventListener('input', onEdit)
      doc.body.removeEventListener('paste', onPaste)
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [])

  // Load a document sent here from Issue Document for edits. Keyed on the
  // incoming doc's id (not the object itself) so this only re-fires when a
  // genuinely different document arrives, not on every parent re-render.
  useEffect(() => {
    if (!shellReady || !editingDoc) return
    const body = frameRef.current?.contentDocument?.body
    if (!body) return
    body.innerHTML = editingDoc.html
    setDraftName(editingDoc.name)
    setEditingId(editingDoc.id)
    scanCanvas()
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [shellReady, editingDoc?.id])

  function cancelEdit() {
    const body = frameRef.current?.contentDocument?.body
    if (body) body.innerHTML = ''
    setDraftName('')
    setDetected([])
    setEditingId(null)
    onDoneEditing?.()
  }

  // Grabs whatever's currently selected in the editor before a toolbar
  // control (about to steal focus) runs its own default mousedown behavior —
  // for a <select>, that default behavior is opening its native dropdown,
  // which calling preventDefault() would block entirely (unlike a <button>,
  // where preventDefault only stops the focus-steal, not the click itself).
  // So selects can't use preserveSelection below; they capture the range
  // here instead and focusAndRestoreSelection() puts it back once the user
  // has actually picked a value.
  function captureSelection() {
    const sel = frameRef.current?.contentDocument?.getSelection()
    if (sel && sel.rangeCount > 0) savedRangeRef.current = sel.getRangeAt(0).cloneRange()
  }

  // Re-focuses the iframe and restores the last-captured selection —
  // execCommand needs the target document to actually be focused, and needs
  // the selection to still be there, to have any effect. A no-op restore
  // when nothing moved focus away (the <button> case, where
  // preserveSelection already kept focus in the iframe).
  function focusAndRestoreSelection() {
    const doc = frameRef.current?.contentDocument
    frameRef.current?.contentWindow?.focus()
    doc?.body.focus()
    const sel = doc?.getSelection()
    if (sel && savedRangeRef.current) {
      sel.removeAllRanges()
      sel.addRange(savedRangeRef.current)
    }
  }

  function exec(command: string, value?: string) {
    focusAndRestoreSelection()
    frameRef.current?.contentDocument?.execCommand(command, false, value)
    scanCanvas()
  }

  // execCommand('fontSize') only supports the legacy HTML levels 1-7, not
  // real point sizes — force a level-7 <font> wrapper around the selection,
  // then rewrite it to the actual pt size requested.
  function setFontSize(pt: string) {
    focusAndRestoreSelection()
    const doc = frameRef.current?.contentDocument
    if (!doc) return
    doc.execCommand('fontSize', false, '7')
    doc.body.querySelectorAll('font[size="7"]').forEach((el) => {
      el.removeAttribute('size')
      ;(el as HTMLElement).style.fontSize = `${pt}pt`
    })
  }

  // Prevents the toolbar button from stealing focus (and the selection)
  // away from the iframe when clicked — execCommand needs that selection
  // to still be there when it runs. Also captures the selection so it can
  // be explicitly restored in exec() (see focusAndRestoreSelection) — a
  // no-op here since focus never actually left, but keeps one code path.
  function preserveSelection(e: ReactMouseEvent) {
    e.preventDefault()
    captureSelection()
  }

  function openPreview() {
    setPreviewError(null)
    const body = frameRef.current?.contentDocument?.body
    if (!(body?.innerText ?? '').trim()) {
      setPreviewError('Paste some document content first.')
      return
    }
    setPreviewHtml(body?.innerHTML ?? '')
  }

  function addToDocument() {
    setAddError(null)
    const body = frameRef.current?.contentDocument?.body
    const html = body?.innerHTML ?? ''
    const plainText = (body?.innerText ?? '').trim()
    if (!plainText) {
      setAddError('Paste some document content first.')
      return
    }
    if (!draftName.trim()) {
      setAddError('Give this document a name before adding it.')
      return
    }
    const name = draftName.trim()
    if (editingId) {
      // Updating an existing document (sent here from Issue Document) —
      // keep its id/createdDate, only the name/html change.
      onChange(documents.map((d) => (d.id === editingId ? { ...d, name, html } : d)))
      setSavedNotice(`Updated "${name}". Find it on the Issue Document tab.`)
      setEditingId(null)
      onDoneEditing?.()
    } else {
      const doc: CreatedDocument = {
        id: `doc_${Date.now().toString(36)}`,
        name,
        html,
        createdDate: today()
      }
      onChange([...documents, doc])
      setSavedNotice(`Added "${doc.name}" to Document. Find it on the Issue Document tab.`)
    }
    if (body) body.innerHTML = ''
    setDraftName('')
    setDetected([])
    setTimeout(() => setSavedNotice(null), 4000)
  }

  return (
    <section className="card">
      <div className="card-head">
        <div className="head-ic">
          <IconDoc />
        </div>
        <div className="titles">
          <h2>Paste a new document</h2>
          <p className="sub">
            Paste content from Word or anywhere else — formatting carries over. Add{' '}
            <code>{'{{Placeholder}}'}</code> markers wherever a value should be filled in, e.g.{' '}
            <code>{'{{Name of the work}}'}</code>. When you're happy with it, name it and add it below.
          </p>
        </div>
      </div>

      {editingId && (
        <div className="notice">
          Editing "{draftName}" — sent from Issue Document.{' '}
          <button className="ghost" onClick={cancelEdit}>
            Cancel
          </button>
        </div>
      )}

      <div className="editor-toolbar">
        <button className="tb-btn" style={{ fontWeight: 700 }} onMouseDown={preserveSelection} onClick={() => exec('bold')}>
          B
        </button>
        <button className="tb-btn" style={{ fontStyle: 'italic' }} onMouseDown={preserveSelection} onClick={() => exec('italic')}>
          I
        </button>
        <button
          className="tb-btn"
          style={{ textDecoration: 'underline' }}
          onMouseDown={preserveSelection}
          onClick={() => exec('underline')}
        >
          U
        </button>

        <span className="tb-sep" />

        <select className="tb-select" defaultValue="" onMouseDown={captureSelection} onChange={(e) => exec('fontName', e.target.value)}>
          <option value="" disabled>
            Font
          </option>
          {FONT_FAMILIES.map((f) => (
            <option value={f} key={f} style={{ fontFamily: f }}>
              {f}
            </option>
          ))}
        </select>

        <select className="tb-select" defaultValue="" onMouseDown={captureSelection} onChange={(e) => setFontSize(e.target.value)}>
          <option value="" disabled>
            Size
          </option>
          {FONT_SIZES.map((s) => (
            <option value={s} key={s}>
              {s} pt
            </option>
          ))}
        </select>

        <span className="tb-sep" />

        <select
          className="tb-select"
          defaultValue=""
          onMouseDown={captureSelection}
          onChange={(e) => exec(e.target.value)}
        >
          <option value="" disabled>
            Align
          </option>
          <option value="justifyLeft">Left</option>
          <option value="justifyCenter">Center</option>
          <option value="justifyRight">Right</option>
          <option value="justifyFull">Justify</option>
        </select>

        <span className="tb-sep" />

        <button className="tb-btn" onMouseDown={preserveSelection} onClick={() => exec('insertUnorderedList')}>
          • List
        </button>
        <button className="tb-btn" onMouseDown={preserveSelection} onClick={() => exec('insertOrderedList')}>
          1. List
        </button>

        <span className="tb-sep" />

        <button className="tb-btn" onMouseDown={preserveSelection} onClick={() => exec('undo')}>
          Undo
        </button>
        <button className="tb-btn" onMouseDown={preserveSelection} onClick={() => exec('redo')}>
          Redo
        </button>
      </div>

      <div className="doc-desk">
        <div className="doc-editor-wrap" style={{ width: PAGE_WIDTH }}>
          <iframe ref={frameRef} className="doc-editor" title="Document canvas" />
          {totalPages > 1 && (
            <span className="doc-page-badge">
              Page {currentPage} of {totalPages}
            </span>
          )}
        </div>
      </div>

      {detected.length > 0 && (
        <div className="tags" style={{ padding: '0 4px' }}>
          {detected.map((d) => (
            <span className="tag accent" key={d}>
              {'{{' + d + '}}'}
            </span>
          ))}
        </div>
      )}

      <div className="card-head-actions" style={{ padding: '12px 4px 4px' }}>
        <input
          className="editor-name"
          placeholder="Document name (e.g. Agreement Bond)"
          value={draftName}
          onChange={(e) => setDraftName(e.target.value)}
        />
        <button className="ghost" onClick={openPreview}>
          <IconEye /> Preview
        </button>
        <button className="primary" onClick={addToDocument}>
          <IconPlus /> {editingId ? 'Save changes' : 'Add this to Document'}
        </button>
      </div>
      {addError && <div className="notice error">{addError}</div>}
      {previewError && <div className="notice error">{previewError}</div>}
      {savedNotice && <div className="notice">{savedNotice}</div>}

      {previewHtml !== null && (
        <div className="editor-overlay" onClick={() => setPreviewHtml(null)}>
          <div className="editor-modal" onClick={(e) => e.stopPropagation()}>
            <div className="editor-head">
              <span className="editor-title">Preview</span>
              <div className="editor-head-actions">
                <button className="ghost" onClick={() => setPreviewHtml(null)}>
                  Close
                </button>
              </div>
            </div>
            <div className="doc-desk">
              <div className="doc-editor-wrap" style={{ width: PAGE_WIDTH }}>
                <iframe
                  ref={previewFrameRef}
                  className="doc-editor"
                  title="Document preview"
                  srcDoc={`<!DOCTYPE html><html><head><meta charset="utf-8"><style>${pageShellStyle()}</style></head><body>${previewHtml}</body></html>`}
                />
                {previewPages > 1 && (
                  <span className="doc-page-badge">
                    Page {previewPage} of {previewPages}
                  </span>
                )}
              </div>
            </div>
          </div>
        </div>
      )}
    </section>
  )
}
