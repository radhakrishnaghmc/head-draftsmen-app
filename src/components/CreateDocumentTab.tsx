import { useEffect, useRef, useState } from 'react'
import { renderAsync } from 'docx-preview'
import { api } from '../ipc'
import { findPlaceholders } from '@core/createDocument'
import { base64ToUint8, DOCX_PREVIEW_OPTIONS } from './docPage'
import type { CreatedDocument } from '@core/types'
import { IconDoc, IconClipboard, IconPlus } from './Icons'

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

/**
 * Paste a document straight from Word (its own RTF clipboard data,
 * converted to a real .docx via a local LibreOffice install — see
 * electron/main.ts's createDocumentFromClipboard) and add {{Placeholder}}
 * markers wherever a value should be filled in later.
 *
 * The canvas below is docx-preview's own faithful render of that real .docx,
 * made contentEditable afterwards — there's no formatting toolbar, since the
 * whole point is that Word's own formatting survives untouched. Only plain
 * text edits are possible (typing a placeholder, fixing a typo); Enter is
 * blocked so the editor can't drift from the original's paragraph count,
 * which the diff-based save below relies on to know what changed.
 */
export default function CreateDocumentTab({ documents, onChange, editingDoc, onDoneEditing }: Props) {
  const editorRef = useRef<HTMLDivElement>(null)
  const [draftName, setDraftName] = useState('')
  const [detected, setDetected] = useState<string[]>([])
  const [savedNotice, setSavedNotice] = useState<string | null>(null)
  const [addError, setAddError] = useState<string | null>(null)
  const [pasteError, setPasteError] = useState<string | null>(null)
  const [pasting, setPasting] = useState(false)
  const [editingId, setEditingId] = useState<string | null>(null)
  // The pasted/loaded document's own base64 .docx, and its paragraphs' text
  // as they were at that moment — the anchor every save diffs the live,
  // possibly-edited canvas against to know which paragraphs actually changed.
  const [docxBase64, setDocxBase64] = useState<string | null>(null)
  const [originalParagraphs, setOriginalParagraphs] = useState<string[]>([])
  const [pageCount, setPageCount] = useState(0)

  function scanPlaceholders() {
    setDetected(findPlaceholders(editorRef.current?.innerText ?? ''))
  }

  function blockNewParagraphs(e: KeyboardEvent) {
    if (e.key === 'Enter') e.preventDefault()
  }

  async function renderDocx(base64: string) {
    const container = editorRef.current
    if (!container) return
    container.removeEventListener('input', scanPlaceholders)
    container.removeEventListener('keydown', blockNewParagraphs)
    container.innerHTML = ''
    await renderAsync(base64ToUint8(base64), container, undefined, DOCX_PREVIEW_OPTIONS)
    container.contentEditable = 'true'
    container.addEventListener('input', scanPlaceholders)
    container.addEventListener('keydown', blockNewParagraphs)
    setPageCount(container.querySelectorAll('section.docx').length)
    scanPlaceholders()
  }

  async function pasteFromWord() {
    setPasteError(null)
    setPasting(true)
    try {
      const base64 = await api.createDocumentFromClipboard()
      if (!base64) {
        setPasteError('Nothing to paste — copy some content in Word first, then try again.')
        return
      }
      const paragraphs = await api.listDocumentParagraphs(base64)
      setDocxBase64(base64)
      setOriginalParagraphs(paragraphs)
      setDraftName((prev) => prev || '')
      await renderDocx(base64)
    } catch (e) {
      setPasteError(e instanceof Error ? e.message : String(e))
    } finally {
      setPasting(false)
    }
  }

  // Load a document sent here from Issue Document for edits. Keyed on the
  // incoming doc's id (not the object itself) so this only re-fires when a
  // genuinely different document arrives, not on every parent re-render.
  useEffect(() => {
    if (!editingDoc) return
    let cancelled = false
    void (async () => {
      const paragraphs = await api.listDocumentParagraphs(editingDoc.docx)
      if (cancelled) return
      setDocxBase64(editingDoc.docx)
      setOriginalParagraphs(paragraphs)
      setDraftName(editingDoc.name)
      setEditingId(editingDoc.id)
      await renderDocx(editingDoc.docx)
    })()
    return () => {
      cancelled = true
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [editingDoc?.id])

  function resetCanvas() {
    if (editorRef.current) editorRef.current.innerHTML = ''
    setDraftName('')
    setDetected([])
    setDocxBase64(null)
    setOriginalParagraphs([])
    setPageCount(0)
  }

  function cancelEdit() {
    resetCanvas()
    setEditingId(null)
    onDoneEditing?.()
  }

  async function addToDocument() {
    setAddError(null)
    const container = editorRef.current
    const plainText = (container?.innerText ?? '').trim()
    if (!plainText || !docxBase64) {
      setAddError('Paste a document from Word first.')
      return
    }
    if (!draftName.trim()) {
      setAddError('Give this document a name before adding it.')
      return
    }
    const name = draftName.trim()

    const currentParagraphs = Array.from(container!.querySelectorAll('article p')).map((p) => p.textContent ?? '')
    const edits = currentParagraphs
      .map((text, index) => ({ index, text }))
      .filter(({ index, text }) => text !== (originalParagraphs[index] ?? ''))

    let finalDocx = docxBase64
    if (edits.length > 0) {
      try {
        finalDocx = await api.saveDocumentEdits(docxBase64, edits)
      } catch (e) {
        setAddError(e instanceof Error ? e.message : String(e))
        return
      }
    }

    if (editingId) {
      // Updating an existing document (sent here from Issue Document) —
      // keep its id/createdDate, only the name/docx change.
      onChange(documents.map((d) => (d.id === editingId ? { ...d, name, docx: finalDocx } : d)))
      setSavedNotice(`Updated "${name}". Find it on the Issue Document tab.`)
      setEditingId(null)
      onDoneEditing?.()
    } else {
      const doc: CreatedDocument = {
        id: `doc_${Date.now().toString(36)}`,
        name,
        docx: finalDocx,
        createdDate: today()
      }
      onChange([...documents, doc])
      setSavedNotice(`Added "${doc.name}" to Document. Find it on the Issue Document tab.`)
    }

    resetCanvas()
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
            Copy a document in Word, then click "Paste from Word" — formatting carries over exactly. Type{' '}
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

      <div className="boq-actions" style={{ padding: '0 4px 12px' }}>
        <button className="primary" onClick={pasteFromWord} disabled={pasting}>
          <IconClipboard /> {pasting ? 'Reading clipboard…' : 'Paste from Word'}
        </button>
      </div>
      {pasteError && <div className="notice error">{pasteError}</div>}

      <div className="doc-desk">
        <div className="doc-editor-wrap">
          <div ref={editorRef} className="docx-editor-canvas" />
          {pageCount > 1 && <span className="doc-page-badge">{pageCount} pages</span>}
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
        <button className="primary" onClick={addToDocument}>
          <IconPlus /> {editingId ? 'Save changes' : 'Add this to Document'}
        </button>
      </div>
      {addError && <div className="notice error">{addError}</div>}
      {savedNotice && <div className="notice">{savedNotice}</div>}
    </section>
  )
}
