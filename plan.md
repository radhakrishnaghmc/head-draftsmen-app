
---

## Addendum — Resilient fill + shareable build (shipped)

- **Crash-proof preview/generation** (`core/generate.ts`): `fillTemplate` now wraps the whole Docxtemplater path (constructor + render) in try/catch. On any parse error (documents containing stray `{{`/`}}` in body text, e.g. embedded JSON/logs), it falls back to `fillTemplateResilient` — a regex fill that only substitutes real `{{field}}` tokens matching data columns and leaves all other braces untouched. General, not doc-specific. Verified on Civil Tender + the two stray non-template files.
- **First-run seed data** (`resources/seed-state.json`, wired in `electron/main.ts` loadState): fresh installs load the works database (143 rows) but no template documents. Bundled via electron-builder `extraResources`.
- **Distribution/Gatekeeper**: packaged app was only linker-signed (resources not sealed) → "damaged/can't be opened" once transferred+quarantined. Fixed by deep ad-hoc sign (`codesign --force --deep --sign - --identifier com.docugen.app`), then DMG rebuilt via hdiutil → `release/ZonalDocs-Install.dmg`. Recipient still must clear quarantine once: `xattr -cr /Applications/DocuGen.app` OR System Settings → Privacy & Security → Open Anyway. Instructions in `release/OPEN-ME-FIRST.txt`.

Next steps (optional): universal (Intel+arm64) build for non-Apple-Silicon Macs; remove diagnostic logging (dlog/dbg/DROP/ALIGN) before final release; proper Developer-ID signing + notarization to remove the Gatekeeper step entirely.

---


New capability: create templates by editing any document in-app instead of hand-typing placeholders.

- **Import any file** (`core/import.ts`): `.docx` via mammoth, `.txt/.md/.csv/.html`, best-effort `.doc` → editable HTML.
- **In-app editor** (`TemplateEditor.tsx`): contentEditable document surface with a name field + Save.
- **Searchable, grouped field picker** (`ColumnPicker.tsx`): one Excel = flat column list; multiple Excels = drill-down by file; search box flattens across all files. Picking inserts a styled `{{Column}}` chip at the cursor.
- **Save as template** (`core/authoring.ts` + IPC `saveTemplate`): edited HTML → `.docx` (html-to-docx), chips normalized to literal `{{Column}}`, placeholders extracted, saved to `userData/templates`, added to the template list.
- Round-trip verified: authored HTML → docx → docxtemplater fill produces correct output.

UI: Templates card now has **New from document** (opens editor) alongside **Add .docx**.

---

## Addendum — In-app preview, tile icons, multi-value filter (shipped)

- **Tile action bar** (`GenerateDocsTab.tsx`, `styles.css`): moved Preview/Print/PDF from a hover-only overlay to an always-visible **bottom footer bar** with labels. Grid min column widened to 200px.
- **In-app .docx preview** (`DocPreview.tsx`): new modal renders the real file with `docx-preview` (high fidelity), reading raw bytes over a new IPC `readDocBytes` (`ipc-contract.ts`/`preload.ts`/`main.ts`). Eye/Preview button opens it instead of external Word. Preview tuned: full render options + let docx-preview draw native page chrome (only page-gap + shadow added in CSS).
- **Multi-value filter** (`GenerateDocsTab.tsx`, `App.tsx`): filter value is now a **multi-select** (`filterValues: string[]`) with removable chips. Rows match if their column value is in the selected set (empty = all). Generation already folders by `groupBy` (the filter column), so each selected value produces its own output subfolder automatically.
- **Calendar spacing**: replaced per-cell `aspect-ratio` with a fixed `34px` row height + `grid-auto-rows: 34px` — uniform weeks, removed the empty band under week 1.
- **Formatting**: confirmed the lossy `importDocument`→HTML warning path (`core/import.ts`) is now **unreachable** (TemplateEditor never mounts; `newFromDocument` uses the preserving `pickTemplates` pipeline). Remaining fidelity gap is the in-app preview render (no LibreOffice installed, so PDF-based preview isn't available; using docx-preview at max fidelity).

---

## Addendum — In-app docx editor: type + drag fields + remove pills (shipped)

`DocPreview.tsx` is now a real editor (not just a viewer) whenever opened with a
`tables` prop (Templates → Edit / New from document). The Generate-tab tile still
opens it read-only (no `tables`).

- **Direct text editing**: the docx-preview render surface is `contentEditable`.
  Typing saves back into the real `.docx` preserving formatting via a new core
  `setParagraphText` (`core/docx-edit.ts`) — a longest common prefix/suffix diff
  rewrites only the changed run(s), leaving every other run + its `rPr` intact.
  Saves are debounced (450 ms) with no re-render, so the caret is preserved.
- **Structure kept stable**: Enter (new paragraphs) and boundary Backspace/Delete
  (cross-paragraph merges) are blocked so the per-paragraph XML index mapping never
  drifts. Paste is forced to plain text. "Open in Word" remains for restructuring.
- **Removable field pills**: `stylizePills` wraps each `{{col}}` in an atomic,
  non-editable pill with an **×**. Clicking × calls new core `removePlaceholder`
  (drops the placeholder's own `w:r`). Occurrence is resolved by pill order within
  the paragraph. The pill's token text stays `{{col}}` (× excluded from saved text
  via `paragraphText`) so docxtemplater still fills it.
- **Drop reliability**: a drop resolves to caret → paragraph-under-cursor →
  **nearest paragraph**, so dropping anywhere on the page inserts. Field column is
  also parsed from `text/plain` as a fallback. Insert errors show as a red ⚠ chip.
- **New IPC**: `removeField`, `saveParagraph` (contract/preload/main), both write
  the file and return refreshed `{ base64, placeholders }`.
- **Print black-screen fix** (`electron/main.ts`): the print-preview `BrowserWindow`
  no longer sets `parent: mainWindow` (macOS compositor bug blacked out the main
  window on child close); both windows get `backgroundColor: '#fff'` and the main
  window is re-focused when the preview closes.
