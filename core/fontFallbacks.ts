/**
 * docx-preview (the library that renders a filled .docx as HTML for the
 * on-screen preview and the Print action — see src/components/docPage.ts)
 * copies each run's font-family straight from the docx's own <w:rFonts>,
 * with no fallback listed (see its own parseFont: `style["font-family"] =
 * [...new Set(fonts)].join(', ')`, built only from the docx's own ascii/
 * asciiTheme/eastAsia names). Real Word substitutes a close-enough fallback
 * font when the exact one isn't installed; a browser with no fallback listed
 * falls back to its own arbitrary default instead, which can be wider and
 * wrap text (e.g. the Intimation letterhead's bold "CYBERABAD MUNICIPAL
 * CORPORATION" Calibri title wrapping to two lines) that fits fine in real
 * Word.
 *
 * This app's templates only ever declare Microsoft-only fonts (Calibri,
 * Segoe UI, Tahoma, Book Antiqua, Cambria, …) that ship with every Windows/
 * Office install — this app's real target — but are absent on macOS/Linux,
 * where this app is also built and tested. "Calibri" is the one that's
 * actually hit this (the corp-name titles on the header2 letterhead
 * variants); Carlito is Google's metric-compatible open substitute for it
 * (the same one LibreOffice itself substitutes internally, which is why the
 * LibreOffice-converted PDF export never showed this wrap — LibreOffice's
 * copy is private to its own app bundle, invisible to Electron's Chromium).
 * Embedding it as a data-URI @font-face makes it work regardless of what's
 * installed on the machine, in both the on-screen preview and the Print
 * window (the generated <style> tag travels with the captured
 * container.innerHTML into printCreatedDocument's HTML).
 *
 * "Gautami" is the mirror-image case: it's the complex-script font the
 * letterhead's Telugu caption declares (see core/docxToPdf.ts's own doc
 * comment on PDF_FONT_SUBSTITUTES), and unlike Calibri it IS installed at
 * the OS level on Windows (shipped there for Indic-script support) but not
 * on macOS/Linux. So without a pinned substitute, the preview renders real
 * Gautami on Windows and some other, narrower, auto-picked Telugu-capable
 * font on Mac/Linux — fitting one line on Mac but wrapping on Windows.
 * Pinning both platforms to the same embedded "Noto Sans Telugu" (the font
 * already used for this run in the LibreOffice/PDF path) makes the preview
 * match the PDF, on every OS, regardless of what's actually installed.
 */
export const FONT_SUBSTITUTES: Record<string, string> = {
  calibri: 'Carlito',
  gautami: 'Noto Sans Telugu'
}

/** Builds the @font-face CSS for the given base64-encoded Carlito/Noto Sans Telugu TTFs (see api.fontFallbackFiles). */
export function buildFontFallbackCss(fonts: { regular: string; bold: string; telugu: string }): string {
  return [
    '@font-face {',
    "  font-family: 'Carlito';",
    `  src: url(data:font/truetype;base64,${fonts.regular}) format('truetype');`,
    '  font-weight: normal;',
    '  font-style: normal;',
    '}',
    '@font-face {',
    "  font-family: 'Carlito';",
    `  src: url(data:font/truetype;base64,${fonts.bold}) format('truetype');`,
    '  font-weight: bold;',
    '  font-style: normal;',
    '}',
    '@font-face {',
    "  font-family: 'Noto Sans Telugu';",
    `  src: url(data:font/truetype;base64,${fonts.telugu}) format('truetype');`,
    '  font-weight: normal;',
    '  font-style: normal;',
    '}'
  ].join('\n')
}
