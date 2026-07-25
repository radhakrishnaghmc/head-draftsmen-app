import libre from 'libreoffice-convert'

/**
 * Convert an RTF buffer (Word's own, much richer clipboard format — see
 * electron/main.ts's createDocumentFromClipboard) to a real .docx via a
 * local LibreOffice install, the same dependency core/docxToPdf.ts already
 * requires for PDF export. Explicitly named `source.rtf` (not the plain
 * `convert()` helper docxToPdf.ts uses) so LibreOffice's extension-based
 * filter selection reliably picks the RTF importer instead of guessing from
 * an extension-less temp file.
 */
export function convertRtfToDocx(rtfBuffer: Buffer): Promise<Buffer> {
  return new Promise((resolve, reject) => {
    libre.convertWithOptions(rtfBuffer, '.docx', undefined, { fileName: 'source.rtf' }, (err, data) => {
      if (err) {
        reject(
          new Error(
            'Pasting from Word requires LibreOffice to be installed on this computer. ' +
              'Install it from libreoffice.org, then try again.\n' +
              (err instanceof Error ? err.message : String(err))
          )
        )
        return
      }
      resolve(data)
    })
  })
}
