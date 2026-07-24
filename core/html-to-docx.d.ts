declare module 'html-to-docx' {
  interface DocumentOptions {
    [key: string]: unknown
  }
  /**
   * Converts an HTML string to a .docx document.
   * In Node this resolves to a Buffer.
   */
  export default function HTMLtoDOCX(
    htmlString: string,
    headerHTMLString?: string,
    documentOptions?: DocumentOptions,
    footerHTMLString?: string
  ): Promise<Buffer | ArrayBuffer | Uint8Array>
}
