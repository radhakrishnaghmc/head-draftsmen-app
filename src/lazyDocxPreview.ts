/**
 * Lazy, drop-in wrapper for docx-preview's `renderAsync`. docx-preview (and its
 * bundled dependencies) is one of the heaviest libraries in the app, but it's
 * only needed once the user renders a document (preview / print / thumbnail /
 * PDF) — never on the initial Works List screen. Importing it dynamically here
 * lets the bundler split it into its own chunk that loads on first use, so the
 * app window appears faster at startup.
 *
 * Call sites keep using `renderAsync(...)` exactly as before (they already
 * await it) — only their import path changes from 'docx-preview' to this file.
 */
type RenderAsync = typeof import('docx-preview')['renderAsync']

let scopeCounter = 0

/**
 * docx-preview writes each document's styles as page-global rules (`.docx p { text-indent… }`,
 * `.docx-wrapper { … }`) in a <style> next to the pages — so whichever document was rendered last
 * restyles every other preview on screen (a tile thumbnail's hanging-indent Normal style was
 * indenting and shifting the File Backer letterhead cells, for instance). Rewrite this render's
 * rules to apply only inside its own wrapper, tagged with a unique `data-dp`. `:where()` keeps the
 * specificity unchanged, so the app's own preview overrides still win exactly as before.
 */
function scopeDocxStyles(bodyContainer: HTMLElement, styleContainer: HTMLElement): void {
  const wrappers = bodyContainer.querySelectorAll<HTMLElement>('.docx-wrapper')
  const wrapper = wrappers[wrappers.length - 1]
  if (!wrapper) return
  const id = String(++scopeCounter)
  wrapper.setAttribute('data-dp', id)
  const scope = `:where(.docx-wrapper[data-dp="${id}"])`
  for (const style of Array.from(styleContainer.querySelectorAll('style'))) {
    if (style.hasAttribute('data-dp-scoped')) continue
    const rules = style.sheet?.cssRules
    if (!rules) continue
    const scoped = Array.from(rules)
      .map((rule) => {
        const styleRule = rule as CSSStyleRule // (duck-typed: the print snapshot's iframe has its own CSSStyleRule class)
        if (typeof styleRule.selectorText !== 'string' || !styleRule.style) return rule.cssText
        const selector = styleRule.selectorText
          .split(',')
          .map((s) => {
            const sel = s.trim()
            return sel.startsWith('.docx-wrapper') ? `.docx-wrapper:where([data-dp="${id}"])${sel.slice('.docx-wrapper'.length)}` : `${scope} ${sel}`
          })
          .join(', ')
        return `${selector} { ${styleRule.style.cssText} }`
      })
      .join('\n')
    style.textContent = scoped
    style.setAttribute('data-dp-scoped', id)
  }
}

export const renderAsync: RenderAsync = (async (...args: Parameters<RenderAsync>) => {
  const m = await import('docx-preview')
  const result = await m.renderAsync(...args)
  const body = args[1] as HTMLElement
  scopeDocxStyles(body, (args[2] as HTMLElement | undefined) ?? body)
  return result
}) as RenderAsync
