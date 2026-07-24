/**
 * Clean up Word-clipboard artifacts before inserting pasted HTML into a
 * continuous (non-paginated) document. Word's clipboard fragment carries
 * things that only make sense inside Word's own paginated layout — forced
 * page breaks and stacks of empty spacer paragraphs it uses to push content
 * onto the next printed page — which, dropped into a document with no real
 * pagination, just render as stray blank lines and forced gaps. This strips
 * those out while leaving the actual visible content and its formatting
 * untouched.
 */
export function cleanPastedHtml(html: string): string {
  let out = html

  // Word's conditional-comment blocks (VML image fallbacks, MSO metadata)
  // are never meant to be visible — safe to drop entirely.
  out = out.replace(/<!--\[if[^\]]*\]>[\s\S]*?<!\[endif\]-->/gi, '')
  out = out.replace(/<!\[if[^\]]*\]>[\s\S]*?<!\[endif\]>/gi, '')

  // Word's <xml> metadata block (document properties etc.) — pure
  // metadata, never needed.
  out = out.replace(/<xml>[\s\S]*?<\/xml>/gi, '')

  out = inlineClassStyles(out)

  // Explicit forced page breaks — meaningless without real pagination, and
  // otherwise render as a stray forced blank line.
  out = out.replace(/<br[^>]*style="[^"]*page-break[^"]*"[^>]*>/gi, '')
  out = out.replace(/<div[^>]*style="[^"]*page-break[^"]*"[^>]*>\s*<\/div>/gi, '')
  out = out.replace(/(?:mso-)?page-break-(?:before|after)\s*:\s*[a-z]+;?/gi, '')
  out = out.replace(/mso-special-character\s*:\s*line-break;?/gi, '')

  // Collapse a run of two or more empty paragraphs (Word's vertical-spacing
  // padding, meant to fill out a page) down to a single blank line.
  const emptyPara = '<p[^>]*>(?:\\s|&nbsp;|<span[^>]*>(?:\\s|&nbsp;)*</span>|<o:p>\\s*</o:p>)*</p>'
  const emptyParaRun = new RegExp(`(?:${emptyPara}\\s*){2,}`, 'gi')
  const emptyParaOne = new RegExp(emptyPara, 'i')
  out = out.replace(emptyParaRun, (run) => emptyParaOne.exec(run)?.[0] ?? run)

  // Collapse long runs of <br> (also used for the same vertical-spacing
  // purpose) down to at most two.
  out = out.replace(/(?:<br[^>]*>\s*){3,}/gi, '<br><br>')

  return out
}

/**
 * Word usually defines actual paragraph spacing (e.g. ".MsoNormal
 * {margin:0in 0in 8.0pt}") in a <style> block, referenced by class, rather
 * than inline on each <p>. Whether a browser's insertHTML/execCommand
 * fragment-insertion path reliably keeps a <style> tag intact is not
 * something we can rely on — so instead of hoping it survives, this reads
 * every class rule out of the <style> block, writes its declarations
 * directly onto the `style` attribute of each element with a matching
 * class, and removes the <style> block entirely. Inline attributes on an
 * element are guaranteed to survive any HTML fragment insertion, since
 * they're part of the element itself rather than separate metadata — this
 * is what actually fixed a real pasted document rendering roughly double
 * its real page count (paragraphs silently losing Word's tight spacing and
 * falling back to the browser's much larger default paragraph margins).
 */
function inlineClassStyles(html: string): string {
  const styleMatch = html.match(/<style[^>]*>([\s\S]*?)<\/style>/i)
  const out = html.replace(/<style[^>]*>[\s\S]*?<\/style>/gi, '')
  if (!styleMatch) return out

  const classRules = extractClassRules(styleMatch[1])
  if (classRules.size === 0) return out

  // Word's clipboard markup very commonly leaves attribute values
  // unquoted (e.g. `class=MsoNormal`, not `class="MsoNormal"`) — matching
  // only the quoted form silently skipped inlining on most real paragraphs.
  const ATTR_VALUE = `(?:"([^"]*)"|'([^']*)'|([^\\s>]+))`
  const classRe = new RegExp(`\\bclass\\s*=\\s*${ATTR_VALUE}`, 'i')
  const styleRe = new RegExp(`\\bstyle\\s*=\\s*${ATTR_VALUE}`, 'i')

  return out.replace(/<(\w+)((?:\s[^<>]*)?)>/g, (full, tag: string, attrs: string) => {
    const classMatch = attrs.match(classRe)
    if (!classMatch) return full
    const classValue = classMatch[1] ?? classMatch[2] ?? classMatch[3] ?? ''
    const declarations = classValue
      .split(/\s+/)
      .filter(Boolean)
      .map((cls) => classRules.get(cls))
      .filter((d): d is string => !!d)
    if (declarations.length === 0) return full
    // CSS values commonly carry double quotes of their own (e.g.
    // font-family:"Calibri") — left as-is, one would prematurely close the
    // style="..." attribute this gets embedded into and corrupt the
    // markup. Single quotes are equally valid CSS string syntax.
    const combined = declarations.join(';').replace(/"/g, "'")

    const styleMatch2 = attrs.match(styleRe)
    const existingStyle = styleMatch2
      ? (styleMatch2[1] ?? styleMatch2[2] ?? styleMatch2[3] ?? '').replace(/"/g, "'")
      : null
    const newAttrs =
      existingStyle !== null
        ? attrs.replace(styleRe, `style="${combined};${existingStyle}"`)
        : `${attrs} style="${combined}"`
    return `<${tag}${newAttrs}>`
  })
}

/** Maps each class name to its declaration text, for simple `.Foo` and `tag.Foo` selectors (including comma-separated lists like Word's `p.MsoNormal, li.MsoNormal, div.MsoNormal`). Descendant/pseudo selectors are skipped — not a pattern Word's own clipboard CSS uses for paragraph spacing. */
function extractClassRules(css: string): Map<string, string> {
  const map = new Map<string, string>()
  const ruleRe = /([^{}]+)\{([^{}]*)\}/g
  let m: RegExpExecArray | null
  while ((m = ruleRe.exec(css)) !== null) {
    const decl = m[2].trim().replace(/;\s*$/, '')
    if (!decl) continue
    // A selector list like Word's "p.MsoNormal, li.MsoNormal, div.MsoNormal"
    // can name the same class more than once — apply this rule's
    // declaration to each distinct class only once.
    const classesInRule = new Set<string>()
    for (const rawSelector of m[1].split(',')) {
      const classMatch = rawSelector.trim().match(/^[a-zA-Z0-9]*\.([\w-]+)$/)
      if (classMatch) classesInRule.add(classMatch[1])
    }
    for (const cls of classesInRule) {
      const existing = map.get(cls)
      map.set(cls, existing ? `${existing};${decl}` : decl)
    }
  }
  return map
}
