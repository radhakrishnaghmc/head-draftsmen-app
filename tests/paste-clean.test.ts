import { describe, it, expect } from 'vitest'
import { cleanPastedHtml } from '../core/pasteClean'

describe('cleanPastedHtml', () => {
  it('strips Word conditional comment blocks', () => {
    const html = '<p>Keep me</p><!--[if gte mso 9]><xml>junk</xml><![endif]-->'
    expect(cleanPastedHtml(html)).toBe('<p>Keep me</p>')
  })

  it('inlines a class rule from <style> onto the matching element and drops the <style> block', () => {
    // Real bug: Word defines paragraph spacing in this <style> block (e.g.
    // .MsoNormal margin), not inline on each <p>. Whether a <style> tag
    // survives the browser's insertHTML fragment-insertion path isn't
    // reliable, so relying on it silently drops that spacing — paragraphs
    // fall back to the browser's much larger default margins and a pasted
    // document measures out far longer than the original. Inlining the
    // declaration straight onto the element's own style attribute removes
    // that uncertainty entirely.
    const html =
      '<style>.MsoNormal{margin:0in 0in 8.0pt;font-size:11.0pt}</style><xml><o:DocumentProperties/></xml><p class="MsoNormal">Body text</p>'
    expect(cleanPastedHtml(html)).toBe(
      '<p class="MsoNormal" style="margin:0in 0in 8.0pt;font-size:11.0pt">Body text</p>'
    )
  })

  it('inlines a comma-separated selector list (Word\'s p.MsoNormal, li.MsoNormal, div.MsoNormal pattern)', () => {
    const html = '<style>p.MsoNormal, li.MsoNormal, div.MsoNormal{margin:0 0 8pt}</style><p class="MsoNormal">Text</p>'
    expect(cleanPastedHtml(html)).toBe('<p class="MsoNormal" style="margin:0 0 8pt">Text</p>')
  })

  it('merges an inlined class rule with an existing inline style, keeping the original style winning on conflicts', () => {
    const html = '<style>.MsoNormal{margin:0 0 8pt;color:red}</style><p class="MsoNormal" style="color:blue">Text</p>'
    expect(cleanPastedHtml(html)).toBe('<p class="MsoNormal" style="margin:0 0 8pt;color:red;color:blue">Text</p>')
  })

  it('ignores a body/html/* rule in <style> — nothing to inline it onto', () => {
    const html = '<style>body{font-family:Calibri} .MsoNormal{margin:0 0 8pt}</style><p class="MsoNormal">Text</p>'
    expect(cleanPastedHtml(html)).toBe('<p class="MsoNormal" style="margin:0 0 8pt">Text</p>')
  })

  it('leaves an element with no matching class rule untouched', () => {
    const html = '<style>.MsoNormal{margin:0 0 8pt}</style><p class="Other">Text</p>'
    expect(cleanPastedHtml(html)).toBe('<p class="Other">Text</p>')
  })

  it('inlines onto an unquoted class attribute (Word\'s real clipboard markup very commonly omits quotes)', () => {
    // Real bug: the first version of this only matched class="...", so it
    // silently skipped inlining on most real Word paragraphs, which is
    // exactly the case that produced a document measuring out at roughly
    // double its real page count.
    const html = '<style>.MsoNormal{margin:0 0 8pt}</style><p class=MsoNormal>Text</p>'
    expect(cleanPastedHtml(html)).toBe('<p class=MsoNormal style="margin:0 0 8pt">Text</p>')
  })

  it('merges onto an existing unquoted style attribute', () => {
    const html = '<style>.MsoNormal{margin:0 0 8pt}</style><p class=MsoNormal style=color:blue>Text</p>'
    expect(cleanPastedHtml(html)).toBe('<p class=MsoNormal style="margin:0 0 8pt;color:blue">Text</p>')
  })

  it('converts embedded double quotes in a declaration (e.g. font-family:"Calibri") so the style attribute is not corrupted', () => {
    // Real bug: a quoted font name left as-is would prematurely close the
    // style="..." attribute it gets embedded in, corrupting the markup.
    const html = '<style>.MsoNormal{font-family:"Calibri",sans-serif}</style><p class="MsoNormal">Text</p>'
    expect(cleanPastedHtml(html)).toBe('<p class="MsoNormal" style="font-family:\'Calibri\',sans-serif">Text</p>')
  })

  it('removes a forced page-break <br>', () => {
    const html = '<p>Page one</p><br clear="all" style="page-break-before:always"><p>Page two</p>'
    expect(cleanPastedHtml(html)).toBe('<p>Page one</p><p>Page two</p>')
  })

  it('removes a forced page-break div', () => {
    const html = '<p>Page one</p><div style="page-break-before:always"></div><p>Page two</p>'
    expect(cleanPastedHtml(html)).toBe('<p>Page one</p><p>Page two</p>')
  })

  it('strips a page-break style property without removing the paragraph it sits on', () => {
    const html = '<p style="page-break-before:always">Section heading</p>'
    expect(cleanPastedHtml(html)).toBe('<p style="">Section heading</p>')
  })

  it('collapses a run of empty spacer paragraphs down to one', () => {
    const html = '<p>Before</p><p>&nbsp;</p><p>&nbsp;</p><p>&nbsp;</p><p>After</p>'
    expect(cleanPastedHtml(html)).toBe('<p>Before</p><p>&nbsp;</p><p>After</p>')
  })

  it('leaves a single empty paragraph alone', () => {
    const html = '<p>Before</p><p>&nbsp;</p><p>After</p>'
    expect(cleanPastedHtml(html)).toBe(html)
  })

  it('does not touch a paragraph that has real content', () => {
    const html = '<p>Before</p><p>Still real text</p><p>After</p>'
    expect(cleanPastedHtml(html)).toBe(html)
  })

  it('collapses a long run of <br> tags down to two', () => {
    const html = '<p>Before</p><br><br><br><br><br><p>After</p>'
    expect(cleanPastedHtml(html)).toBe('<p>Before</p><br><br><p>After</p>')
  })
})
