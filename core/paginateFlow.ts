/// <reference lib="dom" />
/**
 * Cuts every sheet taller than its page into page-height sheets: the article's blocks are measured
 * and moved onto fresh copies of the sheet whenever the next one would overflow, and a table too tall
 * for the room left is split between rows (the tail keeps the same column widths). Run on a laid-out
 * (attached) container, since it reads rendered heights. Each further sheet repeats the section's header and footer.
 * Self-contained (no imports, no outer references) so the PDF converter can inject it into its hidden window via toString().
 */
export function paginateFlowSections(root: HTMLElement): void {
  const num = (v: string): number => parseFloat(v) || 0
  for (const sec of Array.from(root.querySelectorAll<HTMLElement>('section.docx'))) {
    const article = sec.querySelector<HTMLElement>(':scope > article')
    if (!article) continue
    const view = sec.ownerDocument.defaultView ?? window
    const doc = sec.ownerDocument
    const cs = view.getComputedStyle(sec)
    const pageH = num(cs.minHeight)
    // The section is a flex column: a header/footer in it takes page height away from the article.
    const chrome = Array.from(sec.querySelectorAll<HTMLElement>(':scope > header, :scope > footer')).reduce((n, el) => {
      const b = view.getComputedStyle(el)
      return n + el.getBoundingClientRect().height + num(b.marginTop) + num(b.marginBottom)
    }, 0)
    const room = pageH - num(cs.paddingTop) - num(cs.paddingBottom) - chrome
    if (!pageH || room <= 0 || article.offsetHeight <= room) continue

    const blockH = (el: HTMLElement): number => {
      const b = view.getComputedStyle(el)
      return el.getBoundingClientRect().height + num(b.marginTop) + num(b.marginBottom)
    }
    const pages: HTMLElement[][] = [[]]
    let used = 0
    const nextPage = (): void => {
      pages.push([])
      used = 0
    }
    for (const el of Array.from(article.children) as HTMLElement[]) {
      const h = blockH(el)
      if (used + h <= room || pages[pages.length - 1].length === 0 && h <= room) {
        pages[pages.length - 1].push(el)
        used += h
        continue
      }
      const rows = el.tagName === 'TABLE' ? Array.from(el.querySelectorAll<HTMLElement>(':scope > tbody > tr, :scope > tr')) : []
      if (rows.length > 1 && h > room - used) {
        // Spread the table's rows over the pages; each continuation is a copy of the table shell.
        const heights = rows.map((r) => r.getBoundingClientRect().height)
        const shell = h - heights.reduce((x, y) => x + y, 0)
        let target: HTMLElement = el
        let inTarget = 0
        let left = room - used - shell
        rows.forEach((row, i) => {
          if (heights[i] > left && inTarget > 0) {
            const cont = el.cloneNode(false) as HTMLElement
            const grid = el.querySelector(':scope > colgroup')
            if (grid) cont.appendChild(grid.cloneNode(true))
            cont.appendChild(doc.createElement('tbody'))
            nextPage()
            pages[pages.length - 1].push(cont)
            target = cont
            inTarget = 0
            left = room - shell
          } else if (inTarget === 0 && target === el) {
            pages[pages.length - 1].push(el)
          }
          if (target !== el) target.querySelector('tbody')!.appendChild(row)
          inTarget++
          left -= heights[i]
        })
        used = room - Math.max(left, 0)
        continue
      }
      nextPage()
      pages[pages.length - 1].push(el)
      used = h
    }

    // The first page keeps `sec`; each further page is an empty copy of it holding its blocks.
    let after = sec
    for (const blocks of pages.slice(1)) {
      const copy = sec.cloneNode(false) as HTMLElement
      const art = article.cloneNode(false) as HTMLElement
      for (const b of blocks) art.appendChild(b)
      // Word repeats the header/footer on every page; docx-preview draws them once per section.
      const header = sec.querySelector(':scope > header')
      if (header) copy.appendChild(header.cloneNode(true))
      copy.appendChild(art)
      const footer = sec.querySelector(':scope > footer')
      if (footer) copy.appendChild(footer.cloneNode(true))
      after.after(copy)
      after = copy
    }
  }
}
