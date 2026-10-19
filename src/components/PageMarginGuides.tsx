import { useLayoutEffect, useRef, useState, type RefObject } from 'react'
import type { PageMargins } from '@core/docMargins'

const PX_PER_MM = 96 / 25.4
const SIDES = ['top', 'bottom', 'left', 'right'] as const

interface Props {
  /** The element docx-preview renders into; guides sit over its first page. */
  containerRef: RefObject<HTMLElement | null>
  /** Hand-set margins (mm); null = the template's own, which are read off the rendered page. */
  margins: PageMargins | null
  onChange: (m: PageMargins) => void
  /** Changes whenever the page is re-rendered, so the guides re-measure it. */
  renderKey: unknown
}

/**
 * Draggable dashed margin guides over a docx-preview page — the same control as
 * the Note Submitted sheet's. Must be rendered inside a `position: relative`
 * parent that also holds `containerRef`'s element.
 */
export default function PageMarginGuides({ containerRef, margins, onChange, renderKey }: Props) {
  const [box, setBox] = useState<{ left: number; top: number; width: number; height: number; scale: number } | null>(null)
  const [measured, setMeasured] = useState<PageMargins | null>(null)
  const [drag, setDrag] = useState<(typeof SIDES)[number] | null>(null)
  const dragRef = useRef<{ side: (typeof SIDES)[number]; start: number; value: number } | null>(null)

  useLayoutEffect(() => {
    const measure = (): void => {
      const sec = containerRef.current?.querySelector<HTMLElement>('section.docx')
      const parent = sec?.closest<HTMLElement>('.page-margin-host')
      if (!sec || !parent) {
        setBox(null)
        return
      }
      const s = sec.getBoundingClientRect()
      const p = parent.getBoundingClientRect()
      const scale = sec.offsetWidth ? s.width / sec.offsetWidth : 1
      const next = { left: s.left - p.left + parent.scrollLeft, top: s.top - p.top + parent.scrollTop, width: s.width, height: s.height, scale }
      setBox((prev) => (prev && Object.keys(next).every((k) => Math.abs(prev[k as keyof typeof next] - next[k as keyof typeof next]) < 0.5) ? prev : next))
      const cs = getComputedStyle(sec)
      const mm = (v: string): number => Math.round((parseFloat(v) / PX_PER_MM) * 10) / 10
      const m = { top: mm(cs.paddingTop), bottom: mm(cs.paddingBottom), left: mm(cs.paddingLeft), right: mm(cs.paddingRight) }
      setMeasured((prev) => (prev && SIDES.every((k) => prev[k] === m[k]) ? prev : m))
    }
    measure()
    // The page is swapped in after an offscreen render (and the modal can resize), often without the
    // page count changing — so re-measure on any DOM swap or size change, not just on renderKey.
    const container = containerRef.current
    if (!container) return
    const mo = new MutationObserver(measure)
    mo.observe(container, { childList: true })
    const ro = new ResizeObserver(measure)
    ro.observe(container)
    return () => {
      mo.disconnect()
      ro.disconnect()
    }
  }, [renderKey, containerRef])

  const eff = margins ?? measured
  if (!box || !eff) return null

  function startDrag(e: React.PointerEvent, side: (typeof SIDES)[number]) {
    e.preventDefault()
    e.currentTarget.setPointerCapture(e.pointerId)
    const horizontal = side === 'left' || side === 'right'
    dragRef.current = { side, start: horizontal ? e.clientX : e.clientY, value: eff![side] }
    setDrag(side)
  }
  function moveDrag(e: React.PointerEvent) {
    const d = dragRef.current
    if (!d || !box || !eff) return
    const horizontal = d.side === 'left' || d.side === 'right'
    const deltaMm = ((horizontal ? e.clientX : e.clientY) - d.start) / box.scale / PX_PER_MM
    const signed = d.side === 'top' || d.side === 'left' ? deltaMm : -deltaMm
    const next = Math.round(Math.min(80, Math.max(0, d.value + signed)) * 10) / 10
    if (next !== eff[d.side]) onChange({ ...eff, [d.side]: next })
  }
  function endDrag() {
    dragRef.current = null
    setDrag(null)
  }

  return (
    <div
      className="page-margin-guides"
      style={{ position: 'absolute', left: box.left, top: box.top, width: box.width, height: box.height, pointerEvents: 'none' }}
    >
      {SIDES.map((side) => (
        <div
          key={side}
          className={`note-guide note-guide-${side}${drag === side ? ' dragging' : ''}`}
          style={{ [side]: `${eff[side] * PX_PER_MM * box.scale}px`, pointerEvents: 'auto' }}
          title={`${side} margin — drag to adjust`}
          onPointerDown={(e) => startDrag(e, side)}
          onPointerMove={moveDrag}
          onPointerUp={endDrag}
          onPointerCancel={endDrag}
        >
          {drag === side && <span className="note-guide-tip">{eff[side].toFixed(1)} mm</span>}
        </div>
      ))}
    </div>
  )
}
