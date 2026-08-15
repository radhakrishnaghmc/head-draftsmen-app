import type { MouseEvent } from 'react'

/**
 * mousedown handler for a click-outside-to-close modal backdrop. Fires only
 * when the mousedown itself lands on the backdrop, not on a descendant —
 * unlike a plain onClick, it isn't fooled by a drag-selection that starts
 * inside the modal (e.g. selecting text in a field) and happens to release
 * over the backdrop, which would otherwise close the dialog mid-selection.
 */
export function closeOnBackdropMouseDown(onClose: () => void) {
  return (e: MouseEvent<HTMLDivElement>) => {
    if (e.target === e.currentTarget) onClose()
  }
}
