import type { QcOfficeParties, QcPartyDetails } from '@core/types'

const EMPTY_PARTY: QcPartyDetails = { name: '', address: '', phone: '' }
const EMPTY_PARTIES: QcOfficeParties = { third: EMPTY_PARTY, fourth: EMPTY_PARTY }

interface Props {
  /** The current office's 3rd/4th-party agencies (undefined = none set yet). */
  parties: QcOfficeParties | undefined
  onChange: (next: QcOfficeParties) => void
  /** Only a circle (EE) office issues these letters; disable for a zone-only office. */
  circleSelected: boolean
}

/**
 * Enter the 3rd-party (QC college) and 4th-party (testing lab) agency contact
 * details once, per office, right under the office picker on the Works List. The
 * 3rd/4th-party QC letters fill their "To" block from here instead of asking
 * every time (see PrintDocumentTab) — these agencies change yearly, not per work.
 */
export default function QcPartiesEditor({ parties, onChange, circleSelected }: Props) {
  const p = parties ?? EMPTY_PARTIES
  const set = (which: 'third' | 'fourth', field: keyof QcPartyDetails, value: string) =>
    onChange({ ...EMPTY_PARTIES, ...p, [which]: { ...p[which], [field]: value } })

  if (!circleSelected) return null

  const block = (which: 'third' | 'fourth', title: string) => {
    const v = p[which] ?? EMPTY_PARTY
    return (
      <div className="qcp-block">
        <h4>{title}</h4>
        <label>
          Agency name
          <input
            type="text"
            value={v.name}
            placeholder="e.g. Marri Laxman Reddy Institute…"
            onChange={(e) => set(which, 'name', e.target.value)}
          />
        </label>
        <label>
          Address
          <textarea rows={2} value={v.address} onChange={(e) => set(which, 'address', e.target.value)} />
        </label>
        <label>
          Phone / Mobile No.
          <input type="text" value={v.phone} onChange={(e) => set(which, 'phone', e.target.value)} />
        </label>
      </div>
    )
  }

  const anySet = !!(p.third?.name?.trim() || p.fourth?.name?.trim())
  return (
    <details className="qcp card" open={!anySet}>
      <summary>
        3rd / 4th-party QC agency details
        <span className="qcp-hint"> — enter once here; the 3rd/4th-party QC letters fill their address from it</span>
      </summary>
      <div className="qcp-grid">
        {block('third', '3rd Party QC (inspection college)')}
        {block('fourth', '4th Party (testing lab)')}
      </div>
    </details>
  )
}
