import type { QcOfficeParties, QcPartyDetails } from '@core/types'

const EMPTY_PARTY: QcPartyDetails = { name: '', address: '', phone: '' }
const EMPTY_PARTIES: QcOfficeParties = { third: EMPTY_PARTY, fourth: EMPTY_PARTY }

interface Props {
  /** The current office's department QC / 3rd / 4th-party agencies (undefined = none set yet). */
  parties: QcOfficeParties | undefined
  onChange: (next: QcOfficeParties) => void
  /** Only a circle (EE) office issues these letters; disable for a zone-only office. */
  circleSelected: boolean
}

/**
 * Enter the department QC, 3rd-party (QC college) and 4th-party (testing lab) agency contact
 * details once, per office, right under the office picker on the Works List. The
 * 3rd/4th-party QC letters fill their "To" block from here instead of asking
 * every time (see PrintDocumentTab), and so does the Agreement/Work Order page's
 * five-letter QC Intimation/Completion set (see qcPartyToBlock) — these agencies
 * change yearly, not per work.
 */
export default function QcPartiesEditor({ parties, onChange, circleSelected }: Props) {
  const p = parties ?? EMPTY_PARTIES
  const set = (which: 'department' | 'third' | 'fourth', field: keyof QcPartyDetails, value: string) =>
    onChange({ ...EMPTY_PARTIES, ...p, [which]: { ...p[which], [field]: value } })

  if (!circleSelected) return null

  const block = (which: 'department' | 'third' | 'fourth', title: string) => {
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
          Address block (printed under “To” — one line per row)
          <textarea
            rows={5}
            value={v.address}
            placeholder={'e.g. Professor & Head,\nDepartment of Civil Engineering,\nVNR Vignan Jyothi Institute…\nPragathi Nagar, Nizampet\nHyderabad, Telangana'}
            onChange={(e) => set(which, 'address', e.target.value)}
          />
        </label>
        <label>
          Phone / Mobile No.
          <input type="text" value={v.phone} onChange={(e) => set(which, 'phone', e.target.value)} />
        </label>
      </div>
    )
  }

  return (
    <div className="card qcp">
      <div className="qcp-grid">
        {block('department', 'Department QC')}
        {block('third', '3rd Party QC (inspection college)')}
        {block('fourth', '4th Party (testing lab)')}
      </div>
    </div>
  )
}
