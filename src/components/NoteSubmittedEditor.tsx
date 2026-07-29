import { useMemo } from 'react'
import { buildNoteSubmittedHtml, type NoteSubmittedData, type NoteBidder } from '@core/noteSubmitted'
import { IconPlus } from './Icons'

interface Props {
  data: NoteSubmittedData
  onChange: (data: NoteSubmittedData) => void
}

/**
 * Controlled editor for a Note Submitted document: the caller owns the
 * NoteSubmittedData (seeded from a Works List row + the uploaded evaluation
 * PDF's bidder table) and this renders the editable fields alongside a live
 * preview. Every field stays correctable by hand before export.
 */
export default function NoteSubmittedEditor({ data, onChange }: Props) {
  const previewHtml = useMemo(() => buildNoteSubmittedHtml(data), [data])

  function setField<K extends keyof NoteSubmittedData>(key: K, value: NoteSubmittedData[K]) {
    onChange({ ...data, [key]: value })
  }
  function setBidder(i: number, key: keyof NoteBidder, value: string) {
    onChange({ ...data, bidders: data.bidders.map((b, j) => (j === i ? { ...b, [key]: value } : b)) })
  }
  function addBidder() {
    const n = data.bidders.length + 1
    onChange({ ...data, bidders: [...data.bidders, { sno: `${n}.`, name: '', ecv: '', pct: '', tcv: '', rank: `L-${n}` }] })
  }
  function removeBidder(i: number) {
    onChange({
      ...data,
      bidders: data.bidders.filter((_, j) => j !== i).map((b, j) => ({ ...b, sno: `${j + 1}.` }))
    })
  }

  const field = (label: string, key: keyof NoteSubmittedData) => (
    <label className="field-label ns-field">
      {label}
      <input
        value={(data[key] as string) ?? ''}
        onChange={(e) => setField(key, e.target.value as NoteSubmittedData[typeof key])}
      />
    </label>
  )

  return (
    <div className="ns-body">
      <div className="ns-form">
        <div className="ns-group-title">Work &amp; sanction</div>
        <div className="ns-grid">
          {field('Municipal body (CMC / GHMC)', 'body')}
          {field('Circle', 'circle')}
          {field('Estimate (Lakhs)', 'estimateLakhs')}
          {field('Admin sanction date', 'asDate')}
          {field('Financial year', 'financialYear')}
          {field('Reservation (SC / ST / blank)', 'reservation')}
        </div>
        <label className="field-label ns-field">
          Name of the work
          <input value={data.workName} onChange={(e) => setField('workName', e.target.value)} />
        </label>

        <div className="ns-group-title">Tender</div>
        <div className="ns-grid">
          {field('Tender Notice No (Lr No)', 'tenderNoticeNo')}
          {field('Tender notice date', 'tenderNoticeDate')}
          {field('NIT No', 'nitNo')}
          {field('NIT date', 'nitDate')}
        </div>
        {field('Newspapers', 'newspapers')}
        {field('Qualification / rejection note (optional)', 'qualificationNote')}

        <div className="ns-group-title">
          Bidders
          <button className="ns-add" onClick={addBidder} title="Add a bidder row">
            <IconPlus /> Add
          </button>
        </div>
        <div className="ns-bidders">
          <div className="ns-bidder-head">
            <span>S.No</span>
            <span>Name</span>
            <span>ECV</span>
            <span>% Quoted</span>
            <span>TCV</span>
            <span>Rank</span>
            <span />
          </div>
          {data.bidders.map((b, i) => (
            <div className="ns-bidder-row" key={i}>
              <input value={b.sno} onChange={(e) => setBidder(i, 'sno', e.target.value)} />
              <input value={b.name} onChange={(e) => setBidder(i, 'name', e.target.value)} />
              <input value={b.ecv} onChange={(e) => setBidder(i, 'ecv', e.target.value)} />
              <input value={b.pct} onChange={(e) => setBidder(i, 'pct', e.target.value)} />
              <input value={b.tcv} onChange={(e) => setBidder(i, 'tcv', e.target.value)} />
              <input value={b.rank} onChange={(e) => setBidder(i, 'rank', e.target.value)} />
              <button className="ns-del" onClick={() => removeBidder(i)} title="Remove">
                ✕
              </button>
            </div>
          ))}
        </div>

        <div className="ns-group-title">Agreement (L-1)</div>
        <div className="ns-grid">
          {field('L-1 agency', 'l1Name')}
          {field('L-1 % quoted', 'l1PctText')}
          {field('L-1 TCV', 'l1Tcv')}
          {field('Intimation date', 'intimationDate')}
          {field('Online Receipt No', 'receiptNo')}
          {field('Receipt date', 'receiptDate')}
        </div>
      </div>

      <div className="ns-preview">
        <div className="ns-preview-scroll">
          <div className="ns-doc" dangerouslySetInnerHTML={{ __html: previewHtml }} />
        </div>
      </div>
    </div>
  )
}
