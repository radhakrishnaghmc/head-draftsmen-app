import { useEffect, useLayoutEffect, useRef, useState } from 'react'
import { api } from '../ipc'
import { guessHeaderRow } from '@core/sheet'
import { extractElectricalEstimateItems } from '@core/electricalEstimate'
import { extractWorkName } from '@core/estimateExtract'
import { computeEcvFromItems } from '../boqTransform'
import { formatRupees } from '@core/worksAmounts'
import { downloadBoqFromItems, downloadScheduleAFromItems } from '../estimateDownloads'
import { IconBolt, IconDownload, IconWarn, IconCheck } from './Icons'
import type { EstimateWorkItem } from '@core/estimateExtract'

function stripExt(name: string): string {
  return name.replace(/\.[^./\\]+$/, '')
}

/**
 * Tools-workspace tool: generate a BOQ and a Schedule A from a flat electrical
 * estimate, using the dedicated electrical extractor (core/electricalEstimate.ts)
 * rather than the civil detailed-estimate one. Standalone — nothing is read from
 * or written to the Works List, so the Schedule A's amounts come from the
 * estimate's own item total.
 */
export default function ElectricalEstimateTab({
  autoOpen = false,
  onContent
}: {
  autoOpen?: boolean
  onContent?: (hasContent: boolean) => void
}) {
  const [items, setItems] = useState<EstimateWorkItem[] | null>(null)
  const [workName, setWorkName] = useState<string | undefined>(undefined)
  const [fileBase, setFileBase] = useState('Electrical Estimate')
  const [error, setError] = useState<string | null>(null)
  const [busy, setBusy] = useState<null | 'upload' | 'boq' | 'scheduleA'>(null)
  const [saved, setSaved] = useState<string | null>(null)

  // Open the estimate picker straight away when launched from the Tools tile —
  // this is a single-upload tool, so no second click is needed. Guarded against
  // StrictMode's double-mount.
  const autoOpened = useRef(false)
  useEffect(() => {
    if (autoOpen && !autoOpened.current) {
      autoOpened.current = true
      void upload()
    }
  }, [autoOpen])

  // Report to the Tools host whether there's anything to show yet — before
  // paint, so the panel gets its full-width row without a one-frame flash.
  useLayoutEffect(() => {
    onContent?.(items !== null || error !== null)
  }, [items, error, onContent])

  async function upload() {
    setBusy('upload')
    setError(null)
    setSaved(null)
    setItems(null)
    try {
      const g = await api.pickEstimateGrid()
      if (!g) return
      const hr = guessHeaderRow(g.grid)
      const its = extractElectricalEstimateItems(g.grid, hr)
      if (its.length === 0) {
        throw new Error('No line items with a quantity and rate were found — is this a flat electrical estimate?')
      }
      setItems(its)
      setWorkName(extractWorkName(g.grid, hr))
      setFileBase(stripExt(g.name))
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e))
    } finally {
      setBusy(null)
    }
  }

  async function download(kind: 'boq' | 'scheduleA') {
    if (!items) return
    setBusy(kind)
    setError(null)
    setSaved(null)
    try {
      const path =
        kind === 'boq'
          ? await downloadBoqFromItems(items, workName, fileBase)
          : await downloadScheduleAFromItems(items, workName, null, fileBase)
      if (path) setSaved(path)
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e))
    } finally {
      setBusy(null)
    }
  }

  const ecv = items ? computeEcvFromItems(items) : 0

  // Launched from the Tools tile (autoOpen): show nothing until an estimate is
  // actually picked (or reading it fails) — the mount effect fires the folder,
  // so clicking the tile opens it directly with no placeholder panel.
  if (autoOpen && !items && !error) return null

  return (
    <div>
      <button className="primary upload-btn" onClick={upload} disabled={busy === 'upload'}>
        <IconBolt /> {busy === 'upload' ? 'Reading…' : items ? 'Change electrical estimate' : 'Upload electrical estimate'}
      </button>

      {error && (
        <div className="notice error">
          <IconWarn /> {error}
        </div>
      )}
      {saved && (
        <div className="notice ok">
          <IconCheck /> Saved {saved}
        </div>
      )}

      {items && (
        <div style={{ marginTop: 12 }}>
          <p className="estimate-hint">
            {items.length} item{items.length === 1 ? '' : 's'} · ECV (Σ Qty × Rate) {formatRupees(ecv)}
            {workName ? ` · ${workName.slice(0, 70)}` : ''}
          </p>
          <div className="boq-actions">
            <button className="primary" onClick={() => download('boq')} disabled={busy !== null}>
              <IconDownload /> Download BOQ
            </button>
            <button className="primary" onClick={() => download('scheduleA')} disabled={busy !== null}>
              <IconDownload /> Download Schedule A
            </button>
          </div>
          <div className="elec-preview">
            <table>
              <thead>
                <tr>
                  <th>#</th>
                  <th className="elec-desc">Description</th>
                  <th>Qty</th>
                  <th>Rate</th>
                  <th>Unit</th>
                </tr>
              </thead>
              <tbody>
                {items.map((it, i) => (
                  <tr key={i}>
                    <td>{i + 1}</td>
                    <td className="elec-desc">{it.description}</td>
                    <td>{it.quantity}</td>
                    <td>{it.rate}</td>
                    <td>{it.unit}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      )}
    </div>
  )
}
