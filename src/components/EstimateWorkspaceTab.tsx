import EstimateUploadTab from './EstimateUploadTab'
import type { ExcelTable } from '@core/types'
import type { Office } from '../office'

interface Props {
  tables: ExcelTable[]
  onChange: (table: ExcelTable) => void
  /** The chosen office — its corporation is the default Department name on Material Quantity. */
  office: Office
}

/**
 * One workspace for everything derived from an estimate or a BOQ — BOQ,
 * Schedule A, Deviation Statement, and Material Quantity all come from the
 * same uploaded file's items, so there's no reason to upload it three or four
 * times across separate tabs. A flat BOQ is accepted here too (it just yields
 * its Schedule A / Material Quantity, with no estimate to re-derive from).
 * Reading an estimate from photos / a scanned PDF lives on the Tools page.
 */
export default function EstimateWorkspaceTab({ tables, onChange, office }: Props) {
  return <EstimateUploadTab tables={tables} onChange={onChange} office={office} />
}
