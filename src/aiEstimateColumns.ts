import { api } from './ipc'
import {
  extractEstimateItems,
  estimateColumnsMatchedViaEmbedding,
  ESTIMATE_COLUMN_SPECS
} from '@core/estimateExtract'
import type { EstimateWorkItem } from '@core/estimateExtract'
import type { ColumnEmbeddings } from '@core/columnMatch'

/**
 * Only called when the fast, synchronous regex-based column match already
 * failed — computes embeddings for an estimate's header row plus the
 * standard column labels (Serial Number/Quantity/Rate/Unit) so
 * extractEstimateItems gets one more chance via semantic similarity (e.g. a
 * column named "Nos." or "Approx Qty" no fixed pattern anticipated). Local
 * model only, no network — see electron/embeddingsWorker.ts. Returns null
 * (rather than throwing) if the model itself is unavailable, so the
 * caller's original, more specific error stands.
 */
async function embedEstimateColumns(header: string[]): Promise<ColumnEmbeddings | null> {
  try {
    const labels = ESTIMATE_COLUMN_SPECS.map((s) => s.label)
    const vectors = await api.embedTexts([...header, ...labels])
    return { headerVectors: vectors.slice(0, header.length), labelVectors: vectors.slice(header.length) }
  } catch {
    return null
  }
}

export interface EstimateItemsResult {
  items: EstimateWorkItem[]
  /** Column labels (Quantity/Rate/Unit/Serial Number) only resolved via semantic matching, not a plain header match — worth a "please double-check" notice. */
  aiAssisted: string[]
}

/**
 * extractEstimateItems with an automatic embedding-assisted retry: tries the
 * plain regex path first (fast, synchronous, no model dependency); only if
 * that throws does it compute embeddings for this estimate's header row and
 * retry once via semantic column matching. Rethrows the *original* error if
 * the retry still can't resolve every column (a more specific failure than
 * whatever the fallback attempt produced).
 */
export async function extractEstimateItemsWithAi(
  grid: string[][],
  headerRowIndex: number
): Promise<EstimateItemsResult> {
  try {
    return { items: extractEstimateItems(grid, headerRowIndex), aiAssisted: [] }
  } catch (syncError) {
    const header = grid[headerRowIndex] ?? []
    const embeddings = await embedEstimateColumns(header)
    if (!embeddings) throw syncError
    let items: EstimateWorkItem[]
    try {
      items = extractEstimateItems(grid, headerRowIndex, embeddings)
    } catch {
      throw syncError
    }
    return { items, aiAssisted: estimateColumnsMatchedViaEmbedding(grid, headerRowIndex, embeddings) }
  }
}
