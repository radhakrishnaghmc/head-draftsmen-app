import type {
  ExcelTable,
  MergedDataset,
  Collision,
  CollisionResolution,
  ExcelColumn
} from './types'

/**
 * Detect columns that appear in more than one Excel file.
 * A collision means the same header name is present in >= 2 sources.
 */
export function detectCollisions(tables: ExcelTable[]): Collision[] {
  const columnSources = new Map<string, Set<string>>()
  for (const table of tables) {
    for (const header of table.headers) {
      if (!columnSources.has(header)) columnSources.set(header, new Set())
      columnSources.get(header)!.add(table.name)
    }
  }
  const collisions: Collision[] = []
  for (const [column, sources] of columnSources) {
    if (sources.size > 1) {
      collisions.push({ column, sources: [...sources] })
    }
  }
  return collisions
}

/**
 * Merge multiple Excel tables into a single dataset by ROW INDEX:
 * record n is built from row n of every file. Columns are unioned.
 *
 * When a column collides across files, `resolution` selects the winning
 * source file; if a column has an unresolved collision, the first-seen
 * file wins (callers should block generation until collisions resolve).
 */
export function mergeTables(
  tables: ExcelTable[],
  resolution: CollisionResolution = {}
): MergedDataset {
  const collisions = detectCollisions(tables)
  const collisionColumns = new Set(collisions.map((c) => c.column))

  // Determine the ordered unique column list and the source that owns each column.
  const columns: ExcelColumn[] = []
  const owner = new Map<string, string>() // column -> owning file name
  for (const table of tables) {
    for (const header of table.headers) {
      if (!owner.has(header)) {
        const source = collisionColumns.has(header)
          ? resolution[header] ?? table.name
          : table.name
        owner.set(header, source)
        columns.push({ name: header, source })
      } else if (collisionColumns.has(header) && resolution[header] === table.name) {
        // Resolution points at this file; update the owner.
        owner.set(header, table.name)
        const col = columns.find((c) => c.name === header)
        if (col) col.source = table.name
      }
    }
  }

  // Build rows by index across all files, honoring column ownership.
  const rowCount = tables.reduce((max, t) => Math.max(max, t.rows.length), 0)
  const rows: Record<string, string>[] = []
  for (let i = 0; i < rowCount; i++) {
    const rowObj: Record<string, string> = {}
    for (const col of columns) {
      const ownerTable = tables.find((t) => t.name === col.source)
      const value = ownerTable?.rows[i]?.[col.name] ?? ''
      rowObj[col.name] = value
    }
    rows.push(rowObj)
  }

  return { columns, rows, collisions }
}
