/**
 * Regenerates core/data/materialQuantityRag.json from the SE office Data Sheet.
 *   npx vite-node scripts/buildMaterialRag.ts "<path to Data Sheet .xlsx>"
 * Quantities only change when the departmental data does, so this is run
 * rarely — not every year when rates are revised.
 */
import fs from 'node:fs'
import path from 'node:path'
import { readAllSheetGrids } from '../core/excel'
import { buildRateIndex } from '../core/rateDatabase'
import { buildMaterialRag } from '../core/materialQuantityRag'

const file = process.argv[2]
if (!file) throw new Error('Usage: vite-node scripts/buildMaterialRag.ts <Data Sheet .xlsx>')
const sheets = readAllSheetGrids(file)
const index = buildRateIndex(sheets.map((s) => ({ name: s.sheetName, grid: s.grid })))
const rag = buildMaterialRag(index, path.basename(file).replace(/\.xlsx?$/i, ''))
const out = path.join(__dirname, '..', 'core', 'data', 'materialQuantityRag.json')
fs.writeFileSync(out, JSON.stringify(rag))
console.log(`${sheets.length} sheets, ${index.length} rate entries -> ${rag.entries.length} material items`)
for (const [k, v] of Object.entries(rag.byMaterial)) console.log(`  ${k}: ${v.length} items`)
console.log(`wrote ${out} (${(fs.statSync(out).size / 1024).toFixed(0)} KB)`)
