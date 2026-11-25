import { it } from 'vitest'
import { writeFileSync } from 'node:fs'
import { buildMeasurementBookWorkbook } from '../core/measurementBookExcel'
import { emptyMbModel, mergeMbPage } from '../core/measurementBook'

it('writes a sample-equivalent workbook', async () => {
  let m = emptyMbModel()
  const pg = (p: number, o: object) => (m = mergeMbPage(m, p, JSON.stringify(o))!)
  const ln = (location: string, no: number, l: number, b: number[], d: number[], written: number) => ({ location, no, l, b, d, written })
  pg(2, { header: { work: 'Laying of balance CC road in road no (4) Srinivas Housing Society in ward no (15) in NMC under MLF', agency: 'Sri K. Mohan Reddy', estimate: '3.00 lakhs' },
    items: [{ description: 'RBR-EECO-(8)(C): Earthwork excavation for roadway in soil', unit: 'Cum', lines: [ln('Srinivas Housing Society road no (4)', 1, 14, [5.35, 5.35], [0.3], 22.47)], totalWritten: 22.47 }] })
  pg(3, { items: [{ description: 'Providing, laying, spreading and compacting graded stone aggregate to wet mix macadam', unit: 'Cum', lines: [ln('Srinivas Housing Society road no (4)', 1, 14, [5.35, 5.35], [0.18, 0.22], 14.98)], totalWritten: 14.98 }] })
  pg(4, { items: [{ description: 'Supply and placing of ready mix concrete M20', unit: 'Cum', lines: [ln('Srinivas Housing Society road no (4)', 1, 14, [5.35, 5.35], [0.15], 11.24)], totalWritten: 11.24 }] })
  pg(5, { items: [{ description: 'Vacuum dewatering of freshly laid cement concrete floor', unit: 'Sqm', lines: [ln('Srinivas Housing Society road no (4)', 1, 14, [5.35, 5.35], [], 74.9)], totalWritten: 74.9 }] })
  pg(6, { items: [{ description: 'Groove cutting of 2-4mm for 1/3rd depth of concrete slab', unit: 'Rmt', lines: [ln('Srinivas Housing Society road no (4)', 5, 173.7, [], [], 868.5), ln('Srinivas Housing Society road no (4)', 174, 5.5, [], [], 955.35)], totalWritten: 1823.85 }] })
  pg(7, { abstract: [
    { description: 'RBR-EECO-(8)(C): Earthwork excavation', qtyWritten: 22.47, unit: 'Cum', rate: 769.3, amountWritten: 17285, chartSl: null },
    { description: 'Providing, laying, spreading and compacting graded stone aggregate to wet mix macadam', qtyWritten: 14.98, unit: 'Cum', rate: 2399.23, amountWritten: 35940, chartSl: 27 },
    { description: 'Supply and placing of ready mix concrete M20', qtyWritten: 11.24, unit: 'Cum', rate: 5755.13, amountWritten: 64659, chartSl: 2 },
    { description: 'Vacuum dewatering', qtyWritten: 74.9, unit: 'Sqm', rate: 75, amountWritten: 5617, chartSl: null },
    { description: 'Groove cutting', qtyWritten: 1823.85, unit: 'Rmt', rate: 68.17, amountWritten: 124332, chartSl: null }] })
  writeFileSync('/private/tmp/claude-501/-Users-radhakrishnapodugu-HEAD-DRAFTSMEN-APP/864c7810-e40d-48a0-b5ec-7ff69c9a6ce1/scratchpad/mine.xlsx', await buildMeasurementBookWorkbook(m, 'N-271-2025-26'))
})
