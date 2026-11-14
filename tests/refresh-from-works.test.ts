import { expect, it } from 'vitest'
import { refreshAgainstWorks, refreshLetter } from '../core/refreshFromWorks'
import type { StoredLetter } from '../core/types'

const table = {
  id: 't', name: 't', path: '',
  headers: ['Name of the work', 'Amount of estimate', 'Tender ID', 'Tender Notice No', 'Name of the Agency'],
  rows: [{ 'Name of the work': 'Laying of CC road in Ward 5', 'Amount of estimate': '12.5', 'Tender ID': '717574', 'Tender Notice No': '12/DB/EE', 'Name of the Agency': 'M V S' }]
}
const letter = (workName: string): StoredLetter => ({
  id: 'a', kind: 'deviation', nitNo: '', nitDate: 'keep', agency: '', workName, typeOfWork: '', estimateRupees: null, variant: 'default',
  values: { meta: JSON.stringify({ estimateAmountLakhs: 0 }) }, fileName: 'f', createdDate: '2026-01-01'
})

it('refreshes tender id, estimate and agency by work name, leaving blank cells alone', () => {
  const r = refreshAgainstWorks([letter('Laying of CC road in Ward 5'), letter('Unknown work')], table, refreshLetter)
  expect(r.found).toBe(1)
  expect(r.items[0]).toMatchObject({ nitNo: '12/DB/EE', nitDate: 'keep', agency: 'M V S', estimateRupees: 1250000 })
  expect(r.items[0].values.tenderId).toBe('717574')
  expect(JSON.parse(r.items[0].values.meta).estimateAmountLakhs).toBe(12.5)
  expect(r.items[1].nitNo).toBe('')
})
