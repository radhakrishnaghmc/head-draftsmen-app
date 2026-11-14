import { expect, it } from 'vitest'
import { stripItemNoTag } from '../core/tenderAgents/shared'

it('removes the item-number tag in every spelling the portal prints', () => {
  for (const tag of ['(Item No.3)', '(Item No,3)', '(item no: 03)', 'Item No.3', 'item no 3'])
    expect(stripItemNoTag(`Laying of CC road in Nizampet ${tag}`)).toBe('Laying of CC road in Nizampet')
  expect(stripItemNoTag('Supply of safety items such as testers')).toBe('Supply of safety items such as testers')
})

it('removes a wrapped-line "No.01)" prefix and an unbracketed completion period', () => {
  expect(
    stripItemNoTag('No.01) Maintenance of BT Roads in Ward No 273 in Nizampet Circle -58 (Rs.5.00 lakhs) Completion period: 1month')
  ).toBe('Maintenance of BT Roads in Ward No 273 in Nizampet Circle -58 (Rs.5.00 lakhs)')
})
