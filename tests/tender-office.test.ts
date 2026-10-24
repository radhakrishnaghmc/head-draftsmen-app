import { describe, expect, it } from 'vitest'
import { officeKeyFromNit } from '../src/tenderOffice'

describe('officeKeyFromNit', () => {
  it('EE NIT naming the circle -> that circle-level office', () => {
    expect(officeKeyFromNit('29/1/DB/EE/Nizampet Circle-58/CMC/2026-27, Dated: 13.10.2026', 'CMC', 'Quthbullapur')).toBe('CMC|Quthbullapur|Nizampet')
    expect(officeKeyFromNit('27/6/DB/EE/Gajularamaram Circle-57/QBZ/CMC', 'CMC', 'Quthbullapur')).toBe('CMC|Quthbullapur|Gajularamaram')
  })
  it('EE: the work name decides the circle, ahead of the NIT', () => {
    expect(officeKeyFromNit('26/EE-59/QBZ/CMC/2026-27', 'CMC', 'Quthbullapur', 'Laying CC road in Bhandari layout in Nizampet Circle-58, Quthbullapur Zone, CMC')).toBe('CMC|Quthbullapur|Nizampet')
  })
  it('SE: the work name is ignored', () => {
    expect(officeKeyFromNit('28/SE/QBZ/CMC/2026-27', 'CMC', 'Quthbullapur', 'Road in Nizampet Circle-58')).toBe('CMC|Quthbullapur|')
  })
  it('EE NIT carrying only the circle number', () => {
    expect(officeKeyFromNit('NIT No. 26/EE-59/QBZ/CMC/2026-27 Dt:08.10.2026', 'CMC', 'Quthbullapur')).toBe('CMC|Quthbullapur|Dundigal')
  })
  it('SE NIT -> the zone-level office', () => {
    expect(officeKeyFromNit('NIT No. 28/SE/QBZ/CMC/2026-27, Dt. 13.10.2026', 'CMC', 'Quthbullapur')).toBe('CMC|Quthbullapur|')
  })
  it('undecidable NITs give undefined', () => {
    expect(officeKeyFromNit(undefined, 'CMC', 'Quthbullapur')).toBeUndefined()
    expect(officeKeyFromNit('12/Seetharam/2026', 'CMC', 'Quthbullapur')).toBeUndefined()
    expect(officeKeyFromNit('30/EE/QBZ/CMC/2026-27', 'CMC', 'Quthbullapur')).toBeUndefined()
  })
})
