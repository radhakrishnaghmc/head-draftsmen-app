import { describe, expect, it } from 'vitest'
import { inferWorkLocation } from '../core/workLocationExtractor'

// Real work names from a live Monitoring Format "list of works" sheet
// (Quthbullapur Zone, Oct 2026) — each pinned to the location a human
// reading the name would pick out, so a future regex tweak that silently
// regresses one of these gets caught.
describe('inferWorkLocation', () => {
  it('returns null for a blank name', () => {
    expect(inferWorkLocation('')).toBeNull()
    expect(inferWorkLocation('   ')).toBeNull()
  })

  it('returns null when the name has no from/at/in/near-style clause', () => {
    expect(inferWorkLocation('Annual desilting')).toBeNull()
  })

  it('picks the colony/nagar named right before the ward clause', () => {
    expect(
      inferWorkLocation(
        'Laying of CC Road at road no 5 @ Hno :- 6-292/1 at Venkateshwaranagar in ward no : 131, Padma, Quthbullapur Circle - 25, GHMC'
      )
    ).toBe('Venkateshwaranagar')
    expect(
      inferWorkLocation(
        'Removal of water stagnation by Laying of CC Road from Plot No 128 to 124 and Plot No.113 to Hanuman Temple in MN Reddy Nagar colony ward no: 132, Jeedimetla'
      )
    ).toBe('MN Reddy Nagar colony')
    expect(
      inferWorkLocation('Providing Sand in Over Lap Park in Ward no 6 in NMC under MGF 2025-26')
    ).toBe('Over Lap Park')
  })

  it('keeps a road number as a valid location (unlike plot/house/survey numbers)', () => {
    expect(
      inferWorkLocation('Laying of UGD and CC Road in road no:10 And 10C in ward no:18 in Nizampet Municipal Corporation under MGF 2025-26')
    ).toBe('road no:10 And 10C')
  })

  it('strips a plot/house/survey-number prefix fused onto the actual place name', () => {
    expect(
      inferWorkLocation('Laying of CC road from 7-10 to plot no.154 Sri Venkateswara residency in ward No.23 in NMC under MGF')
    ).toBe('Sri Venkateswara residency')
  })

  it('drops corporation/zone/scheme abbreviations left over with nothing more specific named', () => {
    expect(
      inferWorkLocation('Installation of Highmast Lights in NMC under MGF 2024-25')
    ).toBeNull()
  })
})
