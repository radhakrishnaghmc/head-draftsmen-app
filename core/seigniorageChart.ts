/** Seigniorage chart (from "Seignarage 2025-26 -1.xlsx", Sheet1): per work item, the factor of each material per unit of work. */
export type SeigMaterial = 'Metal' | 'M.sand' | 'Sand'

export interface SeigSub {
  material: SeigMaterial
  /** Sub-item label when it differs from the material (e.g. plaster base/top coats). */
  label?: string
  factor: number
}

export interface SeigItem {
  sl: number
  name: string
  subs: SeigSub[]
}

export const SEIG_RATES: Record<SeigMaterial, number> = { Metal: 117, 'M.sand': 117, Sand: 40 }

/** Permit fee is charged on this share of the total, with Sand excluded; SMET and DMFT on the full total. */
export const SEIG_PERMIT_SHARE = 0.8
export const SEIG_SMET_SHARE = 0.02
export const SEIG_DMFT_SHARE = 0.3

const m = (material: SeigMaterial, factor: number, label?: string): SeigSub => ({ material, factor, label })
const sandPair = (f: number, tag = ''): SeigSub[] => [
  m('M.sand', f, tag ? `M.sand (${tag})` : undefined),
  m('Sand', f, tag ? `Sand (${tag})` : undefined)
]

export const SEIG_CHART: SeigItem[] = [
  { sl: 1, name: 'Plain Cement concrete', subs: [m('Metal', 0.9), ...sandPair(0.225)] },
  { sl: 2, name: 'RMC, RCC', subs: [m('Metal', 0.9), ...sandPair(0.225)] },
  { sl: 3, name: 'RR Masonry', subs: [m('Metal', 1.1), ...sandPair(0.17)] },
  { sl: 4, name: 'Brick Masonry in CM (1:6) with 2nd class Bricks traditional size 23 x 11 x 7 cms', subs: sandPair(0.1) },
  { sl: 5, name: 'Brick Masonry in CM (1:3) with Modular Bricks 19 x 9 x 9 cms 2nd class', subs: sandPair(0.11) },
  { sl: 6, name: 'CRS Masonry in CM (1:6) 1st Sort', subs: [m('Metal', 1.1), ...sandPair(0.14)] },
  { sl: 7, name: 'CRS Masonry in CM (1:6) 2nd Sort', subs: [m('Metal', 1.1), ...sandPair(0.16)] },
  { sl: 8, name: 'Flush pointing', subs: sandPair(0.02) },
  { sl: 9, name: 'Raised pointing', subs: sandPair(0.03) },
  { sl: 10, name: 'Plastering with CM, 12 mm thick', subs: sandPair(0.012 * 0.0075) },
  { sl: 11, name: 'Plastering with CM, 15 mm thick', subs: sandPair(0.015 * 0.0095) },
  { sl: 12, name: 'Plastering with CM, 20 mm thick', subs: sandPair(0.02 * 0.011) },
  {
    sl: 13,
    name: 'Plastering with CM 2 coats, 12 mm thick (base coat 8 mm CM 1:5, top coat 4 mm CM 1:3)',
    subs: [...sandPair(0.008 * 0.006, 'base coat'), ...sandPair(0.002 * 0.004, 'top coat')]
  },
  {
    sl: 14,
    name: 'Plastering with CM 2 coats, 20 mm thick (base coat 16 mm CM 1:6, top coat 4 mm CM 1:4)',
    subs: [...sandPair(0.016 * 0.009, 'base coat'), ...sandPair(0.002 * 0.004, 'top coat')]
  },
  { sl: 15, name: 'Damp proof course 40/50 mm thick with cement concrete (1:2:4)', subs: [m('Metal', 0.9), ...sandPair(0.23)] },
  { sl: 16, name: 'Flooring with Rough Cuddapah/Shabad stones', subs: sandPair(0.07) },
  { sl: 17, name: 'Flooring with polished marble slab', subs: sandPair(0.07) },
  { sl: 18, name: 'Flooring with chequered terrazzo tiles of 22 mm thick', subs: sandPair(0.06) },
  { sl: 19, name: 'Flooring with vitrified tiles', subs: sandPair(0.07) },
  { sl: 20, name: 'Flooring with ceramic tiles', subs: sandPair(0.06) },
  { sl: 21, name: 'Flooring with high polished granite', subs: sandPair(0.1) },
  { sl: 22, name: 'Flooring with marble tiles', subs: sandPair(0.1) },
  { sl: 23, name: 'Flooring with granite stone tiles', subs: sandPair(0.1) },
  { sl: 24, name: 'Flooring with decorated glazed tiles', subs: sandPair(0.06) },
  { sl: 25, name: 'Providing skirting to internal walls', subs: sandPair(0.06) },
  { sl: 26, name: 'Raising/Additional Manhole', subs: [m('Metal', 0.0417), m('Sand', 0.0417)] },
  { sl: 27, name: 'Wet Mix Macadam', subs: [m('Metal', 1.32)] },
  { sl: 28, name: 'Aggregates 6 mm', subs: [m('Metal', 1)] },
  { sl: 29, name: 'Stone Dust 3 mm', subs: [m('Metal', 1)] }
]
