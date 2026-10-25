import { describe, it, expect } from 'vitest'
import fs from 'node:fs'
import path from 'node:path'
import { applyCorporationBranding, brandingFromRow, CORPORATION_BRANDING } from '../core/corporationBranding'
import { extractAllText } from '../core/docx-edit'
import PizZip from 'pizzip'

// Retired templates (Work Order type 1 etc.) live in legacy-templates/; their branding retarget still has to work if one is restored.
const res = (f: string) => {
  const bundled = path.join(__dirname, '..', 'resources', f)
  return fs.existsSync(bundled) ? bundled : path.join(__dirname, '..', 'legacy-templates', f)
}
const logo = (f: string) => fs.readFileSync(res(f))
const GHMC = CORPORATION_BRANDING[1]
const MMC = CORPORATION_BRANDING[2]

describe('corporation branding', () => {
  it('resolves the corporation from a filled row', () => {
    expect(brandingFromRow({ 'Corp Full Caps': 'MALKAJGIRI MUNICIPAL CORPORATION' })?.abbr).toBe('MMC')
    expect(brandingFromRow({ Corporation: 'ghmc' })?.abbr).toBe('GHMC')
    expect(brandingFromRow({})).toBeUndefined()
  })

  for (const file of ['work-order-template-header2', 'file-backer-template-header2', 'civil-tender-template-header2', 'civil-tender-template', 'intimation-template-2', 'issue-notice-template-2', 'zonal-work-order-template', 'zonal-work-order-template-header2']) {
    for (const target of [GHMC, MMC]) {
      it(`${file} carries no Cyberabad text after retargeting to ${target.abbr}`, () => {
        const out = applyCorporationBranding(fs.readFileSync(res(`${file}.docx`)), target, logo)
        const z = new PizZip(out)
        const all = Object.keys(z.files).filter((f) => /^word\/(document|header\d*|footer\d*)\.xml$/.test(f)).map((f) => z.file(f)!.asText()).join('')
        expect(all).not.toMatch(/CYBERABAD MUNICIPAL CORPORATION/)
        expect(all).not.toContain('సైబరాబాద్')
        expect(extractAllText(out)).not.toMatch(/cyberabad/i)
      })
    }
  }

  it('writes the Telugu name and swaps the logo', () => {
    const out = applyCorporationBranding(fs.readFileSync(res('file-backer-template-header2.docx')), MMC, logo)
    expect(extractAllText(out)).toContain('మల్కాజిగిరి మున్సిపల్ కార్పొరేషన్')
    expect(extractAllText(out)).toContain('MALKAJGIRI MUNICIPAL CORPORATION')
    const z = new PizZip(out)
    expect(Buffer.from(z.file('word/media/image2.png')!.asUint8Array()).equals(logo('corp-logos/mmc.png'))).toBe(true)
  })

  it('is a no-op for CMC', () => {
    const buf = fs.readFileSync(res('work-order-template-header2.docx'))
    expect(applyCorporationBranding(buf, CORPORATION_BRANDING[0], logo)).toBe(buf)
  })

  it('retargets the agreement sign-off', () => {
    const out = applyCorporationBranding(fs.readFileSync(res('agreement-template.docx')), GHMC, logo)
    expect(extractAllText(out)).toMatch(/circle-\{\{CNO\}\},GHMC/)
  })
})

describe('SE office documents', () => {
  it('retargets SE Work Order / Concluding Agreement / Contract Deed letterheads', () => {
    for (const file of ['zonal-concluding-agreement-template', 'zonal-memo-ee-template', 'contract-deed-template', 'se-agreement-bond-template', 'se-agreement-note-template']) {
      const out = applyCorporationBranding(fs.readFileSync(res(`${file}.docx`)), MMC, logo)
      expect(extractAllText(out)).not.toMatch(/cyberabad/i)
    }
  })

  it('SE evaluation notes name the chosen corporation instead of CMC', async () => {
    const { buildBidEvaluationHtml, buildAgencyApprovalHtml } = await import('../core/seEvaluationNotes')
    const base = { zoneAbbr: 'QBZ', corp: 'MMC', workName: 'W', estimateLakhs: '10', asAuthority: 'Commissioner, MMC', asRefLine: 'x', tsNo: '1', tsDate: 'd', financialYear: '2026-27', nitNo: 'n', nitDate: 'd', itemNo: '', techBidOpenDate: 'd', zone: 'Qutbullapur', bidders: [], nonRespCount: 0, bidEvalApprovedDate: 'd' } as never
    for (const html of [buildBidEvaluationHtml(base), buildAgencyApprovalHtml(base)]) {
      expect(html).toContain('MMC')
      expect(html).not.toMatch(/\bCMC\b/)
    }
  })
})

describe('Work Order type 4 (MMC)', () => {
  it('keeps type 3 body placeholders and carries the horizontal MMC logo as the whole header', async () => {
    const { findPlaceholdersInAllParts } = await import('../core/docx-edit')
    for (const [t3, t4] of [['work-order-template-header2', 'work-order-template-mmc'], ['zonal-work-order-template-header2', 'zonal-work-order-template-mmc']]) {
      const a = findPlaceholdersInAllParts(fs.readFileSync(res(`${t3}.docx`))).sort()
      const b = findPlaceholdersInAllParts(fs.readFileSync(res(`${t4}.docx`))).sort()
      expect(b).toEqual(a)
      const zip = new PizZip(applyCorporationBranding(fs.readFileSync(res(`${t4}.docx`)), MMC, logo))
      // The header image is the cropped horizontal MMC logo (wide strip), not the stacked corp-logos/mmc.png.
      const img = Buffer.from(zip.file('word/media/mmc_header.png')!.asUint8Array())
      expect(img.readUInt32BE(16) / img.readUInt32BE(20)).toBeGreaterThan(4)
    }
  })
})

describe('SE Work Order type 3', () => {
  it('has the same placeholders as the default SE Work Order and carries the MMC logo after branding', async () => {
    const { findPlaceholdersInAllParts } = await import('../core/docx-edit')
    const a = findPlaceholdersInAllParts(fs.readFileSync(res('zonal-work-order-template.docx'))).sort()
    const b = findPlaceholdersInAllParts(fs.readFileSync(res('zonal-work-order-template-header2.docx'))).sort()
    expect(b).toEqual(a)
    const out = applyCorporationBranding(fs.readFileSync(res('zonal-work-order-template-header2.docx')), MMC, logo)
    expect(Buffer.from(new PizZip(out).file('word/media/h2_logo.png')!.asUint8Array()).equals(logo('corp-logos/mmc.png'))).toBe(true)
  })

  it('Concluding Agreement and Memo to EE type 3 keep the default placeholders and take the MMC logo', async () => {
    const { findPlaceholdersInAllParts } = await import('../core/docx-edit')
    for (const base of ['zonal-concluding-agreement-template', 'zonal-memo-ee-template']) {
      const a = findPlaceholdersInAllParts(fs.readFileSync(res(`${base}.docx`)))
      const b = findPlaceholdersInAllParts(fs.readFileSync(res(`${base}-header2.docx`)))
      // Same body placeholders; the plain header's "{{Corp Full Caps}}" title is replaced by the logo header.
      expect(new Set(b)).toEqual(new Set(a.filter((p) => p !== 'Corp Full Caps')))
      const out = applyCorporationBranding(fs.readFileSync(res(`${base}-header2.docx`)), MMC, logo)
      expect(Buffer.from(new PizZip(out).file('word/media/h2_logo.png')!.asUint8Array()).equals(logo('corp-logos/mmc.png'))).toBe(true)
    }
  })
})
