import { describe, it, expect } from 'vitest'
import fs from 'fs'
import path from 'path'
import PizZip from 'pizzip'
import {
  variantsForCorporation,
  variantForCorporation,
  corporationDefaultVariant,
  zonalTemplateVariant,
  workOrderTemplateFileName,
  fileBackerTemplateFileName,
  WORK_ORDER_TEMPLATE_VARIANTS
} from '../core/workOrderTemplateVariants'
import { findPlaceholdersInAllParts } from '../core/docx-edit'

const res = (f: string) => path.resolve(__dirname, '../resources', f)
const BASES = ['work-order', 'zonal-work-order', 'file-backer', 'zonal-concluding-agreement', 'zonal-memo-ee']

describe('corporation logo-header templates (CMC / GHMC)', () => {
  for (const corp of ['cmc', 'ghmc']) {
    it(`${corp}: same body placeholders as MMC, own undistorted header image`, () => {
      for (const base of BASES) {
        const mmc = fs.readFileSync(res(`${base}-template-mmc.docx`))
        const buf = fs.readFileSync(res(`${base}-template-${corp}.docx`))
        expect(findPlaceholdersInAllParts(buf).sort()).toEqual(findPlaceholdersInAllParts(mmc).sort())
        const zip = new PizZip(buf)
        expect(zip.file('word/media/mmc_header.png')).toBeNull()
        const xml = zip.file('word/document.xml')!.asText()
        expect(zip.file('word/_rels/document.xml.rels')!.asText()).toContain(`media/${corp}_header.png`)
        const drawing = [...xml.matchAll(/<w:drawing>[\s\S]*?<\/w:drawing>/g)].map((m) => m[0]).find((d) => d.includes(`rId${corp === 'cmc' ? 'Cmc' : 'Ghmc'}4`))!
        // Same drawing size as the MMC header, so it spans the body width the same way.
        const extent = (x: string) => /<wp:extent cx="(\d+)" cy="(\d+)"/.exec(x)![0]
        const mmcXml = new PizZip(mmc).file('word/document.xml')!.asText()
        expect(extent(drawing)).toBe(extent(mmcXml))
      }
    })
  }

  it('maps variant ids to bundled files that exist', () => {
    for (const v of ['cmc', 'ghmc']) {
      expect(fs.existsSync(res(workOrderTemplateFileName(v)))).toBe(true)
      expect(fs.existsSync(res(fileBackerTemplateFileName(v)))).toBe(true)
      expect(zonalTemplateVariant(v)).toBe(v)
    }
    expect(zonalTemplateVariant('kompally')).toBeUndefined()
  })

  it('offers each office only its own corporation header', () => {
    const ids = (c?: string) => variantsForCorporation(WORK_ORDER_TEMPLATE_VARIANTS, c).map((v) => v.id)
    expect(ids('CMC')).toEqual(['default', 'kompally', 'header2', 'cmc'])
    expect(ids('GHMC')).toEqual(['default', 'kompally', 'header2', 'ghmc'])
    expect(ids('MMC')).toEqual(['default', 'kompally', 'header2', 'mmc'])
    expect(ids(undefined)).toEqual(['default', 'kompally', 'header2'])
    expect(variantForCorporation('mmc', 'CMC')).toBeUndefined()
    expect(variantForCorporation('cmc', 'CMC')).toBe('cmc')
    expect(variantForCorporation('header2', 'GHMC')).toBe('header2')
  })

  it('defaults MMC, CMC and GHMC offices to their own header', () => {
    expect(corporationDefaultVariant('MMC', 'default')).toBe('mmc')
    expect(corporationDefaultVariant('GHMC', 'default')).toBe('ghmc')
    expect(corporationDefaultVariant('CMC', 'default')).toBe('cmc')
    expect(corporationDefaultVariant('XYZ', 'default')).toBe('default')
    expect(corporationDefaultVariant(undefined, 'default')).toBe('default')
  })
})
