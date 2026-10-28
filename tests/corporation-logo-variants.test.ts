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
  intimationTemplateFileName,
  civilTenderTemplateFileName,
  qcIntimationCompletionTemplateFileName,
  noticeTemplateSuffix,
  NOTICE_TEMPLATE_VARIANTS,
  WORK_ORDER_TEMPLATE_VARIANTS
} from '../core/workOrderTemplateVariants'
import { findPlaceholdersInAllParts } from '../core/docx-edit'

const res = (f: string) => path.resolve(__dirname, '../resources', f)
// The retired originals (Notice type 1 etc.) are kept in legacy-templates/; the Notice logo variants are checked against their body.
const legacy = (f: string) => path.resolve(__dirname, '../legacy-templates', f)
const BASES = ['work-order', 'zonal-work-order', 'file-backer', 'intimation', 'zonal-concluding-agreement', 'zonal-memo-ee']

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
      expect(fs.existsSync(res(intimationTemplateFileName(v)))).toBe(true)
      expect(zonalTemplateVariant(v)).toBe(v)
    }
    expect(zonalTemplateVariant('kompally')).toBeUndefined()
    // Retired ids (types 1-2) resolve to a bundled template instead of a missing file.
    for (const v of ['default', 'kompally', 'header2', undefined]) {
      expect(fs.existsSync(res(workOrderTemplateFileName(v)))).toBe(true)
      expect(fs.existsSync(res(fileBackerTemplateFileName(v)))).toBe(true)
      expect(fs.existsSync(res(intimationTemplateFileName(v)))).toBe(true)
      expect(fs.existsSync(res(civilTenderTemplateFileName(v)))).toBe(true)
      expect(fs.existsSync(res(qcIntimationCompletionTemplateFileName(v)))).toBe(true)
    }
  })

  it('offers each office only its own corporation header', () => {
    const ids = (c?: string) => variantsForCorporation(WORK_ORDER_TEMPLATE_VARIANTS, c).map((v) => v.id)
    expect(ids('CMC')).toEqual(['cmc'])
    expect(ids('GHMC')).toEqual(['ghmc'])
    expect(ids('MMC')).toEqual(['mmc'])
    expect(ids(undefined)).toEqual([])
    expect(variantForCorporation('mmc', 'CMC')).toBeUndefined()
    expect(variantForCorporation('cmc', 'CMC')).toBe('cmc')
    expect(variantForCorporation('header2', 'GHMC')).toBe('header2')
    // A saved pick for a retired template (Work Order type 1) falls back to the office's own corporation header.
    expect(variantForCorporation('default', 'CMC', WORK_ORDER_TEMPLATE_VARIANTS)).toBe('cmc')
    expect(variantForCorporation('kompally', undefined, WORK_ORDER_TEMPLATE_VARIANTS)).toBeUndefined()
  })

  it('defaults MMC, CMC and GHMC offices to their own header', () => {
    expect(corporationDefaultVariant('MMC', 'default')).toBe('mmc')
    expect(corporationDefaultVariant('GHMC', 'default')).toBe('ghmc')
    expect(corporationDefaultVariant('CMC', 'default')).toBe('cmc')
    expect(corporationDefaultVariant('XYZ', 'default')).toBe('default')
    expect(corporationDefaultVariant(undefined, 'default')).toBe('default')
  })
})

describe('Notice logo-header templates', () => {
  for (const corp of ['cmc', 'ghmc', 'mmc']) {
    for (const base of ['issue-notice-template', 'issue-notice-ee-template']) {
      it(`${base}-${corp}: logo header over the original body, header image at true aspect`, () => {
        const orig = findPlaceholdersInAllParts(fs.readFileSync(legacy(`${base}.docx`)))
        const buf = fs.readFileSync(res(`${base}-${corp}.docx`))
        const got = findPlaceholdersInAllParts(buf)
        // Only the old text letterhead's corporation-name placeholder goes away.
        expect(got.filter((l) => !orig.includes(l))).toEqual([])
        expect(orig.filter((l) => !got.includes(l)).every((l) => l === 'Corp Full Caps')).toBe(true)
        expect(got).toContain('Agency Name')
        const zip = new PizZip(buf)
        expect(zip.file('word/_rels/document.xml.rels')!.asText()).toContain(`media/${corp}_header.png`)
        expect(zip.file('[Content_Types].xml')!.asText()).toContain('Extension="png" ContentType="image/png"')
        const xml = zip.file('word/document.xml')!.asText()
        // Stretched edge to edge across the body text (page width minus both margins), no table indent or cell padding.
        const pg = /<w:pgSz w:w="(\d+)"/.exec(xml)![1]
        const mar = /<w:pgMar [^>]*w:right="(\d+)"[^>]*w:left="(\d+)"/.exec(xml)!
        const textW = Number(pg) - Number(mar[1]) - Number(mar[2])
        expect(Number(/<wp:extent cx="(\d+)"/.exec(xml)![1])).toBe(textW * 635)
        expect(xml).toContain(`<w:tblW w:w="${textW}" w:type="dxa"/>`)
        expect(xml).toContain('<w:tblInd w:w="0" w:type="dxa"/>')
      })
    }
  }

  it('maps notice variant ids to bundled files that exist and offers each office only its own logo', () => {
    for (const v of ['cmc', 'ghmc', 'mmc', '2', 'default', 'bogus']) {
      expect(fs.existsSync(res(`issue-notice-template${noticeTemplateSuffix(v)}.docx`))).toBe(true)
      expect(fs.existsSync(res(`issue-notice-ee-template${noticeTemplateSuffix(v)}.docx`))).toBe(true)
    }
    expect(variantsForCorporation(NOTICE_TEMPLATE_VARIANTS, 'GHMC').map((v) => v.id)).toEqual(['ghmc'])
    expect(noticeTemplateSuffix('bogus')).toBe('-cmc')
  })
})
