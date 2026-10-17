import PizZip from 'pizzip'
import { fillableParts, replaceTextInPart } from './docx-edit'

/**
 * The bundled templates were authored for Cyberabad Municipal Corporation, so
 * their letterheads still carry its name — English, Telugu, WordArt watermark
 * and CMC logo — as fixed text rather than {{placeholders}}. When the chosen
 * corporation is GHMC or MMC, applyCorporationBranding rewrites those fixed
 * pieces in the finished document so no header names the wrong corporation.
 */
export interface CorporationBranding {
  abbr: string
  english: string
  telugu: string
  /** File name under resources/corp-logos/. */
  logoFile: string
}

export const CORPORATION_BRANDING: CorporationBranding[] = [
  { abbr: 'CMC', english: 'Cyberabad Municipal Corporation', telugu: 'సైబరాబాద్ మున్సిపల్ కార్పొరేషన్', logoFile: 'corp-logos/cmc.png' },
  // Telugu names as printed on https://www.ghmc.gov.in/ and https://mmc.telangana.gov.in/
  { abbr: 'GHMC', english: 'Greater Hyderabad Municipal Corporation', telugu: 'హైదరాబాద్ మహానగర పాలక సంస్థ', logoFile: 'corp-logos/ghmc.png' },
  { abbr: 'MMC', english: 'Malkajgiri Municipal Corporation', telugu: 'మల్కాజిగిరి మున్సిపల్ కార్పొరేషన్', logoFile: 'corp-logos/mmc.png' }
]

const CMC = CORPORATION_BRANDING[0]

/** The branding for the corporation a filled row names (full name or abbreviation), if recognised. */
export function brandingFromRow(row: Record<string, string>): CorporationBranding | undefined {
  for (const key of ['Corp Full Caps', 'Corp Full', 'Corporation Full Name']) {
    const v = (row[key] ?? '').trim().toLowerCase()
    if (v) {
      const hit = CORPORATION_BRANDING.find((b) => b.english.toLowerCase() === v)
      if (hit) return hit
    }
  }
  const abbr = (row['Corporation'] || row['Corp'] || '').trim().toUpperCase()
  return CORPORATION_BRANDING.find((b) => b.abbr === abbr)
}

function pngSize(buf: Buffer): { w: number; h: number } | null {
  if (buf.length < 24 || buf.readUInt32BE(0) !== 0x89504e47) return null
  return { w: buf.readUInt32BE(16), h: buf.readUInt32BE(20) }
}

/** Swap the CMC logo image (matched byte-for-byte) for the target's, keeping its width and re-deriving its height. */
function swapLogo(buffer: Buffer, cmcLogo: Buffer, targetLogo: Buffer): Buffer {
  const zip = new PizZip(buffer)
  const media = Object.keys(zip.files).filter((f) => /^word\/media\/[^/]+\.png$/i.test(f))
  const hit = media.filter((f) => Buffer.from(zip.file(f)!.asUint8Array()).equals(cmcLogo))
  if (hit.length === 0) return buffer
  const size = pngSize(targetLogo)
  const cmcSize = pngSize(cmcLogo)
  for (const f of hit) {
    zip.file(f, targetLogo)
    // Re-proportion every drawing that embeds this image (rels may live in the body or a header).
    const name = f.slice('word/'.length)
    for (const rels of Object.keys(zip.files).filter((r) => /^word\/_rels\/(document|header\d*|footer\d*)\.xml\.rels$/.test(r))) {
      const relsXml = zip.file(rels)!.asText()
      const id = new RegExp(`<Relationship [^>]*Id="([^"]+)"[^>]*Target="${name.replace(/[.]/g, '\\.')}"|<Relationship [^>]*Target="${name.replace(/[.]/g, '\\.')}"[^>]*Id="([^"]+)"`).exec(relsXml)
      const rid = id?.[1] ?? id?.[2]
      if (!rid || !size || !cmcSize) continue
      const partName = 'word/' + rels.slice('word/_rels/'.length).replace(/\.rels$/, '')
      let xml = zip.file(partName)!.asText()
      xml = xml.replace(/<w:drawing>[\s\S]*?<\/w:drawing>/g, (drawing) => {
        if (!drawing.includes(`r:embed="${rid}"`)) return drawing
        const ratio = size.h / size.w
        const fix = (d: string, tag: string) =>
          d.replace(new RegExp(`<${tag} cx="(\\d+)" cy="\\d+"`, 'g'), (_m, cx) => `<${tag} cx="${cx}" cy="${Math.round(Number(cx) * ratio)}"`)
        return fix(fix(drawing, 'wp:extent'), 'a:ext')
      })
      zip.file(partName, xml)
    }
  }
  return zip.generate({ type: 'nodebuffer', compression: 'DEFLATE' })
}

/**
 * Rewrite the fixed Cyberabad letterhead pieces for `target`. A no-op for CMC
 * (the templates' own corporation) and for documents that carry none of it.
 */
export function applyCorporationBranding(
  buffer: Buffer,
  target: CorporationBranding,
  loadLogo?: (file: string) => Buffer | undefined
): Buffer {
  if (target.abbr === CMC.abbr) return buffer
  let buf = buffer
  const textRules: Array<[string | RegExp, string]> = [
    [CMC.english.toUpperCase(), target.english.toUpperCase()],
    [CMC.english, target.english],
    [CMC.telugu, target.telugu],
    // "{{Circle}} circle-{{CNO}},CMC" / "{{Zone}} Zone, CMC" sign-offs
    [/(circle-[^,\n]*,\s*)CMC\b/gi, `$1${target.abbr}`],
    [/(\bZone,?\s+)CMC\b/g, `$1${target.abbr}`]
  ]
  for (const part of fillableParts(buf)) buf = replaceTextInPart(buf, part, textRules)

  // WordArt watermark (<v:textpath string="CYBERABAD MUNICIPAL CORPORATION">) isn't paragraph text.
  const zip = new PizZip(buf)
  let touched = false
  const wordArt = `string="${CMC.english.toUpperCase()}"`
  for (const part of Object.keys(zip.files).filter((f) => /^word\/[^/]+\.xml$/.test(f))) {
    const xml = zip.file(part)!.asText()
    if (!xml.includes(wordArt)) continue
    zip.file(part, xml.split(wordArt).join(`string="${target.english.toUpperCase()}"`))
    touched = true
  }
  if (touched) buf = zip.generate({ type: 'nodebuffer', compression: 'DEFLATE' })

  const cmcLogo = loadLogo?.(CMC.logoFile)
  const targetLogo = loadLogo?.(target.logoFile)
  if (cmcLogo && targetLogo) buf = swapLogo(buf, cmcLogo, targetLogo)
  return buf
}
