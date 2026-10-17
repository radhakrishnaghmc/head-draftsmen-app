/**
 * Every bundled Work Order template variant an office can choose between, in
 * Settings' Document Templates section — different circles' own offices word
 * (and lay out) their Work Order differently, beyond what filling the same
 * placeholders into one shared template can express (see e.g. Kompally
 * Circle-56's real Work Order: a numbered Ref block citing NIT No/Tender ID/
 * Intimation Letter Date, its own "EXECUTIVE ENGINEER" sign-off block, a different Work
 * Order numbering order — the app's original template has none of that).
 *
 * No user uploads: each variant is a real circle's actual template, gathered
 * and added here as a new bundled resource file, not something typed in at
 * runtime — see workOrderTemplateFileName below for the naming convention a
 * new variant follows.
 */
export interface TemplateVariantOption {
  id: string
  label: string
}

export const MMC_WORK_ORDER_TEMPLATE_VARIANT = 'mmc'
export const CMC_WORK_ORDER_TEMPLATE_VARIANT = 'cmc'
export const GHMC_WORK_ORDER_TEMPLATE_VARIANT = 'ghmc'

/** Corporation-specific logo-header variants, keyed by the office's corporation abbreviation. */
const CORPORATION_VARIANT_IDS: Record<string, string> = {
  MMC: MMC_WORK_ORDER_TEMPLATE_VARIANT,
  CMC: CMC_WORK_ORDER_TEMPLATE_VARIANT,
  GHMC: GHMC_WORK_ORDER_TEMPLATE_VARIANT
}
const CORPORATION_VARIANT_ID_SET = new Set(Object.values(CORPORATION_VARIANT_IDS))

/** Drops the other corporations' logo-header variants from a picker list. */
export function variantsForCorporation(variants: TemplateVariantOption[], corporation: string | undefined): TemplateVariantOption[] {
  const own = corporation ? CORPORATION_VARIANT_IDS[corporation] : undefined
  return variants.filter((v) => !CORPORATION_VARIANT_ID_SET.has(v.id) || v.id === own)
}

/** A saved variant pick, or undefined when it's another corporation's logo header (a stale choice). */
export function variantForCorporation(variantId: string | undefined, corporation: string | undefined): string | undefined {
  if (!variantId || !CORPORATION_VARIANT_ID_SET.has(variantId)) return variantId
  return variantId === (corporation ? CORPORATION_VARIANT_IDS[corporation] : undefined) ? variantId : undefined
}

/**
 * The variant an office gets with no saved pick: MMC, CMC and GHMC offices default
 * to their own logo header, everything else keeps `fallback`.
 */
export function corporationDefaultVariant<T extends string | undefined>(corporation: string | undefined, fallback: T): string | T {
  return corporation && CORPORATION_VARIANT_IDS[corporation] ? CORPORATION_VARIANT_IDS[corporation] : fallback
}

/** Resolves a zone-level (SE) template variant id to its filename suffix, or undefined for the default letterhead. */
export function zonalTemplateVariant(variantId: string | undefined): string | undefined {
  return variantId === 'header2' || (variantId && CORPORATION_VARIANT_ID_SET.has(variantId)) ? variantId : undefined
}

export const WORK_ORDER_TEMPLATE_VARIANTS: TemplateVariantOption[] = [
  { id: 'default', label: 'Work Order type 1' },
  { id: 'kompally', label: 'Work Order type 2' },
  { id: 'header2', label: 'Work Order type 3' },
  // Corporation-only: type 3's layout with that corporation's horizontal logo as the whole header (types 4-6).
  { id: MMC_WORK_ORDER_TEMPLATE_VARIANT, label: 'Work Order type 4 (MMC)' },
  { id: CMC_WORK_ORDER_TEMPLATE_VARIANT, label: 'Work Order type 5 (CMC)' },
  { id: GHMC_WORK_ORDER_TEMPLATE_VARIANT, label: 'Work Order type 6 (GHMC)' }
]

export const DEFAULT_WORK_ORDER_TEMPLATE_VARIANT = 'default'

/**
 * SE (zone-level) Concluding Agreement and Memo to EE: the default letterhead,
 * or the logo header shared with Work Order type 3 ("header2") — body text is
 * identical. Resolved in electron/main.ts.
 */
export const CONCLUDING_AGREEMENT_TEMPLATE_VARIANTS: TemplateVariantOption[] = [
  { id: 'default', label: 'Concluding Agreement type 1' },
  { id: 'header2', label: 'Concluding Agreement type 3' },
  // MMC-only: the horizontal Malkajgiri Municipal Corporation logo header (Work Order type 4).
  { id: MMC_WORK_ORDER_TEMPLATE_VARIANT, label: 'Concluding Agreement type 4 (MMC)' },
  { id: CMC_WORK_ORDER_TEMPLATE_VARIANT, label: 'Concluding Agreement type 5 (CMC)' },
  { id: GHMC_WORK_ORDER_TEMPLATE_VARIANT, label: 'Concluding Agreement type 6 (GHMC)' }
]
export const MEMO_EE_TEMPLATE_VARIANTS: TemplateVariantOption[] = [
  { id: 'default', label: 'Memo to EE type 1' },
  { id: 'header2', label: 'Memo to EE type 3' },
  { id: MMC_WORK_ORDER_TEMPLATE_VARIANT, label: 'Memo to EE type 4 (MMC)' },
  { id: CMC_WORK_ORDER_TEMPLATE_VARIANT, label: 'Memo to EE type 5 (CMC)' },
  { id: GHMC_WORK_ORDER_TEMPLATE_VARIANT, label: 'Memo to EE type 6 (GHMC)' }
]

/**
 * Maps a variant id to its bundled resource filename — "work-order-template
 * .docx" for the original default, "work-order-template-<id>.docx" for every
 * other variant. Main-process-only (electron/main.ts is the only caller that
 * actually needs to load the file); falls back to the default filename for
 * an id that isn't (or is no longer) registered above, so a stale saved
 * selection can never point at a file that doesn't exist.
 */
export function workOrderTemplateFileName(variantId: string | undefined): string {
  const found = WORK_ORDER_TEMPLATE_VARIANTS.find((v) => v.id === variantId)
  if (!found || found.id === DEFAULT_WORK_ORDER_TEMPLATE_VARIANT) return 'work-order-template.docx'
  return `work-order-template-${found.id}.docx`
}

/**
 * Same variant-picker shape, for the Agreement Bond — a real circle's own
 * wording (e.g. Kompally Circle-56's "AGREEMENT - DEED", full contractor
 * legal boilerplate, no ECV/Tender-percentage line at all) can differ from
 * the app's own default "A G R E E M E N T" cover page just as much as a
 * Work Order can.
 */
export const AGREEMENT_TEMPLATE_VARIANTS: TemplateVariantOption[] = [
  { id: 'default', label: 'Agreement Bond type 1' },
  { id: 'kompally', label: 'Agreement Bond type 2' }
]

export const DEFAULT_AGREEMENT_TEMPLATE_VARIANT = 'default'

/** Same naming convention as workOrderTemplateFileName, for agreement-template*.docx. */
export function agreementTemplateFileName(variantId: string | undefined): string {
  const found = AGREEMENT_TEMPLATE_VARIANTS.find((v) => v.id === variantId)
  if (!found || found.id === DEFAULT_AGREEMENT_TEMPLATE_VARIANT) return 'agreement-template.docx'
  return `agreement-template-${found.id}.docx`
}

/**
 * Same variant-picker shape, for the Intimation letter (Give Intimation /
 * Tools' Intimation tool) — e.g. Nizampet Circle-58's own template carries a
 * logo header (Telangana emblem + "Cyberabad Municipal Corporation" title +
 * CMC logo) instead of the app's own plain-text "{{Corp Full Caps}}" header.
 */
export const INTIMATION_TEMPLATE_VARIANTS: TemplateVariantOption[] = [
  { id: 'default', label: 'Intimation template 1' },
  { id: '2', label: 'Intimation template 2' }
]

export const DEFAULT_INTIMATION_TEMPLATE_VARIANT = 'default'

/** Same naming convention as workOrderTemplateFileName, for intimation-template*.docx. */
export function intimationTemplateFileName(variantId: string | undefined): string {
  const found = INTIMATION_TEMPLATE_VARIANTS.find((v) => v.id === variantId)
  if (!found || found.id === DEFAULT_INTIMATION_TEMPLATE_VARIANT) return 'intimation-template.docx'
  return `intimation-template-${found.id}.docx`
}

/**
 * Same variant-picker shape, for the File Backer — the file's cover page.
 * "header2" carries Nizampet Circle-58's own logo header (Telangana emblem +
 * "Cyberabad Municipal Corporation" title + CMC logo) in place of the app's
 * plain-text "{{Corp Full Caps}}" heading, matching the Work Order and
 * Intimation "header2"/"2" variants.
 */
export const FILE_BACKER_TEMPLATE_VARIANTS: TemplateVariantOption[] = [
  { id: 'default', label: 'File Backer type 1' },
  { id: 'header2', label: 'File Backer type 2' },
  // MMC-only: the horizontal Malkajgiri Municipal Corporation logo header (Work Order type 4).
  { id: MMC_WORK_ORDER_TEMPLATE_VARIANT, label: 'File Backer type 3 (MMC)' },
  { id: CMC_WORK_ORDER_TEMPLATE_VARIANT, label: 'File Backer type 4 (CMC)' },
  { id: GHMC_WORK_ORDER_TEMPLATE_VARIANT, label: 'File Backer type 5 (GHMC)' }
]

export const DEFAULT_FILE_BACKER_TEMPLATE_VARIANT = 'default'

/** Same naming convention as workOrderTemplateFileName, for file-backer-template*.docx. */
export function fileBackerTemplateFileName(variantId: string | undefined): string {
  const found = FILE_BACKER_TEMPLATE_VARIANTS.find((v) => v.id === variantId)
  if (!found || found.id === DEFAULT_FILE_BACKER_TEMPLATE_VARIANT) return 'file-backer-template.docx'
  return `file-backer-template-${found.id}.docx`
}

/**
 * Same variant-picker shape, for the full Civil Tender Document — "header2"
 * carries the same Nizampet Circle-58 logo header on both the page-1
 * Forwarding Slip and page-2 Notice Inviting Tenders, in place of the app's
 * default plain-text "{{Corp Full}}" heading.
 */
export const CIVIL_TENDER_TEMPLATE_VARIANTS: TemplateVariantOption[] = [
  { id: 'default', label: 'Tender Document type 1' },
  { id: 'header2', label: 'Tender Document type 2' },
  // CMC-only: the horizontal CMC logo header (Work Order type 5) on the Forwarding Slip and NIT pages.
  { id: CMC_WORK_ORDER_TEMPLATE_VARIANT, label: 'Tender Document type 3 (CMC)' }
]

export const DEFAULT_CIVIL_TENDER_TEMPLATE_VARIANT = 'default'

/** Same naming convention as workOrderTemplateFileName, for civil-tender-template*.docx. */
export function civilTenderTemplateFileName(variantId: string | undefined): string {
  const found = CIVIL_TENDER_TEMPLATE_VARIANTS.find((v) => v.id === variantId)
  if (!found || found.id === DEFAULT_CIVIL_TENDER_TEMPLATE_VARIANT) return 'civil-tender-template.docx'
  return `civil-tender-template-${found.id}.docx`
}
