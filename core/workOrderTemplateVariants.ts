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
export function variantForCorporation(
  variantId: string | undefined,
  corporation: string | undefined,
  /** The document's current variants: a saved pick no longer in this list (a retired template) falls back to the corporation's own. */
  validVariants?: TemplateVariantOption[]
): string | undefined {
  if (variantId && validVariants && !validVariants.some((v) => v.id === variantId)) return corporationDefaultVariant(corporation, undefined)
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

// Work Order types 1-3, Intimation / Notice template 1-2, File Backer types 1-2, Tender Document types 1-2, Intimation / Notice / QC Intimation type 1
// are retired: their .docx files live in legacy-templates/ (not bundled with the app). Remaining types keep their numbers.
export const WORK_ORDER_TEMPLATE_VARIANTS: TemplateVariantOption[] = [
  // Corporation-only: type 3's layout with that corporation's horizontal logo as the whole header (types 4-6).
  { id: MMC_WORK_ORDER_TEMPLATE_VARIANT, label: 'Work Order type 4 (MMC)' },
  { id: CMC_WORK_ORDER_TEMPLATE_VARIANT, label: 'Work Order type 5 (CMC)' },
  { id: GHMC_WORK_ORDER_TEMPLATE_VARIANT, label: 'Work Order type 6 (GHMC)' }
]

/** Work Order type 3 is retired (legacy-templates/); an office with no corporation header gets the CMC one. */
export const DEFAULT_WORK_ORDER_TEMPLATE_VARIANT = CMC_WORK_ORDER_TEMPLATE_VARIANT

/**
 * SE (zone-level) Concluding Agreement and Memo to EE: the default letterhead,
 * or the logo header shared with Work Order type 3 ("header2") — body text is
 * identical. Resolved in electron/main.ts.
 */
/** Concluding Agreement types 1 and 3 are retired (legacy-templates/); an office with no corporation header gets the CMC one. */
export const DEFAULT_CONCLUDING_AGREEMENT_TEMPLATE_VARIANT = CMC_WORK_ORDER_TEMPLATE_VARIANT
export const CONCLUDING_AGREEMENT_TEMPLATE_VARIANTS: TemplateVariantOption[] = [
  // MMC-only: the horizontal Malkajgiri Municipal Corporation logo header (Work Order type 4).
  { id: MMC_WORK_ORDER_TEMPLATE_VARIANT, label: 'Concluding Agreement type 4 (MMC)' },
  { id: CMC_WORK_ORDER_TEMPLATE_VARIANT, label: 'Concluding Agreement type 5 (CMC)' },
  { id: GHMC_WORK_ORDER_TEMPLATE_VARIANT, label: 'Concluding Agreement type 6 (GHMC)' }
]
/** Memo to EE type 1 is retired (legacy-templates/); an office with no pick and no corporation header gets type 3. */
export const DEFAULT_MEMO_EE_TEMPLATE_VARIANT = 'header2'
export const MEMO_EE_TEMPLATE_VARIANTS: TemplateVariantOption[] = [
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
  return `work-order-template-${found?.id ?? DEFAULT_WORK_ORDER_TEMPLATE_VARIANT}.docx`
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
  // Corporation-only: the retired template 2's body with that corporation's horizontal logo as the whole header (Work Order types 4-6).
  { id: MMC_WORK_ORDER_TEMPLATE_VARIANT, label: 'Intimation template 3 (MMC)' },
  { id: CMC_WORK_ORDER_TEMPLATE_VARIANT, label: 'Intimation template 4 (CMC)' },
  { id: GHMC_WORK_ORDER_TEMPLATE_VARIANT, label: 'Intimation template 5 (GHMC)' }
]

/** Intimation template 2 is retired (legacy-templates/); an office with no corporation header gets the CMC one. */
export const DEFAULT_INTIMATION_TEMPLATE_VARIANT = CMC_WORK_ORDER_TEMPLATE_VARIANT

/** Same naming convention as workOrderTemplateFileName, for intimation-template*.docx. */
export function intimationTemplateFileName(variantId: string | undefined): string {
  const found = INTIMATION_TEMPLATE_VARIANTS.find((v) => v.id === variantId)
  return `intimation-template-${found?.id ?? DEFAULT_INTIMATION_TEMPLATE_VARIANT}.docx`
}

/**
 * Same variant-picker shape, for the Notice for Non-conclusion of Agreement
 * (Issue Notices tab). "2" is a second wording; the corporation variants carry
 * that corporation's logo header over the original wording.
 */
export const NOTICE_TEMPLATE_VARIANTS: TemplateVariantOption[] = [
  { id: MMC_WORK_ORDER_TEMPLATE_VARIANT, label: 'Notice template 3 (MMC)' },
  { id: CMC_WORK_ORDER_TEMPLATE_VARIANT, label: 'Notice template 4 (CMC)' },
  { id: GHMC_WORK_ORDER_TEMPLATE_VARIANT, label: 'Notice template 5 (GHMC)' }
]

/** Notice template 2 is retired (legacy-templates/); an office with no corporation header gets the CMC one. */
export const DEFAULT_NOTICE_TEMPLATE_VARIANT = CMC_WORK_ORDER_TEMPLATE_VARIANT

/** Filename suffix for issue-notice[-ee]-template*.docx: '-<variant id>' (an unknown id gets the default's). */
export function noticeTemplateSuffix(variantId: string | undefined): string {
  const found = NOTICE_TEMPLATE_VARIANTS.find((v) => v.id === variantId)
  return `-${found?.id ?? DEFAULT_NOTICE_TEMPLATE_VARIANT}`
}

/**
 * Same variant-picker shape, for the five-letter QC Intimation/Completion set.
 * Type 1 is the plain-text letterhead; type 2 swaps the corporation-name heading for that
 * corporation's horizontal logo header (Work Order types 4-6) — corporation-only, like the others.
 */
export const QC_INTIMATION_COMPLETION_TEMPLATE_VARIANTS: TemplateVariantOption[] = [
  { id: MMC_WORK_ORDER_TEMPLATE_VARIANT, label: 'QC Intimation/Completion type 2 (MMC)' },
  { id: CMC_WORK_ORDER_TEMPLATE_VARIANT, label: 'QC Intimation/Completion type 2 (CMC)' },
  { id: GHMC_WORK_ORDER_TEMPLATE_VARIANT, label: 'QC Intimation/Completion type 2 (GHMC)' }
]

export const DEFAULT_QC_INTIMATION_COMPLETION_TEMPLATE_VARIANT = CMC_WORK_ORDER_TEMPLATE_VARIANT

/** Same naming convention as workOrderTemplateFileName, for qc-intimation-completion-template*.docx. */
export function qcIntimationCompletionTemplateFileName(variantId: string | undefined): string {
  const found = QC_INTIMATION_COMPLETION_TEMPLATE_VARIANTS.find((v) => v.id === variantId)
  return `qc-intimation-completion-template-${found?.id ?? DEFAULT_QC_INTIMATION_COMPLETION_TEMPLATE_VARIANT}.docx`
}

/**
 * Same variant-picker shape, for the File Backer — the file's cover page.
 * "header2" carries Nizampet Circle-58's own logo header (Telangana emblem +
 * "Cyberabad Municipal Corporation" title + CMC logo) in place of the app's
 * plain-text "{{Corp Full Caps}}" heading, matching the Work Order and
 * Intimation "header2"/"2" variants.
 */
export const FILE_BACKER_TEMPLATE_VARIANTS: TemplateVariantOption[] = [
  // MMC-only: the horizontal Malkajgiri Municipal Corporation logo header (Work Order type 4).
  { id: MMC_WORK_ORDER_TEMPLATE_VARIANT, label: 'File Backer type 3 (MMC)' },
  { id: CMC_WORK_ORDER_TEMPLATE_VARIANT, label: 'File Backer type 4 (CMC)' },
  { id: GHMC_WORK_ORDER_TEMPLATE_VARIANT, label: 'File Backer type 5 (GHMC)' }
]

export const DEFAULT_FILE_BACKER_TEMPLATE_VARIANT = CMC_WORK_ORDER_TEMPLATE_VARIANT

/** Same naming convention as workOrderTemplateFileName, for file-backer-template*.docx. */
export function fileBackerTemplateFileName(variantId: string | undefined): string {
  const found = FILE_BACKER_TEMPLATE_VARIANTS.find((v) => v.id === variantId)
  return `file-backer-template-${found?.id ?? DEFAULT_FILE_BACKER_TEMPLATE_VARIANT}.docx`
}

/**
 * Same variant-picker shape, for the full Civil Tender Document — "header2"
 * carries the same Nizampet Circle-58 logo header on both the page-1
 * Forwarding Slip and page-2 Notice Inviting Tenders, in place of the app's
 * default plain-text "{{Corp Full}}" heading.
 */
export const CIVIL_TENDER_TEMPLATE_VARIANTS: TemplateVariantOption[] = [
  // CMC-only: the horizontal CMC logo header (Work Order type 5) on the Forwarding Slip and NIT pages.
  { id: CMC_WORK_ORDER_TEMPLATE_VARIANT, label: 'Tender Document type 3 (CMC)' },
  { id: MMC_WORK_ORDER_TEMPLATE_VARIANT, label: 'Tender Document type 4 (MMC)' },
  { id: GHMC_WORK_ORDER_TEMPLATE_VARIANT, label: 'Tender Document type 5 (GHMC)' }
]

export const DEFAULT_CIVIL_TENDER_TEMPLATE_VARIANT = CMC_WORK_ORDER_TEMPLATE_VARIANT

/** Same naming convention as workOrderTemplateFileName, for civil-tender-template*.docx. */
export function civilTenderTemplateFileName(variantId: string | undefined): string {
  const found = CIVIL_TENDER_TEMPLATE_VARIANTS.find((v) => v.id === variantId)
  return `civil-tender-template-${found?.id ?? DEFAULT_CIVIL_TENDER_TEMPLATE_VARIANT}.docx`
}
