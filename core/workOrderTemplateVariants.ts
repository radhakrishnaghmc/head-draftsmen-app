/**
 * Every bundled Work Order template variant an office can choose between, in
 * Settings' Document Templates section — different circles' own offices word
 * (and lay out) their Work Order differently, beyond what filling the same
 * placeholders into one shared template can express (see e.g. Kompally
 * Circle-56's real Work Order: a numbered Ref block citing NIT No/Tender ID/
 * Intimation Letter Date, "EXECUTIVE ENGINEER-II" sign-off, a different Work
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

export const WORK_ORDER_TEMPLATE_VARIANTS: TemplateVariantOption[] = [
  { id: 'default', label: 'Work Order type 1' },
  { id: 'kompally', label: 'Work Order type 2' },
  { id: 'header2', label: 'Work Order type 3' }
]

export const DEFAULT_WORK_ORDER_TEMPLATE_VARIANT = 'default'

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
  { id: 'header2', label: 'File Backer type 2' }
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
  { id: 'header2', label: 'Tender Document type 2' }
]

export const DEFAULT_CIVIL_TENDER_TEMPLATE_VARIANT = 'default'

/** Same naming convention as workOrderTemplateFileName, for civil-tender-template*.docx. */
export function civilTenderTemplateFileName(variantId: string | undefined): string {
  const found = CIVIL_TENDER_TEMPLATE_VARIANTS.find((v) => v.id === variantId)
  if (!found || found.id === DEFAULT_CIVIL_TENDER_TEMPLATE_VARIANT) return 'civil-tender-template.docx'
  return `civil-tender-template-${found.id}.docx`
}
