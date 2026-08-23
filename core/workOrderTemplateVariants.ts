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
  { id: 'kompally', label: 'Work Order type 2' }
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
