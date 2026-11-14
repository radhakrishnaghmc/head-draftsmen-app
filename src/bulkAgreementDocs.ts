import { api } from './ipc'
import { TEMPLATE_KEYS } from './office'
import {
  FILE_BACKER_TEMPLATE_VARIANTS,
  WORK_ORDER_TEMPLATE_VARIANTS,
  corporationDefaultVariant,
  variantForCorporation
} from '@core/workOrderTemplateVariants'
import { fileBackerPlaceholders, agreementPlaceholders, workOrderPlaceholders } from '@core/workOrderAgreement'
import type { StoredLetter } from '@core/types'

type AgreementDocKind = Extract<StoredLetter['kind'], 'filebacker' | 'agreementbond' | 'workorder'>

/**
 * The Agreement-workspace documents Settings can make in bulk from a folder of L-1 sheets and LOAs. Same shape as the
 * QC / Forwarding / DB Note letters: only the placeholder values are stored, the template is re-filled on demand.
 */
export const AGREEMENT_DOCS: Record<
  AgreementDocKind,
  {
    title: string
    docName: string
    /** Prefix of the downloaded file name. */
    fileLabel: string
    templateKey: string
    template: (variant: string) => Promise<string>
    /** The office's picked variant → the one to use (same rule the Agreement workspace applies). */
    variantFor: (picked: string | null, corporation: string | undefined) => string
    placeholders: typeof fileBackerPlaceholders
  }
> = {
  filebacker: {
    title: 'Bulk File Backers',
    docName: 'File Backer',
    fileLabel: 'File Backer',
    templateKey: TEMPLATE_KEYS.fileBacker,
    template: (v) => api.fileBackerTemplate(v),
    variantFor: (picked, corp) =>
      variantForCorporation(picked ?? corporationDefaultVariant(corp, undefined), corp, FILE_BACKER_TEMPLATE_VARIANTS) ?? 'cmc',
    placeholders: fileBackerPlaceholders
  },
  agreementbond: {
    title: 'Bulk Agreement Bonds',
    docName: 'Agreement Bond',
    fileLabel: 'Agreement Bond',
    templateKey: TEMPLATE_KEYS.agreement,
    template: (v) => api.agreementTemplate(v),
    variantFor: (picked) => picked ?? 'default',
    placeholders: agreementPlaceholders
  },
  workorder: {
    title: 'Bulk Work Orders',
    docName: 'Work Order',
    fileLabel: 'Work Order',
    templateKey: TEMPLATE_KEYS.workOrder,
    template: (v) => api.workOrderTemplate(v),
    variantFor: (picked, corp) =>
      variantForCorporation(picked ?? corporationDefaultVariant(corp, undefined), corp, WORK_ORDER_TEMPLATE_VARIANTS) ?? 'cmc',
    placeholders: workOrderPlaceholders
  }
}

export const isAgreementDoc = (k: string): k is AgreementDocKind => k in AGREEMENT_DOCS
