import { describe, it, expect } from 'vitest'
import { parseTenderEvaluation } from '../core/tenderEvaluationPdf'
import { SELECTED_ROW_MARK } from '../core/selectedRowMark'

const tie = (mark: boolean) => [
  'S VENKATESH 940568.00 Less 12.99 818388.22 L-1',
  `BONTHA RANGA RAJU 940568.00 Less 12.99 818388.22 L-1${mark ? ' ' + SELECTED_ROW_MARK : ''}`,
  'Vemula Ramulu 940568.00 Less 12.33 824595.97 L-3'
]

describe('tied L-1 bidders', () => {
  it('takes the bidder whose Select radio is ticked', () => {
    expect(parseTenderEvaluation(tie(true)).l1AgencyName).toBe('BONTHA RANGA RAJU')
  })
  it('falls back to the first L-1 when no tick is detected', () => {
    expect(parseTenderEvaluation(tie(false)).l1AgencyName).toBe('S VENKATESH')
  })
})
