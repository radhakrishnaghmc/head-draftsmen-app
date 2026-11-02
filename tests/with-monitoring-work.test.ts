import { describe, expect, it } from 'vitest'
import { withMonitoringWork } from '../core/workOrderAgreement'

describe('withMonitoringWork', () => {
  const row = { 'Name of the work': 'X', 'Amount of estimate': '9.50' }

  it('overrides the ECV stand-in estimate and fills the Win code from the Monitoring Format', () => {
    const out = withMonitoringWork(row, { estimateAmt: 12.5, winCode: ' W123 ' })
    expect(out['Amount of estimate']).toBe('12.5')
    expect(out.Wincode).toBe('W123')
  })

  it('leaves the row alone without a match or with blank MF values', () => {
    expect(withMonitoringWork(row, null)).toBe(row)
    expect(withMonitoringWork(row, { estimateAmt: 0, winCode: '' })).toEqual(row)
  })
})
