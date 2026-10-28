import { describe, expect, it } from 'vitest'
import { detectNameOfWork } from '../core/tenderAgents/nameOfWork'
import { stripItemNoTag } from '../core/tenderAgents/shared'

describe('detectNameOfWork — title wrapped around the label with a capitalized second half', () => {
  it('keeps the first title line that sits above "Name of Work" and drops the NIT cell tail', () => {
    const lines = [
      'Current Tender Details',
      'NIT No.21/DB/EE-III/JDM/C-',
      'Enquiry/IFB/Tender',
      'Tender ID 742555 55/QBZ/CMC/2026-27, Dt:12.10.2026 (Item',
      'Notice Number',
      'no.6)',
      'Laying of CC road from H.No : 48-516 , 48-522/1 to 48-537/1/1(left side end) and 48-527/2(Right side end) at Ganesh',
      'Name of Work Nagar under Ganesh Nagar-284, Jeedimetla Circle-55, Qpz, CMC (Reserved for SC) (Item no.6) (Completion Peroid 2',
      'Months)',
      'Tender Category Works Tender Evaluation Type Percentage'
    ]
    const name = detectNameOfWork(lines) ?? ''
    expect(name.startsWith('Laying of CC road from H.No')).toBe(true)
    expect(name).toContain('at Ganesh Nagar under Ganesh Nagar-284')
    expect(name).not.toMatch(/^no\.6\)/i)
  })
})

describe('stripItemNoTag — completion period tag', () => {
  it('drops "(Item no.6) (Completion Peroid 2 Months)" after the reservation tag', () => {
    expect(stripItemNoTag('Laying of CC road at Ganesh Nagar, CMC (Reserved for SC) (Item no.6) (Completion Peroid 2 Months)')).toBe(
      'Laying of CC road at Ganesh Nagar, CMC (Reserved for SC)'
    )
  })
  it('handles the reverse order and the correct spelling', () => {
    expect(stripItemNoTag('Laying of road, CMC (Completion Period: 3 Months) (Item No. 4)')).toBe('Laying of road, CMC')
  })
  it('leaves a title without tags alone', () => {
    expect(stripItemNoTag('Laying of road, CMC')).toBe('Laying of road, CMC')
  })
})

describe('stripItemNoTag — keeps the reservation tag', () => {
  it('removes item + completion period tags anywhere but keeps "(Reserved for SC)"', () => {
    expect(stripItemNoTag('Laying CC road, CMC (Item no.6) (Reserved for SC) (Completion Peroid 2 Months)')).toBe(
      'Laying CC road, CMC (Reserved for SC)'
    )
  })
})
