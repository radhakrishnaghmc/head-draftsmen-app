import { describe, expect, it } from 'vitest'
import { gluedLineCount } from '../core/ocrSpacing'

describe('gluedLineCount', () => {
  it('flags OCR lines whose words lost their spaces', () => {
    expect(
      gluedLineCount([
        'Name of Work Constrcutionof Mahila bhavan over existing seniorcitizenbuilding inTSllCcolonyparkinShapurWard291,GajularamaramCircle'
      ])
    ).toBe(1)
  })
  it('leaves normally spaced lines alone', () => {
    expect(
      gluedLineCount([
        'Name of Work Constrcution of Mahila bhavan over existing senior citizen building in TSllC colony park in Shapur Ward 291, Gajularamaram Circle',
        'BRK Associates 1384932-1 19/10/202611:15AM 106.214.2.187'
      ])
    ).toBe(0)
  })
})
