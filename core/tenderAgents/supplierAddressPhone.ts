export interface SupplierAddressPhone {
  /** Agency's postal address, from Street Address + City + Region/State + Zip. */
  address?: string
  /** Agency's mobile/phone — the only one of the Supplier Details / Intimation sources that carries this. */
  phone?: string
}

/**
 * Agent: Address & Phone Number — read from the Telangana eGP portal's
 * "Supplier Details" page (the bidder's own profile, saved/printed as
 * .pdf), not the Intimation/LOA letter (see agencyAddress.ts) or the L1
 * sheet — neither of those carries a phone number at all. Each field is
 * "Label value" on its own line, e.g.:
 *
 *   Street Address 2-1-373, Janmabhoomi Colony, Yellammabanda.
 *   City RANGAREDDY
 *   Region/Province/State Telangana
 *   Zip/Postal Code 500072
 *   Mobile Number 9666054562
 *
 * Best-effort: returns {} for whichever field its label isn't found.
 */
export function detectSupplierAddressPhone(lines: string[]): SupplierAddressPhone {
  const result: SupplierAddressPhone = {}
  const field = (label: string): string => {
    const re = new RegExp(`^${label}\\s*[:\\-]?\\s*(.+)$`, 'i')
    for (const line of lines) {
      const m = re.exec(line.trim())
      if (m && m[1].trim()) return m[1].trim().replace(/\s+/g, ' ')
    }
    return ''
  }

  const street = field('Street Address')
  const cityState = [field('City'), field('Region/Province/State')].filter(Boolean).join(', ')
  const zip = field('Zip/Postal Code')
  const address = [street, cityState].filter(Boolean).join(', ') + (zip ? ` - ${zip}` : '')
  if (address.trim().length > 0) result.address = address.trim()

  const mobile = field('Mobile Number')
  if (mobile) result.phone = mobile

  return result
}
