export interface AgencyAddressBlock {
  /** Contractor/company name, from the "To" address block. */
  agencyName?: string
  /** The agency's postal address — the address-block lines below the name, joined with ", ". */
  address?: string
}

/**
 * Agent: Address of the Agency — the winning bidder's postal address, read
 * from the *printed Intimation / Letter of Acceptance* letter (not the L1
 * sheet — the L1 Commercial Evaluation page never carries an address at
 * all, only the company name). Input is the reconstructed text lines of
 * that letter (see src/pdfToText.ts). The letter's layout:
 *
 *   To,
 *   M V S CONSTRUCTIONS                          <- agency name
 *   13/B, Allwyn Colony, … Telangana             <- address (until "Phone No"/"Sub")
 *   ...
 *
 * Best-effort: returns {} when the letter carries no recognizable "To," block.
 */
export function detectAgencyAddress(lines: string[]): AgencyAddressBlock {
  const result: AgencyAddressBlock = {}
  const toIdx = lines.findIndex((l) => /^to\s*[,:]?\s*$/i.test(l))
  if (toIdx < 0) return result

  const after = lines
    .slice(toIdx + 1)
    .map((l) => l.trim())
    .filter((l) => l.length > 0)
  if (after.length > 0) result.agencyName = after[0]

  const addr: string[] = []
  for (const l of after.slice(1)) {
    if (/^(phone\s*no|sub\s*[:.]|ref\s*[:.]|sir|madam)/i.test(l)) break
    addr.push(l)
  }
  if (addr.length > 0) result.address = addr.join(', ')

  return result
}
