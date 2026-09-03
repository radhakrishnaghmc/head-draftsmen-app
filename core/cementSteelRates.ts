export interface CementSteelRate {
  slNo: number
  description: string
  datePosted: string
  token: string
  /** File extension read out of the token's own path claim (e.g. "pdf"), for a sensible default save name. */
  ext: string
}
