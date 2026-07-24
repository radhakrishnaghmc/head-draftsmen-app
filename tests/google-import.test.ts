import { describe, expect, it } from 'vitest'
import { resolveGoogleDownloadUrl } from '../core/googleImport'

const SHEET_URL = 'https://docs.google.com/spreadsheets/d/abc123/edit?usp=sharing'
const SHEET_URL_WITH_GID = 'https://docs.google.com/spreadsheets/d/abc123/edit#gid=456789'

describe('resolveGoogleDownloadUrl', () => {
  it('exports the whole workbook when the link has no gid', () => {
    expect(resolveGoogleDownloadUrl(SHEET_URL)).toBe(
      'https://docs.google.com/spreadsheets/d/abc123/export?format=xlsx'
    )
  })

  it('respects a #gid= in the link by default (single-sheet import)', () => {
    expect(resolveGoogleDownloadUrl(SHEET_URL_WITH_GID)).toBe(
      'https://docs.google.com/spreadsheets/d/abc123/export?format=xlsx&gid=456789'
    )
  })

  it('ignores a #gid= in the link when wholeWorkbook is requested, so every sheet still downloads', () => {
    expect(resolveGoogleDownloadUrl(SHEET_URL_WITH_GID, { wholeWorkbook: true })).toBe(
      'https://docs.google.com/spreadsheets/d/abc123/export?format=xlsx'
    )
  })

  it('wholeWorkbook has no effect on a link that had no gid to begin with', () => {
    expect(resolveGoogleDownloadUrl(SHEET_URL, { wholeWorkbook: true })).toBe(
      'https://docs.google.com/spreadsheets/d/abc123/export?format=xlsx'
    )
  })
})
