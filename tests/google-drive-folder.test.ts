import { describe, expect, it } from 'vitest'
import { driveFolderId, parseDriveFolderListing } from '../core/googleDriveFolder'

describe('driveFolderId', () => {
  it('extracts the id from a standard folder share link', () => {
    expect(driveFolderId('https://drive.google.com/drive/folders/1AbC_dEf-123?usp=sharing')).toBe('1AbC_dEf-123')
  })
  it('handles a /u/0/ account-scoped folder link', () => {
    expect(driveFolderId('https://drive.google.com/drive/u/0/folders/XYZ789')).toBe('XYZ789')
  })
  it('handles an ?id= style link', () => {
    expect(driveFolderId('https://drive.google.com/open?id=FOLDER42')).toBe('FOLDER42')
  })
  it('returns null for a non-folder link', () => {
    expect(driveFolderId('https://docs.google.com/spreadsheets/d/abc/edit')).toBeNull()
  })
})

describe('parseDriveFolderListing', () => {
  const html = `
    <div class="flip-entry" id="entry-FILE_A"><div class="flip-entry-thumb"></div>
      <div class="flip-entry-title">Stage Selected Form.pdf</div></div>
    <div class="flip-entry" id="entry-FILE_B"><div class="flip-entry-title">Stage Selected Form l1.pdf</div></div>
    <div class="flip-entry" id="entry-FILE_C"><div class="flip-entry-title">notes.txt</div></div>
    <div class="flip-entry" id="entry-FILE_A"><div class="flip-entry-title">Stage Selected Form.pdf</div></div>
  `
  it('extracts each file id + name, deduped, in order', () => {
    expect(parseDriveFolderListing(html)).toEqual([
      { id: 'FILE_A', name: 'Stage Selected Form.pdf' },
      { id: 'FILE_B', name: 'Stage Selected Form l1.pdf' },
      { id: 'FILE_C', name: 'notes.txt' }
    ])
  })
  it('returns an empty array when there are no entries', () => {
    expect(parseDriveFolderListing('<html><body>nothing here</body></html>')).toEqual([])
  })
})
