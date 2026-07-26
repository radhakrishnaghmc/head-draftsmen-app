import { describe, expect, it } from 'vitest'
import { driveFolderId, parseDriveFolderListing, collectFolderPdfs } from '../core/googleDriveFolder'

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

// A file entry links to /file/d/<id>; a subfolder entry links to /drive/folders/<id>.
function fileEntry(id: string, name: string): string {
  return `<div class="flip-entry" id="entry-${id}"><a href="https://drive.google.com/file/d/${id}/view?usp=drive_web"><div class="flip-entry-title">${name}</div></a></div>`
}
function folderEntry(id: string, name: string): string {
  return `<div class="flip-entry" id="entry-${id}"><a href="https://drive.google.com/drive/folders/${id}"><div class="flip-entry-title">${name}</div></a></div>`
}

describe('parseDriveFolderListing', () => {
  it('extracts files and subfolders, tagging which is which', () => {
    const html = [
      folderEntry('SUB1', 'Nizampet works'),
      fileEntry('F_A', 'Stage Selected Form l1.pdf'),
      fileEntry('F_B', 'notes.txt')
    ].join('\n')
    expect(parseDriveFolderListing(html)).toEqual([
      { id: 'SUB1', name: 'Nizampet works', isFolder: true },
      { id: 'F_A', name: 'Stage Selected Form l1.pdf', isFolder: false },
      { id: 'F_B', name: 'notes.txt', isFolder: false }
    ])
  })
  it('dedupes repeated entries and ignores empty HTML', () => {
    expect(parseDriveFolderListing(fileEntry('X', 'a.pdf') + fileEntry('X', 'a.pdf'))).toHaveLength(1)
    expect(parseDriveFolderListing('<html>nothing</html>')).toEqual([])
  })
})

describe('collectFolderPdfs', () => {
  // A tree: root -> [Circle-58 (sub), Contractors (sub, must skip), top.pdf]
  //         Circle-58 -> [deeper (sub), work1.pdf]
  //         deeper -> [work2.pdf]
  //         Contractors -> [should-never-see.pdf]
  const tree: Record<string, string> = {
    ROOT: folderEntry('C58', 'Circle-58') + folderEntry('CON', 'Contractors') + fileEntry('T', 'top.pdf'),
    C58: folderEntry('DEEP', 'deeper') + fileEntry('W1', 'work1.pdf'),
    DEEP: fileEntry('W2', 'work2.pdf'),
    CON: fileEntry('NOPE', 'should-never-see.pdf')
  }
  const fetchHtml = async (id: string) => tree[id] ?? ''

  it('recurses into subfolders and collects every PDF', async () => {
    const pdfs = await collectFolderPdfs('ROOT', fetchHtml)
    const names = pdfs.map((p) => p.name).sort()
    expect(names).toEqual(['top.pdf', 'work1.pdf', 'work2.pdf'])
  })

  it('never descends into a Contractors folder', async () => {
    const pdfs = await collectFolderPdfs('ROOT', fetchHtml)
    expect(pdfs.find((p) => p.name === 'should-never-see.pdf')).toBeUndefined()
  })

  it('does not loop forever on a self-referential tree', async () => {
    const loop: Record<string, string> = { A: folderEntry('B', 'b') + fileEntry('P', 'x.pdf'), B: folderEntry('A', 'a') }
    const pdfs = await collectFolderPdfs('A', async (id) => loop[id] ?? '')
    expect(pdfs.map((p) => p.name)).toEqual(['x.pdf'])
  })
})
