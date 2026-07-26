import { describe, expect, it } from 'vitest'
import { driveFolderId, parseDriveFolderListing, collectFolderFiles, isIntimationFile } from '../core/googleDriveFolder'

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

describe('isIntimationFile', () => {
  it('recognizes intimation notices as .html or .html.pdf', () => {
    expect(isIntimationFile('viewIntimationNoticealeadp circle.html')).toBe(true)
    expect(isIntimationFile('tender.telangana.gov.in_viewIntimationNotice.html.pdf')).toBe(true)
    expect(isIntimationFile('L1.pdf')).toBe(false)
    expect(isIntimationFile('7406082-ravi aadhar.pdf')).toBe(false)
  })
})

describe('collectFolderFiles', () => {
  // The real tree nests files inconsistently:
  //  - W1 (NIT.12): tender PDFs + an intimation .html directly in the work
  //    folder, agency subfolders alongside carry the bidder's own docs.
  //  - W3 (NIT.02): the tender PDFs are inside the agency's "Common Documents"
  //    subfolder, mixed with bidder docs, plus an intimation .html.pdf.
  const tree: Record<string, string> = {
    ROOT: folderEntry('NIT12', 'NIT.12') + folderEntry('NIT02', 'NIT.02(FY-026-27)'),
    NIT12: folderEntry('W1', '1.717574-Aleap circle'),
    W1:
      folderEntry('AG1', 'M V S CONSTRUCTIONS') +
      folderEntry('ASSETS', 'viewIntimationNoticealeadp circle_files') +
      fileEntry('P1', 'Stage Selected Form l1.pdf') +
      fileEntry('P2', 'Stage Selected Form.pdf') +
      fileEntry('H1', 'viewIntimationNoticealeadp circle.html') +
      fileEntry('X', '717574-Evaluations.xlsx'),
    AG1: fileEntry('BID1', '7406082-ravi aadhar.pdf') + fileEntry('BID2', '7406086-ravi new gst.pdf'),
    ASSETS: fileEntry('IMG', 'image001.pdf'),
    NIT02: folderEntry('W3', '1.iTEM.01-699966'),
    W3: folderEntry('AG3', 'N K INFRA PROJECTS'),
    AG3: folderEntry('CD', 'Common Documents'),
    CD:
      fileEntry('P3', 'L1.pdf') +
      fileEntry('P4', 'Stage Selected Form 1.pdf') +
      fileEntry('H2', 'tender.telangana.gov.in_viewIntimationNotice.html.pdf') +
      fileEntry('BID3', '10971896-gst naresh.pdf')
  }
  const fetchHtml = async (id: string) => tree[id] ?? ''

  it('collects tender PDFs and intimation notices, excluding bidder docs by filename', async () => {
    const { tenderPdfs, intimationFiles } = await collectFolderFiles('ROOT', fetchHtml)
    expect(tenderPdfs.map((p) => p.name).sort()).toEqual([
      'L1.pdf',
      'Stage Selected Form 1.pdf',
      'Stage Selected Form l1.pdf',
      'Stage Selected Form.pdf'
    ])
    expect(intimationFiles.map((p) => p.id).sort()).toEqual(['H1', 'H2'])
    const all = [...tenderPdfs, ...intimationFiles].map((p) => p.id)
    expect(all).not.toContain('BID1')
    expect(all).not.toContain('BID3')
    expect(all).not.toContain('IMG') // an intimation _files asset folder isn't descended into
  })

  it('does not loop forever on a self-referential tree', async () => {
    const loop: Record<string, string> = {
      A: folderEntry('B', 'b'),
      B: folderEntry('A', 'a') + fileEntry('P', 'L1.pdf')
    }
    const { tenderPdfs } = await collectFolderFiles('A', async (id) => loop[id] ?? '')
    expect(tenderPdfs.map((p) => p.name)).toEqual(['L1.pdf'])
  })
})
