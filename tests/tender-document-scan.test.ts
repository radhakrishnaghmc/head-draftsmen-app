import { describe, it, expect, beforeEach, afterEach } from 'vitest'
import * as fs from 'fs'
import * as os from 'os'
import * as path from 'path'
import { execSync } from 'child_process'
import { collectTenderDocuments } from '../electron/tenderDocumentScan'

// A throwaway tree shaped like a real office's Tender Evaluations folder
// (see ~/…/GAJULARAMARAM CIRCLE-57-CMC/Tender Evaluations and NIZAMPET
// CIRCLE-58-CMC/Tender Evaluations-58): one folder per work, each holding
// the actual L1/Intimation sheet alongside a "Common Documents" subfolder
// full of unrelated bidder KYC PDFs (PAN, GST, ITR, litigation history, …)
// that should never be read.
let root: string

beforeEach(() => {
  root = fs.mkdtempSync(path.join(os.tmpdir(), 'tender-scan-test-'))
  const work1 = path.join(root, '16).710635-CC road at Sai Keerthi')
  fs.mkdirSync(work1, { recursive: true })
  fs.writeFileSync(path.join(work1, 'evaluation sheet.pdf'), 'l1 sheet content')
  fs.writeFileSync(path.join(work1, 'intimation.pdf'), 'intimation content')
  fs.writeFileSync(path.join(work1, 'Stage Selected Form.pdf'), 'l1 sheet content')
  const commonDocs = path.join(work1, 'Dobbala Anil', 'Common Documents')
  fs.mkdirSync(commonDocs, { recursive: true })
  fs.writeFileSync(path.join(commonDocs, '7256599-RegnPAN GST.pdf'), 'pan card image')
  fs.writeFileSync(path.join(commonDocs, '10702443-Turnover Certificate.pdf'), 'turnover cert')
  fs.writeFileSync(path.join(commonDocs, '10808309-NO LITIGATION HISTORY.pdf'), 'litigation decl')

  const work2 = path.join(root, '2).709803-LED Lights-Recall')
  fs.mkdirSync(work2, { recursive: true })
  fs.writeFileSync(path.join(work2, 'L1.pdf'), 'l1 sheet content 2')
  fs.writeFileSync(path.join(work2, 'tender.telangana.gov.in_viewIntimationNotice.html'), '<html></html>')

  // A non-PDF/HTML file that happens to have "L1" in its name — extension
  // filter should exclude it even though the name would otherwise match.
  fs.writeFileSync(path.join(root, 'L1 notes.txt'), 'not a real document')
})

afterEach(() => {
  fs.rmSync(root, { recursive: true, force: true })
})

describe('collectTenderDocuments', () => {
  it('recursively finds L1/Intimation-named PDFs and HTML inside a picked folder, skipping unrelated KYC documents', async () => {
    const found = await collectTenderDocuments([root])
    const names = found.map((f) => f.name).sort()
    expect(names).toEqual([
      'L1.pdf',
      'Stage Selected Form.pdf',
      'evaluation sheet.pdf',
      'intimation.pdf',
      'tender.telangana.gov.in_viewIntimationNotice.html'
    ])
  })

  it('never includes a file whose extension is not .pdf/.html/.htm, even if the name matches', async () => {
    const found = await collectTenderDocuments([root])
    expect(found.some((f) => f.name === 'L1 notes.txt')).toBe(false)
  })

  it('never reads the unrelated KYC documents in a "Common Documents" subfolder', async () => {
    const found = await collectTenderDocuments([root])
    const names = found.map((f) => f.name)
    expect(names).not.toContain('7256599-RegnPAN GST.pdf')
    expect(names).not.toContain('10702443-Turnover Certificate.pdf')
    expect(names).not.toContain('10808309-NO LITIGATION HISTORY.pdf')
  })

  it('returns the real file bytes, not just names', async () => {
    const found = await collectTenderDocuments([root])
    const l1 = found.find((f) => f.name === 'L1.pdf')!
    expect(Buffer.from(l1.bytes).toString('utf8')).toBe('l1 sheet content 2')
  })

  it('includes a file explicitly picked by path, regardless of its name — the filename filter only applies inside a recursed folder', async () => {
    const oddPath = path.join(root, 'unrelated-name-but-explicitly-picked.pdf')
    fs.writeFileSync(oddPath, 'picked directly')
    const found = await collectTenderDocuments([oddPath])
    expect(found).toHaveLength(1)
    expect(found[0].name).toBe('unrelated-name-but-explicitly-picked.pdf')
  })

  it('accepts a mix of an explicit file and a folder in one call', async () => {
    const oddPath = path.join(root, 'unrelated-name-but-explicitly-picked.pdf')
    fs.writeFileSync(oddPath, 'picked directly')
    const found = await collectTenderDocuments([oddPath, path.join(root, '2).709803-LED Lights-Recall')])
    const names = found.map((f) => f.name).sort()
    expect(names).toEqual([
      'L1.pdf',
      'tender.telangana.gov.in_viewIntimationNotice.html',
      'unrelated-name-but-explicitly-picked.pdf'
    ])
  })

  it('reports scanning progress as matches are found, then reading progress with a real total/percentage', async () => {
    const events: { phase: string; done: number; total: number }[] = []
    const found = await collectTenderDocuments([root], (phase, done, total) => {
      events.push({ phase, done, total })
    })
    const scanning = events.filter((e) => e.phase === 'scanning')
    const reading = events.filter((e) => e.phase === 'reading')
    expect(scanning.length).toBe(found.length) // one 'scanning' tick per match found
    expect(reading.length).toBe(found.length) // one 'reading' tick per file actually read
    // 'reading' has a real, fixed total from the very first tick — 'scanning' never does.
    expect(reading.every((e) => e.total === found.length)).toBe(true)
    expect(reading[reading.length - 1]).toEqual({ phase: 'reading', done: found.length, total: found.length })
  })

  // Real report: a picked folder on a cloud-sync drive (OneDrive/Google
  // Drive "on-demand" placeholder, or a dropped network share) can leave a
  // single file's read stuck forever with no error of its own — a named
  // pipe (FIFO) that nothing ever writes to reproduces exactly that: fs
  // reads it but the read never resolves on its own.
  it('skips a file whose read never completes (e.g. a stalled cloud-sync placeholder) instead of hanging the whole scan forever', async () => {
    const stuckPath = path.join(root, 'L1-stuck.pdf')
    execSync(`mkfifo "${stuckPath}"`)
    const found = await collectTenderDocuments([root])
    const names = found.map((f) => f.name)
    expect(names).not.toContain('L1-stuck.pdf')
    // The rest of the folder's real files still come through despite the stuck one.
    expect(names).toContain('L1.pdf')
    expect(names).toContain('evaluation sheet.pdf')
  }, 25_000)
})
