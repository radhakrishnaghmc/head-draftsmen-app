import sharp from 'sharp'
import { writeFile } from 'fs/promises'
import path from 'path'
import { fileURLToPath } from 'url'

// NSIS only reads 24-bit BMPs. sharp can't write BMP, so encode raw RGB ourselves.
const here = path.dirname(fileURLToPath(import.meta.url))
const icon = path.join(here, 'icon.png')

async function toBmp(svg, w, h, out) {
  const data = await sharp(Buffer.from(svg)).removeAlpha().raw().toBuffer()
  const rowSize = Math.ceil((w * 3) / 4) * 4
  const buf = Buffer.alloc(54 + rowSize * h)
  buf.write('BM', 0)
  buf.writeUInt32LE(buf.length, 2)
  buf.writeUInt32LE(54, 10)
  buf.writeUInt32LE(40, 14)
  buf.writeInt32LE(w, 18)
  buf.writeInt32LE(h, 22)
  buf.writeUInt16LE(1, 26)
  buf.writeUInt16LE(24, 28)
  buf.writeUInt32LE(rowSize * h, 34)
  for (let y = 0; y < h; y++) {
    for (let x = 0; x < w; x++) {
      const s = (y * w + x) * 3
      const d = 54 + (h - 1 - y) * rowSize + x * 3 // BMP rows run bottom-up, BGR
      buf[d] = data[s + 2]
      buf[d + 1] = data[s + 1]
      buf[d + 2] = data[s]
    }
  }
  await writeFile(path.join(here, out), buf)
}

async function iconDataUri(size) {
  const png = await sharp(icon).resize(size, size).png().toBuffer()
  return `data:image/png;base64,${png.toString('base64')}`
}

const gradient =
  '<defs><linearGradient id="g" x1="0" y1="0" x2="1" y2="1"><stop offset="0" stop-color="#0b3b66"/><stop offset="1" stop-color="#0ea5e9"/></linearGradient></defs>'

// Welcome/finish page sidebar: 164x314.
const sidebar = `<svg xmlns="http://www.w3.org/2000/svg" width="164" height="314">
${gradient}
<rect width="164" height="314" fill="url(#g)"/>
<circle cx="150" cy="40" r="90" fill="#fff" fill-opacity="0.06"/>
<circle cx="20" cy="300" r="70" fill="#fff" fill-opacity="0.06"/>
<image href="${await iconDataUri(96)}" x="34" y="60" width="96" height="96"/>
<text x="82" y="196" font-family="Arial, sans-serif" font-size="15" font-weight="700" fill="#fff" text-anchor="middle">Head Draughtsman</text>
<text x="82" y="216" font-family="Arial, sans-serif" font-size="15" font-weight="700" fill="#fff" text-anchor="middle">App</text>
<rect x="62" y="230" width="40" height="3" rx="1.5" fill="#7dd3fc"/>
</svg>`

// Installing-page header strip: 150x57.
const header = `<svg xmlns="http://www.w3.org/2000/svg" width="150" height="57">
${gradient}
<rect width="150" height="57" fill="url(#g)"/>
<image href="${await iconDataUri(40)}" x="12" y="8" width="40" height="40"/>
<text x="60" y="26" font-family="Arial, sans-serif" font-size="11" font-weight="700" fill="#fff">Head</text>
<text x="60" y="40" font-family="Arial, sans-serif" font-size="11" font-weight="700" fill="#fff">Draughtsman</text>
</svg>`

await toBmp(sidebar, 164, 314, 'installerSidebar.bmp')
await toBmp(sidebar, 164, 314, 'uninstallerSidebar.bmp')
await toBmp(header, 150, 57, 'installerHeader.bmp')
console.log('Generated installerSidebar.bmp, uninstallerSidebar.bmp, installerHeader.bmp')
