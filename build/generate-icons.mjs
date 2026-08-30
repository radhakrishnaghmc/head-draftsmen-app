import sharp from 'sharp'
import pngToIco from 'png-to-ico'
import { mkdir, rm, writeFile } from 'fs/promises'
import path from 'path'
import { fileURLToPath } from 'url'

const here = path.dirname(fileURLToPath(import.meta.url))
const sourcePath = path.join(here, 'icon-source.png')
const iconsetDir = path.join(here, 'icon.iconset')

// icon-source.png is pure black line art on a transparent background. Left
// as-is, that's invisible against a dark taskbar/desktop/dock (Windows dark
// theme in particular). Bake in a white rounded-square backing plate so the
// glyph stays visible on any background, on every platform.
const CANVAS = 1024
const PLATE_MARGIN = 32
const PLATE_RADIUS = 220
const GLYPH_SIZE = 720

async function buildFlatSource() {
  const plateSvg = Buffer.from(
    `<svg xmlns="http://www.w3.org/2000/svg" width="${CANVAS}" height="${CANVAS}">` +
      `<rect x="${PLATE_MARGIN}" y="${PLATE_MARGIN}" width="${CANVAS - PLATE_MARGIN * 2}" height="${CANVAS - PLATE_MARGIN * 2}" rx="${PLATE_RADIUS}" fill="#ffffff"/>` +
      `</svg>`
  )
  const glyph = await sharp(sourcePath, { limitInputPixels: false })
    .resize(GLYPH_SIZE, GLYPH_SIZE, { fit: 'contain', background: { r: 0, g: 0, b: 0, alpha: 0 } })
    .toBuffer()
  const offset = Math.round((CANVAS - GLYPH_SIZE) / 2)
  return sharp(plateSvg)
    .composite([{ input: glyph, left: offset, top: offset }])
    .png()
    .toBuffer()
}

async function main() {
  await rm(iconsetDir, { recursive: true, force: true })
  await mkdir(iconsetDir, { recursive: true })

  const flatSource = await buildFlatSource()

  // macOS .iconset naming: icon_16x16.png, icon_16x16@2x.png (= 32px), etc.
  const pairs = [
    [16, 'icon_16x16.png'],
    [32, 'icon_16x16@2x.png'],
    [32, 'icon_32x32.png'],
    [64, 'icon_32x32@2x.png'],
    [128, 'icon_128x128.png'],
    [256, 'icon_128x128@2x.png'],
    [256, 'icon_256x256.png'],
    [512, 'icon_256x256@2x.png'],
    [512, 'icon_512x512.png'],
    [1024, 'icon_512x512@2x.png']
  ]
  for (const [size, name] of pairs) {
    await sharp(flatSource)
      .resize(size, size, { fit: 'contain', background: { r: 0, g: 0, b: 0, alpha: 0 } })
      .png()
      .toFile(path.join(iconsetDir, name))
  }

  // Linux / general-purpose PNG.
  await sharp(flatSource)
    .resize(512, 512, { fit: 'contain', background: { r: 0, g: 0, b: 0, alpha: 0 } })
    .png()
    .toFile(path.join(here, 'icon.png'))

  // Windows .ico from a few key sizes.
  const icoSizes = [16, 24, 32, 48, 64, 128, 256]
  const icoBuffers = []
  for (const size of icoSizes) {
    icoBuffers.push(
      await sharp(flatSource)
        .resize(size, size, { fit: 'contain', background: { r: 0, g: 0, b: 0, alpha: 0 } })
        .png()
        .toBuffer()
    )
  }
  const icoBuffer = await pngToIco(icoBuffers)
  await writeFile(path.join(here, 'icon.ico'), icoBuffer)

  console.log('Generated icon.png, icon.ico, and icon.iconset/ — run iconutil next for icon.icns')
}

main().catch((e) => {
  console.error(e)
  process.exit(1)
})
