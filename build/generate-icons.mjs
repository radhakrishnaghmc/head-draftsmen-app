import sharp from 'sharp'
import pngToIco from 'png-to-ico'
import { mkdir, rm, writeFile } from 'fs/promises'
import path from 'path'
import { fileURLToPath } from 'url'

const here = path.dirname(fileURLToPath(import.meta.url))
const source = path.join(here, 'icon-source.png')
const iconsetDir = path.join(here, 'icon.iconset')

const icnsSizes = [16, 32, 64, 128, 256, 512, 1024]

async function main() {
  await rm(iconsetDir, { recursive: true, force: true })
  await mkdir(iconsetDir, { recursive: true })

  // macOS .iconset naming: icon_16x16.png, icon_16x16@2x.png (= 32px), etc.
  for (const size of icnsSizes) {
    const base = size / 2
    if (icnsSizes.includes(base) || base === 8) {
      // covered by the smaller size's @2x below where applicable
    }
  }
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
    await sharp(source, { limitInputPixels: false })
      .resize(size, size, { fit: 'contain', background: { r: 0, g: 0, b: 0, alpha: 0 } })
      .png()
      .toFile(path.join(iconsetDir, name))
  }

  // Linux / general-purpose PNG.
  await sharp(source)
    .resize(512, 512, { fit: 'contain', background: { r: 0, g: 0, b: 0, alpha: 0 } })
    .png()
    .toFile(path.join(here, 'icon.png'))

  // Windows .ico from a few key sizes.
  const icoSizes = [16, 24, 32, 48, 64, 128, 256]
  const icoBuffers = []
  for (const size of icoSizes) {
    icoBuffers.push(
      await sharp(source, { limitInputPixels: false })
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
