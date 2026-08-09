import { recognizeImage } from './ocr'

/**
 * OCR the GPS overlay stamped on a field photo and return its text lines.
 *
 * The overlay is white text over a photo, which a single OCR pass reads
 * unreliably, so we binarise the (greyscaled) image at several brightness
 * thresholds and OCR each — different thresholds recover different lines as the
 * scene behind the text varies — then return the union of all lines. The caller
 * parses the coordinates out of them (core/gpsOverlay.ts), preferring the passes
 * that captured the full °/'/" form. Runs off the shared, crash-isolated OCR
 * child (electron/ocr.ts).
 */
export async function ocrGpsOverlay(imageBuffer: Buffer): Promise<string[]> {
  const sharp = (await import('sharp')).default
  const meta = await sharp(imageBuffer).metadata()
  const W = meta.width ?? 0
  const H = meta.height ?? 0
  // Upscale small (WhatsApp-compressed) photos so the overlay text is big enough
  // for the recogniser; leave larger photos as-is.
  const width = Math.max(1800, W || 1800)
  const base = sharp(imageBuffer).rotate().grayscale().resize({ width })

  // Every line from every pass is kept (NOT de-duplicated): core/gpsOverlay.ts
  // votes across passes, so repeated identical reads must count.
  const lines: string[] = []
  const collect = async (buf: Buffer) => {
    let ocrLines
    try {
      ocrLines = await recognizeImage(buf)
    } catch {
      return
    }
    for (const l of ocrLines) {
      const s = l.text.trim()
      if (s) lines.push(s)
    }
  }

  // Pass set 1 — whole-image binarisation at several brightness thresholds.
  // Recovers the large, high-contrast centre overlay (DMS "N 17°52'…" stamp).
  for (const threshold of [205, 230, 250]) {
    try {
      await collect(await base.clone().threshold(threshold).negate().png().toBuffer())
    } catch {
      /* skip a failed pass */
    }
  }

  // Pass set 2 — several contrast-boosted (CLAHE) variants of the bottom-left
  // strip. The other common overlay is small, faint, semi-transparent text in the
  // bottom corner ("Lat/Long: Lat 17.611674 Long 78.162789"); thresholding loses
  // it, but local-contrast equalisation on an upscaled crop brings it out.
  // Several variants give the parser independent reads to vote on.
  if (W > 0 && H > 0) {
    const stripH = Math.max(1, Math.round(H * 0.34))
    const stripW = Math.max(1, Math.round(W * 0.65))
    const crop = (upscale: number) =>
      sharp(imageBuffer)
        .extract({ left: 0, top: H - stripH, width: stripW, height: stripH })
        .resize({ width: Math.min(2600, Math.round(stripW * upscale)) })
        .grayscale()
    const variants = [
      () => crop(2).clahe({ width: 15, height: 15 }),
      () => crop(3).clahe({ width: 8, height: 8 }),
      () => crop(3).normalise().clahe({ width: 20, height: 20 })
    ]
    for (const v of variants) {
      try {
        await collect(await v().png().toBuffer())
      } catch {
        /* skip a failed variant */
      }
    }
  }

  return lines
}
