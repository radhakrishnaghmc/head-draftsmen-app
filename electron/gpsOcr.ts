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

  // Each pass preprocesses the image a different way and OCRs it. They're
  // independent, so instead of running all 8 back-to-back on one photo we build
  // them as tasks and run them concurrently — the OCR child pool (electron/ocr.ts)
  // spreads them across CPU cores. Kept in pass order and NOT de-duplicated:
  // core/gpsOverlay.ts votes across passes (and, on a tie, favours the earlier
  // pass), so repeated reads must count and order must stay stable.
  const passes: Array<() => Promise<Buffer>> = []

  // Pass set 1 — whole-image binarisation at several brightness thresholds.
  // Recovers the large, high-contrast centre overlay (DMS "N 17°52'…" stamp).
  // 190 is included for fainter white-on-photo text.
  for (const threshold of [190, 205, 230, 250]) {
    passes.push(() => base.clone().threshold(threshold).negate().png().toBuffer())
  }

  // Pass set 2 — several contrast-boosted variants of the FULL-WIDTH bottom
  // strip. The other common overlay is small, faint, semi-transparent text in
  // the bottom corner ("Lat/Long: Lat 17.611674 Long 78.162789"); thresholding
  // loses it, but local-contrast equalisation / normalisation on an upscaled crop
  // brings it out. Full width (not 65%) so a right-aligned longitude is never
  // cropped off; the extra normalise/gamma variants recover very faint stamps on
  // bright backgrounds. Several variants give the parser independent reads to vote
  // on (core/gpsOverlay.ts also reconstructs a dropped decimal / a wrapped value).
  if (W > 0 && H > 0) {
    const stripH = Math.max(1, Math.round(H * 0.38))
    const crop = () =>
      sharp(imageBuffer)
        .extract({ left: 0, top: H - stripH, width: W, height: stripH })
        .resize({ width: Math.min(3200, Math.round(W * 2.2)) })
        .grayscale()
    passes.push(() => crop().clahe({ width: 15, height: 15 }).png().toBuffer())
    passes.push(() => crop().normalise().clahe({ width: 20, height: 20 }).png().toBuffer())
    passes.push(() => crop().normalise().sharpen().png().toBuffer())
    passes.push(() => crop().gamma(2.5).normalise().clahe({ width: 18, height: 18 }).png().toBuffer())
  }

  // Run one pass: build its preprocessed image, OCR it, return its trimmed lines.
  // A failed preprocess or a failed OCR pass just contributes nothing.
  const runPass = async (make: () => Promise<Buffer>): Promise<string[]> => {
    let buf: Buffer
    try {
      buf = await make()
    } catch {
      return []
    }
    try {
      const ocrLines = await recognizeImage(buf)
      return ocrLines.map((l) => l.text.trim()).filter((s) => s.length > 0)
    } catch {
      return []
    }
  }

  const perPass = await Promise.all(passes.map(runPass))
  return perPass.flat()
}
