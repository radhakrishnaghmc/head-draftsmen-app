import path from 'path'
import type Ocr from '@gutenye/ocr-node'
import type { OcrLine } from './ocr'

// Runs as an isolated CHILD PROCESS (child_process.fork), not a worker thread.
// A large photo can trip a native SIGTRAP deep inside onnxruntime's memory
// arena on some machines; inside a worker thread that fault takes down the
// whole Electron process, but in a separate OS process it only kills this
// child — letting the parent (electron/ocr.ts) retry the page at a smaller,
// long-proven-safe resolution instead of the whole app going down.

let ocrPromise: Promise<Ocr> | null = null

async function getOcr(modelDir: string): Promise<Ocr> {
  if (!ocrPromise) {
    ocrPromise = (async () => {
      const { default: OcrClass } = await import('@gutenye/ocr-node')
      return OcrClass.create({
        models: {
          detectionPath: path.join(modelDir, 'ch_PP-OCRv4_det_infer.onnx'),
          recognitionPath: path.join(modelDir, 'ch_PP-OCRv4_rec_infer.onnx'),
          dictionaryPath: path.join(modelDir, 'ppocr_keys_v1.txt')
        }
      })
    })()
  }
  return ocrPromise
}

interface Request {
  id: number
  imagePath: string
  modelDir: string
}

process.on('message', async (msg: Request) => {
  try {
    const ocr = await getOcr(msg.modelDir)
    const result = await ocr.detect(msg.imagePath)
    const lines: OcrLine[] = result.map((line) => {
      const box = line.box ?? []
      const ys = box.map((p) => p[1])
      return { text: line.text, top: ys.length > 0 ? Math.min(...ys) : 0 }
    })
    process.send?.({ id: msg.id, lines })
  } catch (e) {
    process.send?.({ id: msg.id, error: e instanceof Error ? e.message : String(e) })
  }
})
