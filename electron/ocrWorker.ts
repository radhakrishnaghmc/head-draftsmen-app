import { parentPort } from 'worker_threads'
import path from 'path'
import type Ocr from '@gutenye/ocr-node'
import type { OcrLine } from './ocr'

// Runs on its own OS thread, same reasoning as embeddingsWorker.ts — OCR on
// a full-page photo can take a few seconds, and doing that on Electron's
// single-threaded main process would freeze the entire app (every IPC
// call, window repaint) for the duration.

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

parentPort?.on('message', async (msg: Request) => {
  try {
    const ocr = await getOcr(msg.modelDir)
    const result = await ocr.detect(msg.imagePath)
    const lines: OcrLine[] = result.map((line) => {
      const box = line.box ?? []
      const ys = box.map((p) => p[1])
      return { text: line.text, top: ys.length > 0 ? Math.min(...ys) : 0 }
    })
    parentPort?.postMessage({ id: msg.id, lines })
  } catch (e) {
    parentPort?.postMessage({ id: msg.id, error: e instanceof Error ? e.message : String(e) })
  }
})
