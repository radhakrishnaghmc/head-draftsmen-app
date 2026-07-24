import { parentPort } from 'worker_threads'

// Runs on its own OS thread — the whole point is to keep the ~15s+ model
// load and batch-inference work off Electron's single-threaded main process,
// which would otherwise freeze the entire app (every IPC call, window
// repaint, everything) for the duration, and can get force-quit by the OS
// as "not responding".

const MODEL_NAME = 'all-MiniLM-L6-v2'

let pipelinePromise: Promise<import('@huggingface/transformers').FeatureExtractionPipeline> | null = null

async function getPipeline(modelDir: string) {
  if (!pipelinePromise) {
    pipelinePromise = (async () => {
      const { env, pipeline } = await import('@huggingface/transformers')
      // Never reach out to the network — the model is fully bundled.
      env.allowRemoteModels = false
      env.allowLocalModels = true
      env.localModelPath = modelDir
      // Node's build of this library only supports cpu/coreml/webgpu
      // execution providers (no "wasm" device — that's web-build only), so
      // onnxruntime-node's native binary is unavoidable here.
      return pipeline('feature-extraction', MODEL_NAME, { device: 'cpu', dtype: 'q8' })
    })()
  }
  return pipelinePromise
}

interface Request {
  id: number
  texts: string[]
  modelDir: string
}

parentPort?.on('message', async (msg: Request) => {
  try {
    const extractor = await getPipeline(msg.modelDir)
    const output = await extractor(msg.texts, { pooling: 'mean', normalize: true })
    parentPort?.postMessage({ id: msg.id, vectors: output.tolist() })
  } catch (e) {
    parentPort?.postMessage({ id: msg.id, error: e instanceof Error ? e.message : String(e) })
  }
})
