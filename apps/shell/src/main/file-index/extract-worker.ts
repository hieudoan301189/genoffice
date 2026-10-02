/**
 * Worker-thread entry: text extraction and directory walks run here so a slow
 * PDF or a stalled network folder never blocks the main process.
 */
import { parentPort } from 'node:worker_threads'
import { extractDvhDocId, extractText } from './extract'
import { scanFiles } from './scan'

export type WorkerRequest =
  { id: number; type: 'extract'; path: string } | { id: number; type: 'scan'; root: string }

export type WorkerResponse =
  | {
      id: number
      type: 'extract'
      result: Awaited<ReturnType<typeof extractText>>
      /** the DVH docId of an Office file with Smart Data (P3 link resolution) */
      dvhDocId?: string | null
    }
  | { id: number; type: 'scan'; files: ReturnType<typeof scanFiles> }

parentPort?.on('message', async (req: WorkerRequest) => {
  if (req.type === 'extract') {
    const result = await extractText(req.path)
    const dvhDocId = await extractDvhDocId(req.path)
    parentPort?.postMessage({
      id: req.id,
      type: 'extract',
      result,
      dvhDocId,
    } satisfies WorkerResponse)
  } else {
    parentPort?.postMessage({
      id: req.id,
      type: 'scan',
      files: scanFiles(req.root),
    } satisfies WorkerResponse)
  }
})
