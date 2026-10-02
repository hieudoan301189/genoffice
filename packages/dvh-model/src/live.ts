/**
 * Live link channel (P3 automatic links): a workbook open in Sheets publishes
 * its Smart Data after every edit; the main process relays it to every open
 * document, which updates the links fed by that workbook without a save.
 */

import { z } from 'zod'

/** renderer → main (ipcRenderer.send) */
export const DVH_LIVE_PUBLISH_CHANNEL = 'dvh:live-publish'
/** main → document renderers */
export const DVH_LIVE_DOCS_CHANNEL = 'docs:dvh-live'

export const dvhLivePayloadSchema = z
  .object({
    /** the workbook's DVH docId: links match on it, so a renamed or unsaved file still feeds them */
    docId: z.string().min(1),
    /** the workbook's path on disk, when it has one */
    path: z.string().nullable(),
    /** the model as serializeModelXml writes it, values read from the cells just now */
    modelXml: z
      .string()
      .min(1)
      .max(64 * 1024 * 1024),
  })
  .strict()
export type DvhLivePayload = z.infer<typeof dvhLivePayloadSchema>
