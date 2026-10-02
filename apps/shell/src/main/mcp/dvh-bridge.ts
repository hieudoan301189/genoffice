import type { WebContents } from 'electron'
import {
  createBridgeEndpoint,
  type DvhBridgeRequest,
  type DvhBridgeTarget,
} from '@genoffice/dvh-actions'
import { newDvhId } from '@genoffice/dvh-model'
import { applySessionTxn } from '../../../../slides/src/main/slides-main'
import { sessions } from '../../../../slides/src/main/session-state'
import {
  createSlidesDvhActions,
  slidesSagaParticipant,
} from '../../../../slides/src/main/dvh-actions'
import type { OpenDocumentTab } from '../../shared/tabs-api'
import type { DocsControl } from './tools/document-tools'
import type { DvhActionsControl, DvhFamily } from './tools/dvh-tools'
import { FAMILY_BY_KIND, resolveOpenDocument } from './tools/open-documents-tools'
import type { SheetsControl } from './tools/sheets-tools'

/**
 * Shell half of the DVH action bridge (P4). Docs and Sheets answer bridge
 * requests in their renderers (`dvh_actions` over their MCP bridges); a
 * Slides session lives in this process, so its endpoint is built here, one
 * per tab, over the same applySessionTxn the app's own batch surface uses.
 */

export interface DvhBridgeDeps {
  list: () => Promise<OpenDocumentTab[]>
  webContentsFor: (tabId: string) => WebContents | undefined
  docs?: DocsControl
  sheets?: SheetsControl
  confirm: (title: string, lines: readonly string[]) => Promise<boolean>
  bulkThreshold: () => number
}

const slidesEndpoints = new Map<number, (request: unknown) => Promise<unknown>>()

function slidesEndpoint(wcId: number): (request: unknown) => Promise<unknown> {
  const known = slidesEndpoints.get(wcId)
  if (known) return known
  const host = {
    session: () => sessions.get(wcId) ?? null,
    applyTxn: applySessionTxn,
  }
  const registry = createSlidesDvhActions(host)
  const docId = newDvhId('doc')
  const target: DvhBridgeTarget = {
    docId,
    registry,
    participant: slidesSagaParticipant(host, registry, docId),
  }
  const endpoint = createBridgeEndpoint(() => (host.session() ? target : null))
  slidesEndpoints.set(wcId, endpoint)
  return endpoint
}

export function createDvhActionsControl(deps: DvhBridgeDeps): DvhActionsControl {
  return {
    async resolve(document) {
      const documents = await deps.list()
      const doc = resolveOpenDocument(documents, document)
      if (!doc) throw new Error(`no open document matches "${document}"; see open_documents`)
      const family = FAMILY_BY_KIND[doc.kind]
      if (family !== 'docx' && family !== 'xlsx' && family !== 'pptx') {
        throw new Error(
          `"${doc.title}" has no DVH actions (only Word, Excel and PowerPoint documents do)`,
        )
      }
      const contents = deps.webContentsFor(doc.id)
      if (!contents || contents.isDestroyed()) throw new Error(`"${doc.title}" is no longer open`)
      return { wcId: contents.id, family, title: doc.title }
    },
    async send(wcId: number, family: DvhFamily, request: DvhBridgeRequest) {
      if (family === 'docx') {
        if (!deps.docs) throw new Error('the Word editor is unavailable in this build')
        return deps.docs.runCommand(wcId, 'dvh_actions', request)
      }
      if (family === 'xlsx') {
        if (!deps.sheets) throw new Error('the spreadsheet editor is unavailable in this build')
        return deps.sheets.runCommand(wcId, 'dvh_actions', request)
      }
      return slidesEndpoint(wcId)(request)
    },
    confirm: deps.confirm,
    bulkThreshold: deps.bulkThreshold,
  }
}
