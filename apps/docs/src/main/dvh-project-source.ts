/**
 * Project data in Docs (P10): the main-process side of the project store.
 *
 * - `dvh-project://<id>` is a link source like a workbook: reading it gives
 *   the project's Smart Data model, writing fields to it goes through the P3
 *   write-back check, and every change tells the documents linked to it.
 * - Schema packs are installed from a pack file; objects are edited from the
 *   Smart Data panel.
 * - Two-way exchange with QLCL-DVH: the partner file (a QLCL workbook or a
 *   dvh-exchange JSON) is synced by id, the merged data written back to it;
 *   conflicts wait for the user and touch neither side.
 */

import { readFile } from 'node:fs/promises'
import { extname } from 'node:path'
import {
  serializeModelXml,
  type FieldWriteResult,
  type Scalar,
  type WriteFieldsRequest,
} from '@genoffice/dvh-model'
import {
  createObject,
  deleteObject,
  installPack,
  parseProjectData,
  parseSchemaPack,
  projectIdOfUri,
  projectSourceModel,
  projectTypes,
  projectUri,
  pendingConflicts,
  resolveAllSyncConflicts,
  resolveSyncConflict,
  setValues,
  syncProject,
  validateProject,
  writeProjectFields,
  type ProjectData,
  type SyncConflict,
} from '@genoffice/dvh-project'
import {
  exportExchangeJson,
  exportQlclWorkbook,
  importExchangeJson,
  importQlclWorkbook,
} from '@genoffice/dvh-template'
import type { DvhProjectInfo as ProjectInfo } from '../shared/ipc'
import { atomicWriteFile } from './atomic-write'

/** What this module needs from the project store (ProjectStore in the app). */
export interface ProjectDataStore {
  resolveProjectForFile(filePath: string): string
  ensureDefaultProject(): { id: string; name: string }
  getProject(projectId: string): { id: string; name: string } | null
  readProjectData(projectId: string): unknown | null
  writeProjectData(projectId: string, data: unknown): void
}

export class ProjectDataService {
  constructor(
    private readonly store: () => ProjectDataStore,
    /** tells the documents linked to a project that it changed */
    private readonly changed: (uri: string) => void,
    private readonly now: () => string = () => new Date().toISOString(),
  ) {}

  /** The project of a document (unsaved documents use the default project). */
  projectOf(filePath: string | null): string {
    const store = this.store()
    return filePath ? store.resolveProjectForFile(filePath) : store.ensureDefaultProject().id
  }

  load(projectId: string): ProjectData {
    return parseProjectData(this.store().readProjectData(projectId), projectId)
  }

  private save(data: ProjectData): void {
    this.store().writeProjectData(data.projectId, data)
    this.changed(projectUri(data.projectId))
  }

  info(projectId: string): ProjectInfo {
    const data = this.load(projectId)
    return {
      projectId,
      name: this.store().getProject(projectId)?.name ?? projectId,
      uri: projectUri(projectId),
      docId: data.docId,
      revision: data.revision,
      packs: data.packs.map((p) => ({ id: p.id, name: p.name, packVersion: p.packVersion })),
      types: projectTypes(data).map((t) => ({
        name: t.name,
        ...(t.label ? { label: t.label } : {}),
        single: t.single === true,
        count: data.objects.filter((o) => o.type === t.name).length,
        fields: t.fields.map((f) => ({
          key: f.key,
          ...(f.label ? { label: f.label } : {}),
          type: f.type,
          ...(f.enumValues ? { enumValues: f.enumValues } : {}),
        })),
      })),
      issues: validateProject(data).slice(0, 100),
      partners: Object.entries(data.sync).map(([path, base]) => ({
        path,
        at: base.at,
        conflicts: base.conflicts ?? [],
      })),
    }
  }

  /** The `dvh-project://` source as a model part, as readDvhSource gives a workbook's. */
  readSource(uri: string): { path: string; modelXml: string | null } | null {
    const projectId = projectIdOfUri(uri)
    if (!projectId || !this.store().getProject(projectId)) return null
    const data = this.load(projectId)
    return {
      path: uri,
      modelXml: data.objects.length ? serializeModelXml(projectSourceModel(data)) : null,
    }
  }

  /** Field write-back into the project; null when the request is not for a project. */
  writeFields(request: WriteFieldsRequest): {
    via: 'file'
    path: string
    results: FieldWriteResult[]
  } | null {
    const uri = request.paths.find((p) => projectIdOfUri(p) !== null)
    const projectId = uri ? projectIdOfUri(uri) : null
    if (!uri || !projectId) return null
    const data = this.load(projectId)
    if (data.docId !== request.docId) return { via: 'file', path: uri, results: [] }
    const out = writeProjectFields(data, request.writes, {
      source: 'link',
      action: 'Link.WriteBack',
      at: this.now(),
      ...(request.origin.name ? { actor: { user: request.origin.name } } : {}),
    })
    if (out.data !== data) this.save(out.data)
    return { via: 'file', path: uri, results: out.results }
  }

  installPack(projectId: string, packJson: string): ProjectInfo {
    const pack = parseSchemaPack(JSON.parse(packJson))
    this.save(
      installPack(this.load(projectId), pack, {
        source: 'ui',
        action: 'Project.InstallPack',
        at: this.now(),
      }),
    )
    return this.info(projectId)
  }

  objects(projectId: string, type: string): { id: string; values: Record<string, Scalar> }[] {
    return this.load(projectId)
      .objects.filter((o) => o.type === type)
      .map((o) => ({ id: o.id, values: o.values }))
  }

  /** Creates (objectId null) or edits an object; returns its id. */
  setObject(
    projectId: string,
    request: { type: string; objectId: string | null; values: Record<string, Scalar> },
  ): string {
    const ctx = { source: 'ui' as const, action: 'Project.Edit', at: this.now() }
    const data = this.load(projectId)
    if (request.objectId === null) {
      const created = createObject(data, request.type, request.values, ctx)
      this.save(created.data)
      return created.object.id
    }
    this.save(setValues(data, request.objectId, request.values, ctx))
    return request.objectId
  }

  deleteObject(projectId: string, objectId: string): void {
    this.save(
      deleteObject(this.load(projectId), objectId, {
        source: 'ui',
        action: 'Project.Delete',
        at: this.now(),
      }),
    )
  }

  /** Syncs with a partner file and writes the merged data back to it. */
  async sync(
    projectId: string,
    partnerPath: string,
  ): Promise<{ stats: Record<string, number>; conflicts: readonly SyncConflict[] }> {
    const json = extname(partnerPath).toLowerCase() === '.json'
    const bytes = new Uint8Array(await readFile(partnerPath))
    const incoming = json
      ? importExchangeJson(new TextDecoder().decode(bytes))
      : await importQlclWorkbook(bytes)
    const at = this.now()
    const result = syncProject(this.load(projectId), incoming, partnerPath, {
      source: 'external',
      action: 'Project.Sync',
      at,
    })
    const out = json
      ? Buffer.from(exportExchangeJson(result.outgoing, at), 'utf8')
      : Buffer.from(await exportQlclWorkbook(result.outgoing))
    // the partner file first: if it is locked (open in Excel), nothing changes here either
    await atomicWriteFile(partnerPath, out)
    // pending conflicts are part of the saved data: they survive a restart
    this.save(result.data)
    return { stats: { ...result.stats }, conflicts: result.conflicts }
  }

  /**
   * Settles one pending conflict of a partner (by its place in the pending
   * list), or all of them the same way (index null).
   */
  resolve(
    projectId: string,
    partnerPath: string,
    index: number | null,
    keep: 'mine' | 'theirs',
  ): void {
    const data = this.load(projectId)
    const ctx = { source: 'ui' as const, action: 'Project.ResolveConflict', at: this.now() }
    if (index === null) {
      this.save(resolveAllSyncConflicts(data, partnerPath, keep, ctx))
      return
    }
    const conflict = pendingConflicts(data, partnerPath)[index]
    if (conflict) this.save(resolveSyncConflict(data, partnerPath, conflict, keep, ctx))
  }
}
