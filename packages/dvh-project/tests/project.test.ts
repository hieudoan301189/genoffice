/**
 * P10: the project data model. Schema packs carry the business vocabulary
 * (the core has none, D10); objects are validated against types and rules; a
 * document links to project objects through the virtual source and writes
 * fields back with the P3 conflict check; the two-way exchange with QLCL-DVH
 * matches objects by id and detects conflicts like P3.
 */
import { readdirSync, readFileSync } from 'node:fs'
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'
import { parseModelXml, serializeModelXml, type DvhModel } from '@genoffice/dvh-model'
import { exportQlclWorkbook, importQlclWorkbook } from '@genoffice/dvh-template'
import {
  createObject,
  deleteObject,
  emptyProjectData,
  installPack,
  parseProjectData,
  parseSchemaPack,
  pendingConflicts,
  resolveAllSyncConflicts,
  projectFieldId,
  projectIdOfUri,
  projectSourceModel,
  projectUri,
  resolveSyncConflict,
  setValues,
  syncProject,
  validateProject,
  writeProjectFields,
  type ProjectData,
} from '../src/index'

const PACK = parseSchemaPack(
  JSON.parse(readFileSync(join(__dirname, '../packs/qlcl-xay-dung/pack.json'), 'utf8')),
)
const ui = { source: 'ui' as const, action: 'Project.Edit', at: '2026-10-02T08:00:00.000Z' }

function project(): ProjectData {
  let data = installPack(emptyProjectData('p1'), PACK, ui)
  data = createObject(
    data,
    'Contractor',
    { Code: 'NT-01', Name: 'Công ty Xây dựng Số 1' },
    ui,
    'o_nt1',
  ).data
  data = createObject(
    data,
    'Project',
    { Code: 'DA-01', Name: 'Cầu Rạch Miễu 2', Owner: 'Ban QLDA 7', Contractor: 'o_nt1' },
    ui,
    'o_da',
  ).data
  data = createObject(
    data,
    'WorkItem',
    {
      Code: 'HM-01',
      Name: 'Móng trụ T1',
      Unit: 'm3',
      Quantity: '125.5',
      Contractor: 'o_nt1',
      Status: 'Đang thi công',
    },
    ui,
    'o_hm1',
  ).data
  data = createObject(
    data,
    'WorkItem',
    { Code: 'HM-02', Name: 'Thân trụ T1', Unit: 'm3', Quantity: 80, Status: 'Chưa thi công' },
    ui,
    'o_hm2',
  ).data
  return data
}

describe('schema packs and objects', () => {
  it('the core has no business vocabulary: it comes from the pack (D10)', () => {
    const src = join(__dirname, '../src')
    for (const file of readdirSync(src)) {
      const text = readFileSync(join(src, file), 'utf8')
      expect(text, file).not.toMatch(/WorkItem|Contractor|Acceptance|Material|nghiệm thu|hạng mục/)
    }
    expect(PACK.types.map((t) => t.name)).toEqual([
      'Project',
      'Contractor',
      'WorkItem',
      'Acceptance',
      'Material',
      'Test',
    ])
  })

  it('rejects broken packs and clashing types', () => {
    expect(() =>
      parseSchemaPack({ ...PACK, types: [{ name: 'A', fields: [{ key: 'x', type: 'ref' }] }] }),
    ).toThrow(/names its target type/)
    expect(() =>
      parseSchemaPack({
        ...PACK,
        rules: [{ id: 'r', type: 'Nope', expr: '1', message: 'm', severity: 'error' }],
      }),
    ).toThrow(/unknown type Nope/)
    expect(() =>
      parseSchemaPack({
        ...PACK,
        types: [{ name: 'A', fields: [{ key: 'id', type: 'text' }] }],
        rules: [],
      }),
    ).toThrow(/reserved/)
    const other = parseSchemaPack({
      ...PACK,
      id: 'other',
      types: [{ name: 'Project', fields: [{ key: 'x', type: 'text' }] }],
      rules: [],
    })
    expect(() => installPack(project(), other, ui)).toThrow(/already comes from another pack/)
  })

  it('values carry revisions and every change is a change set', () => {
    let data = project()
    const hm1 = () => data.objects.find((o) => o.id === 'o_hm1')!
    expect(hm1().values.Quantity).toBe(125.5)
    expect(hm1().revs).toMatchObject({ Quantity: 1 })
    data = setValues(data, 'o_hm1', { Quantity: '130', Name: 'Móng trụ T1' }, ui)
    expect(hm1().values.Quantity).toBe(130)
    expect(hm1().revs).toMatchObject({ Quantity: 2, Name: 1 })
    expect(data.history.at(-1)).toMatchObject({
      action: 'Project.Edit',
      changes: [{ objectId: 'o_hm1', path: 'values.Quantity', before: 125.5, after: 130 }],
    })
    expect(() => setValues(data, 'o_hm1', { Colour: 'x' }, ui)).toThrow(/no field Colour/)
    expect(() => createObject(data, 'Project', {}, ui)).toThrow(/one per project/)
    data = deleteObject(data, 'o_hm2', ui)
    expect(data.objects.map((o) => o.id)).not.toContain('o_hm2')
    // a damaged store starts empty instead of failing
    expect(parseProjectData({ format: 'nope' }, 'p1').objects).toEqual([])
    expect(parseProjectData(JSON.parse(JSON.stringify(data)), 'p1')).toEqual(data)
  })

  it('validates types, references and the pack rules', () => {
    let data = project()
    data = createObject(
      data,
      'Acceptance',
      { Code: 'BB-01', WorkItem: 'o_hm1', Result: 'Đạt' },
      ui,
      'o_bb1',
    ).data
    data = createObject(
      data,
      'Acceptance',
      { Code: 'BB-02', WorkItem: 'o_gone', Result: 'Không đạt' },
      ui,
      'o_bb2',
    ).data
    data = setValues(data, 'o_hm2', { Quantity: -3, Status: 'Xong' }, ui)
    const issues = validateProject(data).map((i) => `${i.objectId} ${i.severity} ${i.message}`)
    expect(issues).toEqual([
      'o_hm2 error WorkItem.Status must be one of Chưa thi công, Đang thi công, Đã nghiệm thu',
      'o_bb2 error Acceptance.WorkItem refers to a missing WorkItem',
      'o_hm2 error Khối lượng hạng mục phải lớn hơn 0',
      'o_bb1 warning Hạng mục đã nghiệm thu đạt nhưng trạng thái chưa là "Đã nghiệm thu"',
      'o_bb2 error Biên bản không đạt phải ghi rõ lý do',
    ])
  })
})

describe('documents link to project objects', () => {
  it('the virtual source: fields for one-per-project types, collections with an ID column', () => {
    const data = project()
    expect(projectIdOfUri(projectUri('p1'))).toBe('p1')
    expect(projectIdOfUri('D:\\a.xlsx')).toBeNull()
    const model = projectSourceModel(data)
    expect(model.docId).toBe('proj_p1')
    expect(model.fields.map((f) => [f.name, f.value])).toEqual([
      ['Project.Code', 'DA-01'],
      ['Project.Name', 'Cầu Rạch Miễu 2'],
      ['Project.Owner', 'Ban QLDA 7'],
      ['Project.Location', null],
      ['Project.StartDate', null],
      ['Project.Contractor', 'o_nt1'],
    ])
    const items = model.collections.find((c) => c.name === 'WorkItem')!
    expect(items.columns.map((c) => c.title)).toEqual([
      'ID',
      'Mã hạng mục',
      'Tên hạng mục',
      'Đơn vị',
      'Khối lượng',
      'Nhà thầu',
      'Trạng thái',
    ])
    expect(items.rows[0]).toEqual([
      'o_hm1',
      'HM-01',
      'Móng trụ T1',
      'm3',
      125.5,
      'o_nt1',
      'Đang thi công',
    ])
    // it is a valid model part (a document stores it like any linked source)
    expect(parseModelXml(serializeModelXml(model))).toEqual(model)
  })

  it('writes fields back with the P3 conflict check', () => {
    let data = project()
    const id = projectFieldId('o_da', 'Name')
    const base = { rev: 1, text: 'Cầu Rạch Miễu 2' }
    const doc = { source: 'link' as const, action: 'Link.WriteBack' }
    const first = writeProjectFields(
      data,
      [{ fieldId: id, value: 'Cầu Rạch Miễu II', expected: base, force: false }],
      doc,
    )
    expect(first.results).toEqual([
      { fieldId: id, status: 'written', current: { rev: 2, text: 'Cầu Rạch Miễu II' } },
    ])
    data = first.data
    // a second document still on the old base: conflict, nothing written
    const stale = writeProjectFields(
      data,
      [{ fieldId: id, value: 'Cầu RM2', expected: base, force: false }],
      doc,
    )
    expect(stale.results[0]).toMatchObject({
      status: 'conflict',
      current: { rev: 2, text: 'Cầu Rạch Miễu II' },
    })
    expect(stale.data).toBe(data)
    const forced = writeProjectFields(
      data,
      [{ fieldId: id, value: 'Cầu RM2', expected: base, force: true }],
      doc,
    )
    expect(forced.results[0]!.status).toBe('written')
    expect(
      writeProjectFields(
        data,
        [{ fieldId: 'pf_o_none_Name', value: 'x', expected: null, force: false }],
        doc,
      ).results[0]!.status,
    ).toBe('missing')
  })
})

describe('two-way exchange with QLCL-DVH', () => {
  const sync = {
    source: 'external' as const,
    action: 'Project.Sync',
    at: '2026-10-02T09:00:00.000Z',
  }
  const partner = 'D:\\QLCL\\DA-01.xlsx'
  /** the partner's file: written by us, read back as QLCL-DVH would */
  const roundTrip = async (model: DvhModel) => importQlclWorkbook(await exportQlclWorkbook(model))
  const edit = (
    model: DvhModel,
    type: string,
    id: string,
    title: string,
    value: string | number,
  ) => {
    const c = model.collections.find((x) => x.name === type)!
    const col = c.columns.findIndex((x) => x.title === title)
    const row = c.rows.find((r) => r[0] === id)!
    row[col] = value
    return model
  }

  it('first sync pulls the partner; later both sides merge by id; conflicts stay on both sides', async () => {
    // the partner (QLCL-DVH) has the project and one work item; we only have the pack
    const theirs = projectSourceModel(project())
    const workItems = theirs.collections.find((c) => c.name === 'WorkItem')!
    workItems.rows = workItems.rows.slice(0, 1)
    const incoming = await roundTrip(theirs)
    let data = installPack(emptyProjectData('p1'), PACK, ui)
    const first = syncProject(data, incoming, partner, sync)
    expect(first.conflicts).toEqual([])
    expect(first.stats).toMatchObject({ created: 3 })
    data = first.data
    expect(data.objects.find((o) => o.id === 'o_hm1')!.values).toMatchObject({
      Name: 'Móng trụ T1',
      Quantity: 125.5,
    })
    expect(data.objects.find((o) => o.type === 'Project')!.values.Name).toBe('Cầu Rạch Miễu 2')

    // both sides edit: different values merge, the same value both ways is a conflict
    data = setValues(data, 'o_hm1', { Status: 'Đã nghiệm thu', Unit: 'm³' }, ui)
    let partnerFile = await roundTrip(first.outgoing)
    partnerFile = edit(partnerFile, 'WorkItem', 'o_hm1', 'Khối lượng', 130)
    partnerFile = edit(partnerFile, 'WorkItem', 'o_hm1', 'Đơn vị', 'khối')
    // a row the partner added without an id
    partnerFile.collections
      .find((c) => c.name === 'WorkItem')!
      .rows.push(['', 'HM-03', 'Xà mũ T1', 'm3', 40, null, 'Chưa thi công'])
    const second = syncProject(data, await roundTrip(partnerFile), partner, sync)
    data = second.data
    const hm1 = data.objects.find((o) => o.id === 'o_hm1')!.values
    expect(hm1).toMatchObject({ Quantity: 130, Status: 'Đã nghiệm thu', Unit: 'm³' })
    expect(second.conflicts).toEqual([
      {
        objectId: 'o_hm1',
        type: 'WorkItem',
        key: 'Unit',
        kind: 'value',
        base: 'm3',
        local: 'm³',
        remote: 'khối',
      },
    ])
    const hm3 = data.objects.find((o) => o.values.Code === 'HM-03')!
    expect(hm3.id).toMatch(/^o_/)
    // the outgoing file: our status, the new id, and the partner's own value where we disagree
    const out = second.outgoing.collections.find((c) => c.name === 'WorkItem')!
    expect(out.rows.find((r) => r[0] === 'o_hm1')).toEqual([
      'o_hm1',
      'HM-01',
      'Móng trụ T1',
      'khối',
      130,
      'o_nt1',
      'Đã nghiệm thu',
    ])
    expect(out.rows.some((r) => r[0] === hm3.id)).toBe(true)

    // the user keeps theirs: taken now, and the next sync is quiet
    data = resolveSyncConflict(data, partner, second.conflicts[0]!, 'theirs', ui)
    expect(data.objects.find((o) => o.id === 'o_hm1')!.values.Unit).toBe('khối')
    const third = syncProject(data, await roundTrip(second.outgoing), partner, sync)
    expect(third.conflicts).toEqual([])
    expect(third.stats).toMatchObject({ pulled: 0, created: 0, deleted: 0 })
  })

  it('"keep mine" is pushed at the next sync', async () => {
    let data = installPack(emptyProjectData('p1'), PACK, ui)
    const start = await roundTrip(projectSourceModel(project()))
    data = syncProject(data, start, partner, sync).data
    data = setValues(data, 'o_hm2', { Name: 'Thân trụ T1 (sửa)' }, ui)
    const theirs = edit(
      await roundTrip(start),
      'WorkItem',
      'o_hm2',
      'Tên hạng mục',
      'Thân trụ số 1',
    )
    const s = syncProject(data, theirs, partner, sync)
    expect(s.conflicts).toHaveLength(1)
    data = resolveSyncConflict(s.data, partner, s.conflicts[0]!, 'mine', ui)
    const next = syncProject(data, await roundTrip(s.outgoing), partner, sync)
    expect(next.conflicts).toEqual([])
    expect(
      next.outgoing.collections
        .find((c) => c.name === 'WorkItem')!
        .rows.find((r) => r[0] === 'o_hm2')![2],
    ).toBe('Thân trụ T1 (sửa)')
  })

  it('deletions are three-way: unchanged → deleted, changed → conflict', async () => {
    let data = installPack(emptyProjectData('p1'), PACK, ui)
    const start = await roundTrip(projectSourceModel(project()))
    data = syncProject(data, start, partner, sync).data
    // they delete HM-01 (unchanged here) and HM-02 (edited here)
    data = setValues(data, 'o_hm2', { Quantity: 81 }, ui)
    const theirs = await roundTrip(start)
    const items = theirs.collections.find((c) => c.name === 'WorkItem')!
    items.rows = items.rows.filter((r) => r[0] !== 'o_hm1' && r[0] !== 'o_hm2')
    const s = syncProject(data, theirs, partner, sync)
    expect(s.data.objects.some((o) => o.id === 'o_hm1')).toBe(false)
    expect(s.conflicts).toEqual([
      expect.objectContaining({ objectId: 'o_hm2', kind: 'deleted-there' }),
    ])
    expect(s.data.objects.some((o) => o.id === 'o_hm2')).toBe(true)
    // we delete the contractor they edited: it comes back as a conflict, the partner keeps it
    let here = deleteObject(s.data, 'o_nt1', ui)
    here = resolveSyncConflict(here, partner, s.conflicts[0]!, 'theirs', ui)
    const theirs2 = edit(
      await roundTrip(s.outgoing),
      'Contractor',
      'o_nt1',
      'Người đại diện',
      'Ông A',
    )
    const s2 = syncProject(here, theirs2, partner, sync)
    expect(s2.conflicts).toEqual([
      expect.objectContaining({ objectId: 'o_nt1', kind: 'deleted-here' }),
    ])
    expect(
      s2.outgoing.collections
        .find((c) => c.name === 'Contractor')!
        .rows.some((r) => r[0] === 'o_nt1'),
    ).toBe(true)
    expect(s2.data.objects.some((o) => o.id === 'o_hm2')).toBe(false)
    const back = resolveSyncConflict(s2.data, partner, s2.conflicts[0]!, 'theirs', ui)
    expect(back.objects.find((o) => o.id === 'o_nt1')!.values.Representative).toBe('Ông A')
  })

  it('pending conflicts are saved with the data, settled one by one or all at once', async () => {
    let data = installPack(emptyProjectData('p1'), PACK, ui)
    const start = await roundTrip(projectSourceModel(project()))
    data = syncProject(data, start, partner, sync).data
    data = setValues(data, 'o_hm1', { Name: 'Móng T1 (đây)', Unit: 'cái' }, ui)
    data = setValues(data, 'o_hm2', { Name: 'Thân T1 (đây)' }, ui)
    let theirs = edit(await roundTrip(start), 'WorkItem', 'o_hm1', 'Tên hạng mục', 'Móng T1 (QLCL)')
    theirs = edit(theirs, 'WorkItem', 'o_hm1', 'Đơn vị', 'tấn')
    theirs = edit(theirs, 'WorkItem', 'o_hm2', 'Tên hạng mục', 'Thân T1 (QLCL)')
    const s = syncProject(data, theirs, partner, sync)
    expect(s.conflicts).toHaveLength(3)
    // they survive a restart: stored in the project data
    const reopened = parseProjectData(JSON.parse(JSON.stringify(s.data)), 'p1')
    expect(pendingConflicts(reopened, partner)).toEqual(s.conflicts)

    // one settled: off the list; settling it again changes nothing
    const first = pendingConflicts(reopened, partner)[0]!
    data = resolveSyncConflict(reopened, partner, first, 'theirs', ui)
    expect(pendingConflicts(data, partner)).toHaveLength(2)
    expect(resolveSyncConflict(data, partner, first, 'mine', ui)).toBe(data)
    expect(data.objects.find((o) => o.id === 'o_hm1')!.values.Name).toBe('Móng T1 (QLCL)')

    // the rest at once: mine, pushed by the next sync, which is then quiet
    data = resolveAllSyncConflicts(data, partner, 'mine', ui)
    expect(pendingConflicts(data, partner)).toEqual([])
    const next = syncProject(data, theirs, partner, sync)
    expect(next.conflicts).toEqual([])
    const rows = next.outgoing.collections.find((c) => c.name === 'WorkItem')!.rows
    expect(rows.find((r) => r[0] === 'o_hm1')!.slice(2, 4)).toEqual(['Móng T1 (QLCL)', 'cái'])
    expect(rows.find((r) => r[0] === 'o_hm2')![2]).toBe('Thân T1 (đây)')
  })

  it('a new sync replaces the pending list; "theirs" brings back an object deleted meanwhile', async () => {
    let data = installPack(emptyProjectData('p1'), PACK, ui)
    const start = await roundTrip(projectSourceModel(project()))
    data = syncProject(data, start, partner, sync).data
    data = setValues(data, 'o_hm1', { Name: 'đây' }, ui)
    const theirs = edit(await roundTrip(start), 'WorkItem', 'o_hm1', 'Tên hạng mục', 'QLCL')
    data = syncProject(data, theirs, partner, sync).data
    const conflict = pendingConflicts(data, partner)[0]!
    // the conflicting object is deleted here before the user decides
    const gone = deleteObject(data, 'o_hm1', ui)
    const back = resolveSyncConflict(gone, partner, conflict, 'theirs', ui)
    expect(back.objects.find((o) => o.id === 'o_hm1')!.values).toMatchObject({
      Name: 'QLCL',
      Code: 'HM-01',
    })
    // agreeing by hand, then syncing again: the stale conflict is gone
    data = setValues(data, 'o_hm1', { Name: 'QLCL' }, ui)
    const again = syncProject(data, theirs, partner, sync)
    expect(again.conflicts).toEqual([])
    expect(pendingConflicts(again.data, partner)).toEqual([])
    expect(again.data.sync[partner]).not.toHaveProperty('conflicts')
  })
})
