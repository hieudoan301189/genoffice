/**
 * P5 history model: base + snapshots survive the part round-trip; snapshots
 * every SNAPSHOT_EVERY change sets; compaction keeps the net effect; a model
 * changed outside DVH yields one `external` change set; a transaction revert
 * reports later edits as conflicts; size is measured against the P0 limit.
 */
import { describe, expect, it } from 'vitest'
import {
  cellDiff,
  compactHistory,
  diffModels,
  emptyModel,
  externalChangeSet,
  historyForSave,
  historyFrom,
  historySizeBytes,
  HISTORY_SIZE_LIMIT,
  objectHistory,
  parseHistoryXml,
  revertPlan,
  serializeHistoryXml,
  SNAPSHOT_EVERY,
  valueAfter,
  type ChangeSet,
  type DvhModel,
} from '../src/index'

const model = (name: string, rows: (string | number)[][] = [['A', 1]]): DvhModel => ({
  ...emptyModel('doc_aaaaaaaaaaaaaaaa'),
  fields: [{ id: 'f_name', name: 'Project.Name', type: 'text', value: name, access: 'readwrite' }],
  collections: [
    {
      id: 'c_items',
      name: 'Items',
      columns: [
        { id: 'k1', key: 'code', title: 'Mã', type: 'text' },
        { id: 'k2', key: 'qty', title: 'KL', type: 'number' },
      ],
      rows,
    },
  ],
})

let n = 0
const cs = (
  objectId: string,
  before: unknown,
  after: unknown,
  txId = `tx_${n}`,
  path = 'value',
): ChangeSet => ({
  id: `cs_${n++}`,
  txId,
  docId: 'doc_aaaaaaaaaaaaaaaa',
  at: `2026-10-02T00:${String(n).padStart(2, '0')}:00.000Z`,
  source: 'ui',
  action: 'Data.SetField',
  changes: [{ objectId, path, before, after }],
})

describe('history part (P5)', () => {
  it('keeps base and snapshots through serialize/parse', () => {
    const saved = historyForSave(null, [cs('f_name', 'A', 'B')], model('B'))
    const parsed = parseHistoryXml(serializeHistoryXml(saved))
    expect(parsed.base).toEqual(model('B'))
    expect(parsed.changes).toHaveLength(1)
    expect(parsed.modelHash).toBe(saved.modelHash)
  })

  it('takes a snapshot every SNAPSHOT_EVERY change sets and compacts behind the anchors', () => {
    let history = historyForSave(null, [], model('v0'))
    for (let i = 0; i < SNAPSHOT_EVERY * 3; i++) {
      history = historyForSave(history, [cs('f_name', `v${i}`, `v${i + 1}`)], model(`v${i + 1}`))
    }
    expect(history.snapshots).toHaveLength(3)
    const compacted = compactHistory(history)
    expect(compacted.snapshots).toHaveLength(2)
    expect(compacted.changes[0]).toMatchObject({
      action: 'History.Compact',
      changes: [{ objectId: 'f_name', before: 'v0', after: `v${SNAPSHOT_EVERY * 2}` }],
    })
    // the change sets after the older of the two kept snapshots stay as they were
    expect(compacted.changes).toHaveLength(1 + SNAPSHOT_EVERY)
    expect(historySizeBytes(compacted)).toBeLessThan(historySizeBytes(history))
    expect(historySizeBytes(history)).toBeLessThan(HISTORY_SIZE_LIMIT)
  })

  it('records a Word/Excel edit as one external change set', () => {
    const saved = historyForSave(null, [], model('A'))
    expect(externalChangeSet(saved, model('A'))).toBeNull()
    const ext = externalChangeSet(saved, model('Sửa trong Word', [['A', 2]]))!
    expect(ext.source).toBe('external')
    expect(ext.changes.map((c) => `${c.objectId}:${c.path}`)).toEqual([
      'f_name:value',
      'c_items:data',
    ])
    // no base (file from before P5): one model-level entry
    const legacy = externalChangeSet({ ...saved, base: undefined }, model('X'))!
    expect(legacy.changes[0]!.path).toBe('model')
  })

  it('diffs collections and tables; cell diffs for the panel', () => {
    const changes = diffModels(
      model('A', [['A', 1]]),
      model('A', [
        ['A', 5],
        ['B', 2],
      ]),
    )
    expect(changes).toHaveLength(1)
    const data = changes[0]!
    const diff = cellDiff(data.before as never, data.after as never)
    expect(diff.rowsBefore).toBe(1)
    expect(diff.rowsAfter).toBe(2)
    expect(diff.cells).toEqual([
      { row: 1, column: 'Mã', before: null, after: 'B' },
      { row: 0, column: 'KL', before: 1, after: 5 },
      { row: 1, column: 'KL', before: null, after: 2 },
    ])
  })

  it('reverts a transaction and reports later edits of the same object', () => {
    const changes = [
      cs('f_name', 'A', 'B', 'tx_ai'),
      cs('f_other', 1, 2, 'tx_ai'),
      cs('f_other', 2, 3, 'tx_user'),
    ]
    const plan = revertPlan(changes, 'tx_ai')
    expect(plan.writes).toEqual([
      { objectId: 'f_name', path: 'value', value: 'A' },
      { objectId: 'f_other', path: 'value', value: 1 },
    ])
    expect(plan.conflicts).toEqual([
      { objectId: 'f_other', path: 'value', changeSet: changes[2]!.id },
    ])
    expect(objectHistory(changes, 'f_other')).toHaveLength(2)
    expect(valueAfter(changes[0]!, 'f_name', 'value')).toEqual({ found: true, value: 'B' })
    expect(historyFrom({ docId: 'd', changes }, changes[2]!.at).changes).toHaveLength(1)
  })
})
