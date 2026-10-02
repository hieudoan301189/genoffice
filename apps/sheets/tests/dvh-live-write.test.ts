/**
 * P3 write-back, Sheets side: a linked document's write sets the bound cell
 * (Data.SetField, recorded as a link change), bumps the field revision, and
 * reports a conflict when the cell moved away from the document's base.
 */
import { describe, expect, it } from 'vitest'
import { serializeModelXml, type DvhModel } from '@genoffice/dvh-model'
import { handleSetFields } from '../src/renderer/dvh-live'
import { loadDvhSmartData, registerDvhSheetsContext } from '../src/renderer/dvh-smart-data'
import type { LazyWorkbookState, UniverRuntime } from '../src/renderer/univer-state'

const DOC = 'doc_workbook00000001'

function fakeRuntime(cells: Map<string, { value: unknown; formula?: string }>) {
  const names: { name: string; formula: string }[] = []
  const range = (key: string) => ({
    getValue: () => cells.get(key)?.value ?? '',
    getFormula: () => cells.get(key)?.formula ?? '',
    setValue: (value: unknown) => void cells.set(key, { value }),
  })
  const workbook = {
    getSheetByName: (sheet: string) => ({ getRange: (cell: string) => range(`${sheet}!${cell}`) }),
    getDefinedNames: () =>
      names.map((n) => ({ getName: () => n.name, getFormulaOrRefString: () => n.formula })),
    newDefinedNameBuilder: () => ({
      load: (p: Record<string, unknown>) => ({
        build: () => ({ name: p.name as string, formula: p.formulaOrRefString as string }),
      }),
    }),
    insertDefinedNameBuilder: (n: { name: string; formula: string }) => void names.push(n),
  }
  return { univerAPI: { getActiveWorkbook: () => workbook } } as unknown as UniverRuntime
}

const model: DvhModel = {
  docId: DOC,
  schemaVersion: 1,
  fields: [
    { id: 'f_name', name: 'Project.Name', type: 'text', value: 'A', access: 'readwrite', rev: 2 },
    { id: 'f_sum', name: 'Total', type: 'number', value: 10, access: 'readwrite' },
  ],
  collections: [],
  tables: [],
  links: [],
}

async function setup(nameCell: string) {
  const cells = new Map<string, { value: unknown; formula?: string }>([
    ['Info!B1', { value: nameCell }],
    ['Info!B3', { value: 10, formula: '=B2*2' }],
  ])
  const runtime = fakeRuntime(cells)
  const state = {
    file: { sessionId: 's', path: 'C:\\data\\nguon.xlsx' },
  } as unknown as LazyWorkbookState
  const dvh = await loadDvhSmartData(
    runtime,
    state,
    async () =>
      ({
        names: [
          { name: '_dvh.f.f_name', formula: 'Info!$B$1' },
          { name: '_dvh.f.f_sum', formula: 'Info!$B$3' },
        ],
        model: { xml: serializeModelXml(model), storeItemId: '{A}' },
        history: null,
      }) as never,
  )
  let pending = 0
  registerDvhSheetsContext({
    getRuntime: () => runtime,
    getState: () => state,
    markPending: () => void pending++,
    notify: () => {},
  })
  return { cells, runtime, dvh, pending: () => pending }
}

const request = (value: string, expected: { rev: number; text: string } | null, force = false) => ({
  requestId: 'r1',
  docId: DOC,
  writes: [{ fieldId: 'f_name', value, expected, force }],
  origin: { docId: 'doc_document00000001' },
})

describe('handleSetFields', () => {
  it('writes the bound cell, bumps rev, records a link change set', async () => {
    const s = await setup('A')
    const reply = handleSetFields(s.runtime, request('B', { rev: 2, text: 'A' }))!
    expect(reply).toMatchObject({ handled: true, path: 'C:\\data\\nguon.xlsx' })
    expect(reply.results).toEqual([
      { fieldId: 'f_name', status: 'written', current: { rev: 3, text: 'B' } },
    ])
    expect(s.cells.get('Info!B1')!.value).toBe('B')
    expect(s.dvh.pending.at(-1)).toMatchObject({ action: 'Data.SetField', source: 'link' })
    expect(s.pending()).toBe(1)
  })

  it('reports a conflict when the cell was edited (unsaved) since the base; force writes', async () => {
    const s = await setup('typed in Sheets')
    const reply = handleSetFields(s.runtime, request('B', { rev: 2, text: 'A' }))!
    expect(reply.results[0]).toEqual({
      fieldId: 'f_name',
      status: 'conflict',
      current: { rev: 2, text: 'typed in Sheets' },
    })
    expect(s.cells.get('Info!B1')!.value).toBe('typed in Sheets')
    const forced = handleSetFields(s.runtime, request('B', { rev: 2, text: 'A' }, true))!
    expect(forced.results[0]!.status).toBe('written')
    expect(s.cells.get('Info!B1')!.value).toBe('B')
  })

  it('leaves formulas alone and declines other workbooks', async () => {
    const s = await setup('A')
    const formula = handleSetFields(s.runtime, {
      ...request('x', null),
      writes: [{ fieldId: 'f_sum', value: 5, expected: null, force: false }],
    })!
    expect(formula.results[0]!.status).toBe('formula')
    expect(handleSetFields(s.runtime, { ...request('x', null), docId: 'doc_other' })).toMatchObject(
      {
        handled: false,
      },
    )
    expect(handleSetFields(s.runtime, { nonsense: true })).toBeNull()
  })
})
