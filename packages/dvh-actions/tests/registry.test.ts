import { describe, expect, it, vi } from 'vitest'
import { z } from 'zod'
import { emptyModel, type DvhModel } from '@genoffice/dvh-model'
import {
  ActionError,
  ActionRegistry,
  defaultPermissions,
  type ActionDescriptor,
} from '../src/index'

function setup() {
  const model: DvhModel = {
    ...emptyModel('doc_aaaaaaaaaaaaaaaa'),
    fields: [{ id: 'f_a', name: 'Project.Name', type: 'text', value: 'A', access: 'readwrite' }],
  }
  const registry = new ActionRegistry()
  const setField: ActionDescriptor<{ fieldId: string; value: string }, string> = {
    name: 'Data.SetField',
    group: 'Data',
    summary: 'Set a Smart Data field value',
    input: z.object({ fieldId: z.string(), value: z.string() }),
    effect: 'write',
    preview: ({ fieldId }) => ({
      objects: model.fields.some((f) => f.id === fieldId) ? 1 : 0,
      files: [],
      summary: [],
      warnings: [],
    }),
    execute: ({ fieldId, value }, ctx) => {
      const field = model.fields.find((f) => f.id === fieldId)!
      ctx.emit([{ objectId: fieldId, path: 'value', before: field.value, after: value }])
      field.value = value
      return value
    },
  }
  const execute = vi.fn()
  const clear: ActionDescriptor<Record<string, never>, void> = {
    name: 'Data.ClearAll',
    group: 'Data',
    summary: 'Clear every field',
    input: z.object({}),
    effect: 'destructive',
    preview: () => ({
      objects: model.fields.length,
      files: [],
      summary: [],
      warnings: ['clears every field'],
    }),
    execute,
  }
  registry.register(setField)
  registry.register(clear)
  return { registry, model, execute }
}

const options = (caller: 'ui' | 'ai' = 'ui') => ({
  caller,
  docId: 'doc_aaaaaaaaaaaaaaaa',
  permissions: defaultPermissions(caller),
  now: () => '2026-10-02T00:00:00.000Z',
})

describe('ActionRegistry', () => {
  it('runs an action and reports one change set for its transaction', async () => {
    const { registry, model } = setup()
    const result = await registry.run(
      'Data.SetField',
      { fieldId: 'f_a', value: 'B' },
      { ...options(), txId: 'tx_1' },
    )
    expect(result.output).toBe('B')
    expect(model.fields[0]!.value).toBe('B')
    expect(result.changeSet).toMatchObject({
      txId: 'tx_1',
      source: 'ui',
      action: 'Data.SetField',
      changes: [{ objectId: 'f_a', path: 'value', before: 'A', after: 'B' }],
    })
  })

  it('dry run previews without executing', async () => {
    const { registry, model } = setup()
    const result = await registry.run(
      'Data.SetField',
      { fieldId: 'f_a', value: 'B' },
      { ...options(), dryRun: true },
    )
    expect(result.preview.objects).toBe(1)
    expect(result.output).toBeUndefined()
    expect(model.fields[0]!.value).toBe('A')
  })

  it('refuses actions outside the catalog and malformed input (AI safety)', async () => {
    const { registry } = setup()
    await expect(registry.run('File.DeleteEverything', {}, options('ai'))).rejects.toMatchObject({
      code: 'unknown_action',
    })
    await expect(
      registry.run('Data.SetField', { fieldId: 1 }, options('ai')),
    ).rejects.toMatchObject({ code: 'invalid_input' })
  })

  it('denies effects the caller lacks: AI cannot run destructive actions by default', async () => {
    const { registry, execute } = setup()
    await expect(registry.run('Data.ClearAll', {}, options('ai'))).rejects.toBeInstanceOf(
      ActionError,
    )
    await expect(registry.run('Data.ClearAll', {}, options('ai'))).rejects.toMatchObject({
      code: 'permission_denied',
    })
    expect(execute).not.toHaveBeenCalled()
  })

  it('asks before destructive actions and refuses without a yes', async () => {
    const { registry, execute } = setup()
    await expect(registry.run('Data.ClearAll', {}, options())).rejects.toMatchObject({
      code: 'confirmation_required',
    })
    const confirm = vi.fn(() => true)
    await registry.run('Data.ClearAll', {}, { ...options(), confirm })
    expect(confirm).toHaveBeenCalledWith('Data.ClearAll', expect.objectContaining({ objects: 1 }))
    expect(execute).toHaveBeenCalledTimes(1)
  })

  it('publishes a catalog with JSON Schema inputs and rejects duplicate names', () => {
    const { registry } = setup()
    const catalog = registry.catalog()
    expect(catalog.map((a) => a.name)).toEqual(['Data.SetField', 'Data.ClearAll'])
    expect(catalog[0]!.input).toMatchObject({ type: 'object', required: ['fieldId', 'value'] })
    expect(() => registry.register(registry.list()[0]!)).toThrow(/already registered/)
  })
})
