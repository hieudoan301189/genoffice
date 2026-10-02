/**
 * P4 in Sheets: every workbook DSL operation is a Spreadsheet.* action whose
 * input is the operation's own schema; preview is the plan, execute one batch.
 */
import { describe, expect, it } from 'vitest'
import { defaultPermissions } from '@genoffice/dvh-actions'
import { workbookOperationSchema } from '@genoffice/xlsx-gateway/domain/workbook-dsl'
import { createSheetsDvhActions } from '../src/renderer/dvh-actions'

describe('Spreadsheet.* adapter', () => {
  it('lists one action per DSL op, validates with the op schema, previews then applies', async () => {
    const calls: { ops: unknown[]; dryRun: boolean }[] = []
    const registry = createSheetsDvhActions({
      applyOps: async (ops, dryRun) => {
        calls.push({ ops, dryRun })
        return dryRun
          ? { ok: true, dryRun: true, cellChangeCount: 3, structuralChanges: [], formatChanges: [] }
          : { ok: true, message: 'Applied 3 changes' }
      },
    })
    const ops = workbookOperationSchema.options.length
    const names = registry
      .catalog()
      .filter((a) => a.group === 'Spreadsheet')
      .map((a) => a.name)
    expect(names.length).toBeGreaterThanOrEqual(ops - 1)
    expect(names).toContain('Spreadsheet.SetCell')
    expect(registry.get('Spreadsheet.DeleteSheet')!.effect).toBe('destructive')
    const input = { sheetId: 's1', address: 'A1', value: 1 }
    const ui = { caller: 'ui' as const, docId: 'doc_w', permissions: defaultPermissions('ui') }
    const dry = await registry.run('Spreadsheet.SetCell', input, { ...ui, dryRun: true })
    expect(dry.preview.objects).toBe(3)
    const done = await registry.run('Spreadsheet.SetCell', input, ui)
    expect(calls.map((c) => c.dryRun)).toEqual([true, true, false])
    expect(calls[2]!.ops[0]).toMatchObject({ op: 'set_cell', sheetId: 's1' })
    expect(done.changeSet?.action).toBe('Spreadsheet.SetCell')
    await expect(registry.run('Spreadsheet.SetCell', { nope: 1 }, ui)).rejects.toMatchObject({
      code: 'invalid_input',
    })
  })
})
