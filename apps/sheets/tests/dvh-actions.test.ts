import { describe, expect, it } from 'vitest'
import { defaultPermissions } from '@genoffice/dvh-actions'
import { createSheetsDvhActions, parseCellRef } from '../src/renderer/dvh-actions'

describe('DVH actions (Sheets)', () => {
  it('registers the P1 actions with their effects', () => {
    expect(
      createSheetsDvhActions()
        .catalog()
        .map((a) => `${a.name}:${a.effect}`),
    ).toEqual([
      'Data.ListFields:read',
      'Data.GetField:read',
      'Data.SetField:write',
      'Spreadsheet.BindField:write',
    ])
  })

  it('resolves cell references against the active sheet', () => {
    expect(parseCellRef('B5', 'Sheet1')).toEqual({ sheetName: 'Sheet1', row: 4, column: 1 })
    expect(parseCellRef('AA10', 'Dữ liệu')).toEqual({ sheetName: 'Dữ liệu', row: 9, column: 26 })
    expect(parseCellRef("'Bảng kê'!$C$3", 'Sheet1')).toEqual({
      sheetName: 'Bảng kê',
      row: 2,
      column: 2,
    })
    expect(() => parseCellRef('B5:C6', 'Sheet1')).toThrow('not a single cell')
  })

  it('refuses to run without an open workbook', async () => {
    await expect(
      createSheetsDvhActions().run(
        'Data.ListFields',
        {},
        {
          caller: 'ai',
          docId: 'doc_x',
          permissions: defaultPermissions('ai'),
        },
      ),
    ).rejects.toThrow('no workbook is open')
  })
})
