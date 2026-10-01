import { describe, expect, it } from 'vitest'
import { changeTextIndent, DEFAULT_DVH_HOME } from '../src/renderer/dvh-home-presets'
import {
  createEditJournal,
  recordSetRangeValues,
  toSaveEdits,
  toNeutralStyle,
} from '../src/renderer/edit-journal'
import { StylesheetEditor } from '@genoffice/xlsx-gateway/gateway/xlsx-styles'
import { workbookStyleEditSchema } from '@genoffice/xlsx-gateway/shared/edit-schemas'

describe('DVH Home defaults', () => {
  it('ports all saved indent templates and removes the longest prefix first', () => {
    for (const preset of DEFAULT_DVH_HOME.indent) {
      let text = 'Công tác bê tông'
      for (let i = 1; i <= 10; i++) text = changeTextIndent(text, preset, 'add')
      expect(text).toBe(preset.first + preset.tab.repeat(10) + preset.end + 'Công tác bê tông')
      expect(changeTextIndent(text, preset, 'add')).toBe(text)
      expect(changeTextIndent(text, preset, 'clear')).toBe('Công tác bê tông')
      expect(changeTextIndent(changeTextIndent('Nội dung', preset, 'add'), preset, 'remove')).toBe(
        'Nội dung',
      )
    }
  })
  it('journals center-across and shrink without replacing content or losing original font size', () => {
    const journal = createEditJournal()
    recordSetRangeValues(journal, 's1', {
      0: {
        0: {
          s: { ht: 2, fs: 6, tb: 2 },
          custom: { dvhCenterAcross: true, dvhShrink: true, dvhOriginalSize: 14 },
        },
      },
    })
    const edits = toSaveEdits(journal)
    expect(edits[0]?.style).toMatchObject({
      horizontalAlignment: 'centerContinuous',
      shrinkToFit: true,
      fontSize: 14,
    })
    expect(workbookStyleEditSchema.safeParse(edits[0]?.style).success).toBe(true)
  })
  it('writes native XLSX alignment and diagonal borders', () => {
    const xml =
      '<styleSheet><fonts count="1"><font><sz val="11"/></font></fonts><fills count="1"><fill/></fills><borders count="1"><border/></borders><cellXfs count="1"><xf fontId="0" fillId="0" borderId="0"/></cellXfs></styleSheet>'
    const editor = new StylesheetEditor(xml)
    editor.resolveStyle(0, {
      horizontalAlignment: 'centerContinuous',
      shrinkToFit: true,
      borderDiagonal: { style: 'thin', color: '#FF0000' },
      diagonalDown: true,
      diagonalUp: false,
    })
    const result = editor.serialize()
    expect(result).toContain('horizontal="centerContinuous"')
    expect(result).toContain('shrinkToFit="1"')
    expect(result).toContain('diagonalDown="1"')
    expect(result).toContain('<diagonal style="thin">')
  })
  it('records removed diagonal edges and default alignment', () => {
    expect(toNeutralStyle({ ht: 0, bd: { tl_br: null, bl_tr: null } })).toMatchObject({
      horizontalAlignment: 'general',
      borderDiagonal: null,
      diagonalUp: false,
      diagonalDown: false,
    })
  })
})
