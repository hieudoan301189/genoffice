import { loadDvhHome } from './dvh-home-presets'
import { DVH_TOOL_ICONS } from './ribbon-icons'
import { useI18n, type StringKey } from './i18n/locale'
import { openDvhToolDialog } from './DvhToolDialogs'
import './dvh-home.css'

/** One DVH Tool button: its icon, label/tip strings and what it does. */
export interface DvhToolButton {
  readonly id: string
  readonly label: StringKey
  readonly tip: StringKey
  readonly run: (command: (action: string, payload?: unknown) => void) => void
}

/**
 * The DVH Tool commands in DVH-Tool's own order (Home group, columns of
 * three): text case, visibility, alignment and borders, indents, formulas,
 * visible-cells copy.
 */
export const DVH_TOOL_BUTTONS: readonly DvhToolButton[] = [
  {
    id: 'upper',
    label: 'dvhToolUpper',
    tip: 'dvhToolUpperTip',
    run: (c) => c('case', { kind: 'upper' }),
  },
  {
    id: 'lower',
    label: 'dvhToolLower',
    tip: 'dvhToolLowerTip',
    run: (c) => c('case', { kind: 'lower' }),
  },
  {
    id: 'proper',
    label: 'dvhToolProper',
    tip: 'dvhToolProperTip',
    run: (c) => c('case', { kind: 'proper' }),
  },
  {
    id: 'firstChar',
    label: 'dvhToolFirstChar',
    tip: 'dvhToolFirstCharTip',
    run: (c) => c('case', { kind: 'firstChar' }),
  },
  { id: 'hidden', label: 'dvhToolHidden', tip: 'dvhToolHiddenTip', run: (c) => c('hidden') },
  {
    id: 'zeroHide',
    label: 'dvhToolZeroHide',
    tip: 'dvhToolZeroHideTip',
    run: (c) => c('zero-hide'),
  },
  { id: 'centerAcross', label: 'dvhToolCenter', tip: 'dvhToolCenterTip', run: (c) => c('center') },
  { id: 'justify', label: 'dvhToolJustify', tip: 'dvhToolJustifyTip', run: (c) => c('justify') },
  {
    id: 'beautyBorder',
    label: 'dvhToolBorder',
    tip: 'dvhToolBorderTip',
    run: (c) => {
      const s = loadDvhHome()
      c('border', s.border[s.borderIndex])
    },
  },
  { id: 'fitMerge', label: 'dvhToolFitMerge', tip: 'dvhToolFitMergeTip', run: (c) => c('fit') },
  { id: 'shrink', label: 'dvhToolShrink', tip: 'dvhToolShrinkTip', run: (c) => c('shrink') },
  {
    id: 'borderSettings',
    label: 'dvhToolSettings',
    tip: 'dvhToolSettingsTip',
    run: () => openDvhToolDialog('settings'),
  },
  {
    id: 'indentIncrease',
    label: 'dvhToolIndentAdd',
    tip: 'dvhToolIndentAddTip',
    run: (c) => {
      const s = loadDvhHome()
      c('indent-add', s.indent[s.indentIndex])
    },
  },
  {
    id: 'indentDecrease',
    label: 'dvhToolIndentRemove',
    tip: 'dvhToolIndentRemoveTip',
    run: (c) => {
      const s = loadDvhHome()
      c('indent-remove', s.indent[s.indentIndex])
    },
  },
  {
    id: 'indentClear',
    label: 'dvhToolIndentClear',
    tip: 'dvhToolIndentClearTip',
    run: (c) => {
      const s = loadDvhHome()
      c('indent-clear', s.indent[s.indentIndex])
    },
  },
  {
    id: 'refStyle',
    label: 'dvhToolRefStyle',
    tip: 'dvhToolRefStyleTip',
    run: () => openDvhToolDialog('ref'),
  },
  {
    id: 'roundAdd',
    label: 'dvhToolRoundAdd',
    tip: 'dvhToolRoundAddTip',
    run: () => openDvhToolDialog('round'),
  },
  {
    id: 'roundRemove',
    label: 'dvhToolRoundRemove',
    tip: 'dvhToolRoundRemoveTip',
    run: (c) => c('round-remove'),
  },
  {
    id: 'copyVisible',
    label: 'dvhToolCopyVisible',
    tip: 'dvhToolCopyVisibleTip',
    run: (c) => c('copy-visible'),
  },
  {
    id: 'pasteVisible',
    label: 'dvhToolPasteVisible',
    tip: 'dvhToolPasteVisibleTip',
    run: (c) => c('paste-visible'),
  },
  {
    id: 'readNumber',
    label: 'dvhToolReadNumber',
    tip: 'dvhToolReadNumberTip',
    run: () => openDvhToolDialog('read-number'),
  },
]

/** The DVH Tool group of the Home tab: icon-only buttons, three per column, as in DVH-Tool. */
export function DvhHomeTools({ onCommand }: { onCommand: (command: string) => void }) {
  const { t } = useI18n()
  const command = (action: string, payload?: unknown) =>
    onCommand(`dvh:${JSON.stringify({ action, payload })}`)
  return (
    <div className="dvh-home-tools" role="group" aria-label="DVH Tool">
      <div className="dvh-home-buttons">
        {DVH_TOOL_BUTTONS.map((b) => (
          <button
            key={b.id}
            data-tip={t(b.label)}
            data-tip-detail={t(b.tip)}
            aria-label={t(b.label)}
            onClick={() => b.run(command)}
          >
            {DVH_TOOL_ICONS[b.id]}
          </button>
        ))}
      </div>
      <span className="dvh-home-caption">DVH Tool</span>
    </div>
  )
}
