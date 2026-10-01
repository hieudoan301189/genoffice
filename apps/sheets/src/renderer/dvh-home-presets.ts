// Ported from DVH-Excel/WPF/Window/LeVB/TabHandler.cs and Border/BorderManager.cs.
// Default library captured from this user's DVH-Tool settings on 2026-09-28.
export interface IndentPreset {
  name: string
  tab: string
  first: string
  end: string
}
export interface BorderEdge {
  style: number
  color: string
}
export interface BorderPreset {
  name: string
  edges: BorderEdge[]
}
export interface DvhHomeSettings {
  indent: IndentPreset[]
  border: BorderPreset[]
  indentIndex: number
  borderIndex: number
}
const edge = (style: number, color: string): BorderEdge => ({ style, color })
export const DEFAULT_DVH_HOME: DvhHomeSettings = {
  indent: [
    { name: '7 dấu cách +', tab: '       ', first: '', end: '+' },
    { name: '7 dấu cách −', tab: '       ', first: '', end: '-' },
    { name: 'Dấu sao +', tab: '*****', first: '', end: '+' },
  ],
  border: [
    {
      name: 'Viền cam, trong mảnh',
      edges: [13, 13, 13, 13, 1, 1, 0, 0].map((s) => edge(s, '#FF8000')),
    },
    {
      name: 'Viền đỏ, ngang nét đứt',
      edges: [13, 13, 13, 13, 1, 4, 0, 0].map((s, i) => edge(s, i < 4 ? '#FF0000' : '#000000')),
    },
  ],
  indentIndex: 0,
  borderIndex: 1,
}
export const DVH_HOME_KEY = 'dvh-office.home-presets.v1'
export function loadDvhHome(): DvhHomeSettings {
  try {
    const value = JSON.parse(localStorage.getItem(DVH_HOME_KEY) ?? 'null') as DvhHomeSettings | null
    if (
      value &&
      value.indent?.length &&
      value.border?.length &&
      value.indent.every((p) =>
        [p.name, p.tab, p.first, p.end].every((v) => typeof v === 'string'),
      ) &&
      value.border.every(
        (p) =>
          typeof p.name === 'string' &&
          p.edges?.length === 8 &&
          p.edges.every(
            (e) =>
              Number.isInteger(e.style) &&
              e.style >= 0 &&
              e.style <= 13 &&
              /^#[\da-f]{6}$/i.test(e.color),
          ),
      ) &&
      value.indent[value.indentIndex] &&
      value.border[value.borderIndex]
    )
      return value
  } catch {
    /* Fall back to the bundled library if a saved setting is invalid. */
  }
  return structuredClone(DEFAULT_DVH_HOME)
}
export function changeTextIndent(
  text: string,
  preset: IndentPreset,
  change: 'add' | 'remove' | 'clear',
): string {
  const levels = Array.from({ length: 11 }, (_, i) =>
    i ? preset.first + preset.tab.repeat(i) + preset.end : '',
  )
  let level = 0
  for (let i = 10; i > 0; i--)
    if (levels[i] && text.startsWith(levels[i]!)) {
      level = i
      break
    }
  const next =
    change === 'clear' ? 0 : Math.max(0, Math.min(10, level + (change === 'add' ? 1 : -1)))
  return levels[next]! + text.slice(levels[level]!.length)
}
