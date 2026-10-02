import { useEffect, useState } from 'react'
import { DEFAULT_DVH_HOME, DVH_HOME_KEY, loadDvhHome } from './dvh-home-presets'
import { DvhSmartDataPanel } from './DvhSmartDataPanel'
import { useI18n } from './i18n/locale'
import type { RefKind, RoundFn } from './dvh-tool-actions'
import './dvh-home.css'

export type DvhToolDialog = 'settings' | 'round' | 'ref' | 'read-number' | 'smart-data'

/** Opens a DVH Tool dialog (the Home group and the DVH tab share them). */
export function openDvhToolDialog(kind: DvhToolDialog): void {
  window.dispatchEvent(new CustomEvent<DvhToolDialog>('dvh-tool-dialog', { detail: kind }))
}

const EDGE_NAMES = [
  'Trên',
  'Dưới',
  'Trái',
  'Phải',
  'Dọc trong',
  'Ngang trong',
  'Chéo xuống',
  'Chéo lên',
]
const STYLES = [
  [0, 'Không viền'],
  [3, 'Chấm mảnh'],
  [6, 'Gạch chấm chấm mảnh'],
  [5, 'Gạch chấm mảnh'],
  [4, 'Nét đứt mảnh'],
  [1, 'Liền mảnh'],
  [11, 'Gạch chấm chấm vừa'],
  [10, 'Gạch chấm vừa'],
  [9, 'Nét đứt vừa'],
  [8, 'Liền vừa'],
  [13, 'Liền đậm'],
  [7, 'Đôi'],
] as const
/**
 * The dialogs of the DVH Tool commands, mounted once by the shell: border and
 * indent presets, ROUND options, reference style, Read number, and the Smart
 * Data panel.
 */
export function DvhToolDialogs({ onCommand }: { onCommand: (command: string) => void }) {
  const { t } = useI18n()
  const [dialog, setDialog] = useState<DvhToolDialog | null>(null)
  const [smartDataOpen, setSmartDataOpen] = useState(false)
  const [settings, setSettings] = useState(loadDvhHome)
  const [error, setError] = useState('')
  const [roundFn, setRoundFn] = useState<RoundFn>('ROUND')
  const [digits, setDigits] = useState('2')
  const [ref, setRef] = useState<RefKind>('absolute')
  const [reading, setReading] = useState({
    input: '',
    output: '',
    before: '',
    after: '',
    intUnit: 'đồng',
    decUnit: '',
  })
  const indent = settings.indent[settings.indentIndex]!
  const border = settings.border[settings.borderIndex]!
  const command = (action: string, payload?: unknown) =>
    onCommand(`dvh:${JSON.stringify({ action, payload })}`)

  useEffect(() => {
    const open = (event: Event) => {
      const kind = (event as CustomEvent<DvhToolDialog>).detail
      if (kind === 'smart-data') {
        setSmartDataOpen((value) => !value)
        return
      }
      if (kind === 'settings') {
        setSettings(loadDvhHome())
        setError('')
      }
      setDialog(kind)
    }
    window.addEventListener('dvh-tool-dialog', open)
    return () => window.removeEventListener('dvh-tool-dialog', open)
  }, [])

  const small = (title: string, body: React.ReactNode, apply: () => void) => (
    <div
      className="dvh-home-backdrop"
      onKeyDown={(e) => {
        if (e.key === 'Escape') setDialog(null)
      }}
    >
      <section
        className="dvh-home-dialog narrow"
        role="dialog"
        aria-modal="true"
        aria-label={title}
      >
        <h2>{title}</h2>
        {body}
        <footer>
          <button onClick={() => setDialog(null)}>{t('dvhToolCancel')}</button>
          <button
            className="primary"
            onClick={() => {
              apply()
              setDialog(null)
            }}
          >
            {t('dvhToolApply')}
          </button>
        </footer>
      </section>
    </div>
  )

  return (
    <>
      {dialog === 'settings' && (
        <div
          className="dvh-home-backdrop"
          onKeyDown={(e) => {
            if (e.key === 'Escape') setDialog(null)
          }}
        >
          <section
            className="dvh-home-dialog"
            role="dialog"
            aria-modal="true"
            aria-label="Kiểu lề và đường viền DVH"
          >
            <h2>Kiểu lề và đường viền DVH</h2>
            <p>
              Chọn mẫu mặc định cho các nút Home. Lề tự tạo thêm tiền tố vào nội dung ô, bỏ qua công
              thức.
            </p>
            <div className="dvh-home-columns">
              <fieldset>
                <legend>Lề văn bản</legend>
                <select
                  aria-label="Mẫu lề"
                  value={settings.indentIndex}
                  onChange={(e) =>
                    setSettings({ ...settings, indentIndex: Number(e.target.value) })
                  }
                >
                  {settings.indent.map((p, i) => (
                    <option key={i} value={i}>
                      {p.name}
                    </option>
                  ))}
                </select>
                {(['name', 'first', 'tab', 'end'] as const).map((key, i) => (
                  <label key={key}>
                    {['Tên mẫu', 'Trước lề', 'Đơn vị lề (có thể nhập dấu cách)', 'Sau lề'][i]}
                    <input
                      value={indent[key]}
                      maxLength={80}
                      onChange={(e) =>
                        setSettings({
                          ...settings,
                          indent: settings.indent.map((p, j) =>
                            j === settings.indentIndex ? { ...p, [key]: e.target.value } : p,
                          ),
                        })
                      }
                    />
                  </label>
                ))}
                <pre>
                  {[1, 2, 3]
                    .map(
                      (i) =>
                        indent.first + indent.tab.repeat(i) + indent.end + ' Nội dung cấp ' + i,
                    )
                    .join('\n')}
                </pre>
                <button
                  onClick={() =>
                    setSettings({
                      ...settings,
                      indent: [...settings.indent, { ...indent, name: 'Mẫu lề mới' }],
                      indentIndex: settings.indent.length,
                    })
                  }
                >
                  Thêm mẫu lề
                </button>
              </fieldset>
              <fieldset>
                <legend>Đường viền</legend>
                <select
                  aria-label="Mẫu đường viền"
                  value={settings.borderIndex}
                  onChange={(e) =>
                    setSettings({ ...settings, borderIndex: Number(e.target.value) })
                  }
                >
                  {settings.border.map((p, i) => (
                    <option key={i} value={i}>
                      {p.name}
                    </option>
                  ))}
                </select>
                <label>
                  Tên mẫu
                  <input
                    value={border.name}
                    maxLength={80}
                    onChange={(e) =>
                      setSettings({
                        ...settings,
                        border: settings.border.map((p, i) =>
                          i === settings.borderIndex ? { ...p, name: e.target.value } : p,
                        ),
                      })
                    }
                  />
                </label>
                {border.edges.map((edge, i) => (
                  <div className="dvh-border-edge" key={i}>
                    <span>{EDGE_NAMES[i]}</span>
                    <select
                      aria-label={`Kiểu viền ${EDGE_NAMES[i]}`}
                      value={edge.style}
                      onChange={(e) =>
                        setSettings({
                          ...settings,
                          border: settings.border.map((p, j) =>
                            j === settings.borderIndex
                              ? {
                                  ...p,
                                  edges: p.edges.map((v, k) =>
                                    k === i ? { ...v, style: Number(e.target.value) } : v,
                                  ),
                                }
                              : p,
                          ),
                        })
                      }
                    >
                      {STYLES.map(([v, label]) => (
                        <option key={v} value={v}>
                          {label}
                        </option>
                      ))}
                    </select>
                    <input
                      aria-label={`Màu viền ${EDGE_NAMES[i]}`}
                      type="color"
                      value={edge.color}
                      onChange={(e) =>
                        setSettings({
                          ...settings,
                          border: settings.border.map((p, j) =>
                            j === settings.borderIndex
                              ? {
                                  ...p,
                                  edges: p.edges.map((v, k) =>
                                    k === i ? { ...v, color: e.target.value } : v,
                                  ),
                                }
                              : p,
                          ),
                        })
                      }
                    />
                  </div>
                ))}
                <button
                  onClick={() =>
                    setSettings({
                      ...settings,
                      border: [
                        ...settings.border,
                        {
                          ...border,
                          name: 'Mẫu viền mới',
                          edges: border.edges.map((e) => ({ ...e })),
                        },
                      ],
                      borderIndex: settings.border.length,
                    })
                  }
                >
                  Thêm mẫu viền
                </button>
              </fieldset>
            </div>
            {error && <p role="alert">{error}</p>}
            <footer>
              <button onClick={() => setSettings(structuredClone(DEFAULT_DVH_HOME))}>
                Khôi phục mẫu DVH
              </button>
              <button
                onClick={() => {
                  setSettings(loadDvhHome())
                  setDialog(null)
                }}
              >
                Hủy
              </button>
              <button
                onClick={() => {
                  if (!indent.tab && !indent.first && !indent.end) {
                    setError('Hãy nhập ít nhất một thành phần lề.')
                    return
                  }
                  try {
                    localStorage.setItem(DVH_HOME_KEY, JSON.stringify(settings))
                    setDialog(null)
                  } catch {
                    setError('Không lưu được cài đặt. Vui lòng thử lại.')
                  }
                }}
              >
                Lưu mặc định
              </button>
            </footer>
          </section>
        </div>
      )}
      {dialog === 'round' &&
        small(
          t('dvhToolRoundAdd'),
          <>
            <label>
              {t('dvhToolRoundFunction')}
              <select value={roundFn} onChange={(e) => setRoundFn(e.target.value as RoundFn)}>
                <option value="ROUND">ROUND</option>
                <option value="ROUNDUP">ROUNDUP</option>
                <option value="ROUNDDOWN">ROUNDDOWN</option>
              </select>
            </label>
            <label>
              {t('dvhToolRoundDigits')}
              <input
                type="number"
                min={-15}
                max={15}
                value={digits}
                onChange={(e) => setDigits(e.target.value)}
              />
            </label>
          </>,
          () =>
            command('round-add', {
              fn: roundFn,
              digits: Math.max(-15, Math.min(15, Math.trunc(Number(digits) || 0))),
            }),
        )}
      {dialog === 'ref' &&
        small(
          t('dvhToolRefStyle'),
          <>
            <p>{t('dvhToolRefStyleTip')}</p>
            {(
              [
                ['absolute', '$A$1'],
                ['absRow', 'A$1'],
                ['absCol', '$A1'],
                ['relative', 'A1'],
              ] as const
            ).map(([kind, sample]) => (
              <label key={kind} className="inline">
                <input
                  type="radio"
                  name="dvh-ref"
                  checked={ref === kind}
                  onChange={() => setRef(kind)}
                />
                <code>{sample}</code>
              </label>
            ))}
          </>,
          () => command('ref-style', { ref }),
        )}
      {dialog === 'read-number' &&
        small(
          t('dvhToolReadNumber'),
          <>
            <p>{t('dvhToolReadNumberHelp')}</p>
            {(
              [
                ['input', 'dvhToolReadInput'],
                ['output', 'dvhToolReadOutput'],
                ['before', 'dvhToolReadBefore'],
                ['after', 'dvhToolReadAfter'],
                ['intUnit', 'dvhToolReadIntUnit'],
                ['decUnit', 'dvhToolReadDecUnit'],
              ] as const
            ).map(([key, label]) => (
              <label key={key}>
                {t(label)}
                <input
                  value={reading[key]}
                  maxLength={120}
                  onChange={(e) => setReading({ ...reading, [key]: e.target.value })}
                />
              </label>
            ))}
          </>,
          () => command('read-number', reading),
        )}
      {smartDataOpen && <DvhSmartDataPanel onClose={() => setSmartDataOpen(false)} />}
    </>
  )
}
