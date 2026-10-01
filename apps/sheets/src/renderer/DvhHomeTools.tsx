import { useState } from 'react'
import { DEFAULT_DVH_HOME, DVH_HOME_KEY, loadDvhHome } from './dvh-home-presets'
import './dvh-home.css'

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
export function DvhHomeTools({ onCommand }: { onCommand: (command: string) => void }) {
  const [settings, setSettings] = useState(loadDvhHome)
  const [open, setOpen] = useState(false)
  const [error, setError] = useState('')
  const indent = settings.indent[settings.indentIndex]!
  const border = settings.border[settings.borderIndex]!
  const command = (action: string, payload?: unknown) =>
    onCommand(`dvh:${JSON.stringify({ action, payload })}`)
  return (
    <div className="dvh-home-tools">
      <div className="dvh-home-buttons">
        <button onClick={() => command('center')}>Căn giữa không gộp</button>
        <button onClick={() => command('justify')}>Căn đều hai bên</button>
        <button onClick={() => command('shrink')}>Thu chữ vừa ô</button>
        <button onClick={() => command('fit')}>Giãn dòng ô gộp</button>
        <button onClick={() => command('border', border)}>Đường viền đẹp</button>
        <button
          onClick={() => {
            setSettings(loadDvhHome())
            setError('')
            setOpen(true)
          }}
        >
          Kiểu lề và đường viền…
        </button>
        <button onClick={() => command('indent-add', indent)}>Tăng lề tự tạo</button>
        <button onClick={() => command('indent-remove', indent)}>Giảm lề tự tạo</button>
        <button onClick={() => command('indent-clear', indent)}>Xóa lề tự tạo</button>
      </div>
      <span className="dvh-home-caption">DVH Tool</span>
      {open && (
        <div
          className="dvh-home-backdrop"
          onKeyDown={(e) => {
            if (e.key === 'Escape') setOpen(false)
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
                  setOpen(false)
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
                    setOpen(false)
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
    </div>
  )
}
