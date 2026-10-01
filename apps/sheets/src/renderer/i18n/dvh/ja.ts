import type { zh } from './zh'

export const ja = {
  dvhSmartData: 'スマートデータ',
  dvhPanelTitle: 'スマートデータ フィールド',
  dvhColName: 'フィールド',
  dvhColCell: 'セル',
  dvhColValue: '値',
  dvhNamePlaceholder: 'フィールド名（例: Project.Name）',
  dvhBindCell: '選択したセルをバインド',
  dvhRemove: '削除',
  dvhEmpty: 'フィールドはまだありません。セルを選択し、名前を付けてバインドします。',
  dvhBroken: '未バインド',
  dvhErrSelectCell: '先にセルを 1 つ選択してください。',
  dvhErrName: '未使用のフィールド名を入力してください。',
  dvhErrNoFile: '先にブックを開いてください。',
  dvhUnboundToast: 'フィールド「{name}」のセルが削除されました。再度バインドしてください。',
  dvhClose: '閉じる',
  dvhBound: 'フィールド「{name}」を {cell} にバインドしました。',
} satisfies Record<keyof typeof zh, string>
