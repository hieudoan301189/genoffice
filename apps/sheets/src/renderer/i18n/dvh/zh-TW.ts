import type { zh } from './zh'

export const zhTW = {
  dvhSmartData: '智慧資料',
  dvhPanelTitle: '智慧資料欄位',
  dvhColName: '欄位',
  dvhColCell: '儲存格',
  dvhColValue: '值',
  dvhNamePlaceholder: '欄位名稱，例如 Project.Name',
  dvhBindCell: '繫結所選儲存格',
  dvhRemove: '移除',
  dvhEmpty: '尚無欄位。選取一個儲存格，命名後繫結。',
  dvhBroken: '未繫結',
  dvhErrSelectCell: '請先選取一個儲存格。',
  dvhErrName: '請輸入尚未使用的欄位名稱。',
  dvhErrNoFile: '請先開啟活頁簿。',
  dvhUnboundToast: '欄位「{name}」的儲存格已刪除，請重新繫結。',
  dvhClose: '關閉',
  dvhBound: '欄位「{name}」已繫結到 {cell}。',
} satisfies Record<keyof typeof zh, string>
