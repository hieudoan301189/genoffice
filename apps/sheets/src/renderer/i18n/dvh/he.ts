import type { zh } from './zh'

export const he = {
  dvhSmartData: 'נתונים חכמים',
  dvhPanelTitle: 'שדות נתונים חכמים',
  dvhColName: 'שדה',
  dvhColCell: 'תא',
  dvhColValue: 'ערך',
  dvhNamePlaceholder: 'שם שדה, למשל Project.Name',
  dvhBindCell: 'קשר את התא שנבחר',
  dvhRemove: 'הסר',
  dvhEmpty: 'אין עדיין שדות. בחר תא, תן לו שם וקשר אותו.',
  dvhBroken: 'לא מקושר',
  dvhErrSelectCell: 'בחר תחילה תא אחד.',
  dvhErrName: 'הזן שם שדה שעדיין לא בשימוש.',
  dvhErrNoFile: 'פתח תחילה חוברת עבודה.',
  dvhUnboundToast: 'השדה "{name}" איבד את התא שלו; קשר אותו מחדש.',
  dvhClose: 'סגור',
  dvhBound: 'השדה "{name}" קושר לתא {cell}.',
} satisfies Record<keyof typeof zh, string>
