import type { zh } from './zh'

export const th = {
  dvhSmartData: 'Smart Data',
  dvhPanelTitle: 'ฟิลด์ Smart Data',
  dvhColName: 'ฟิลด์',
  dvhColCell: 'เซลล์',
  dvhColValue: 'ค่า',
  dvhNamePlaceholder: 'ชื่อฟิลด์ เช่น Project.Name',
  dvhBindCell: 'ผูกเซลล์ที่เลือก',
  dvhRemove: 'ลบ',
  dvhEmpty: 'ยังไม่มีฟิลด์ เลือกเซลล์ ตั้งชื่อ แล้วผูก',
  dvhBroken: 'ยังไม่ได้ผูก',
  dvhErrSelectCell: 'โปรดเลือกเซลล์เดียวก่อน',
  dvhErrName: 'โปรดป้อนชื่อฟิลด์ที่ยังไม่ถูกใช้',
  dvhErrNoFile: 'โปรดเปิดเวิร์กบุ๊กก่อน',
  dvhUnboundToast: 'ฟิลด์ "{name}" สูญเสียเซลล์แล้ว โปรดผูกอีกครั้ง',
  dvhClose: 'ปิด',
  dvhBound: 'ผูกฟิลด์ "{name}" กับ {cell} แล้ว',
} satisfies Record<keyof typeof zh, string>
