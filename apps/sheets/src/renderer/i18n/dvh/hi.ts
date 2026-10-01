import type { zh } from './zh'

export const hi = {
  dvhSmartData: 'स्मार्ट डेटा',
  dvhPanelTitle: 'स्मार्ट डेटा फ़ील्ड',
  dvhColName: 'फ़ील्ड',
  dvhColCell: 'सेल',
  dvhColValue: 'मान',
  dvhNamePlaceholder: 'फ़ील्ड नाम, जैसे Project.Name',
  dvhBindCell: 'चयनित सेल बाइंड करें',
  dvhRemove: 'हटाएँ',
  dvhEmpty: 'अभी कोई फ़ील्ड नहीं। सेल चुनें, नाम दें और बाइंड करें।',
  dvhBroken: 'बाइंड नहीं',
  dvhErrSelectCell: 'पहले एक सेल चुनें।',
  dvhErrName: 'ऐसा फ़ील्ड नाम दर्ज करें जो उपयोग में न हो।',
  dvhErrNoFile: 'पहले वर्कबुक खोलें।',
  dvhUnboundToast: 'फ़ील्ड "{name}" का सेल हट गया; उसे फिर से बाइंड करें।',
  dvhClose: 'बंद करें',
  dvhBound: 'फ़ील्ड "{name}" को {cell} से बाइंड किया गया।',
} satisfies Record<keyof typeof zh, string>
