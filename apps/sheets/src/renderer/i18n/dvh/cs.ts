import type { zh } from './zh'

export const cs = {
  dvhSmartData: 'Smart Data',
  dvhPanelTitle: 'Pole Smart Data',
  dvhColName: 'Pole',
  dvhColCell: 'Buňka',
  dvhColValue: 'Hodnota',
  dvhNamePlaceholder: 'Název pole, např. Project.Name',
  dvhBindCell: 'Svázat vybranou buňku',
  dvhRemove: 'Odebrat',
  dvhEmpty: 'Zatím žádná pole. Vyberte buňku, pojmenujte ji a svažte.',
  dvhBroken: 'Nesvázáno',
  dvhErrSelectCell: 'Nejprve vyberte jednu buňku.',
  dvhErrName: 'Zadejte nepoužitý název pole.',
  dvhErrNoFile: 'Nejprve otevřete sešit.',
  dvhUnboundToast: 'Pole „{name}“ ztratilo buňku; svažte ho znovu.',
  dvhClose: 'Zavřít',
  dvhBound: 'Pole „{name}“ svázáno s {cell}.',
} satisfies Record<keyof typeof zh, string>
