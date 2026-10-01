import type { zh } from './zh'

export const pl = {
  dvhSmartData: 'Smart Data',
  dvhPanelTitle: 'Pola Smart Data',
  dvhColName: 'Pole',
  dvhColCell: 'Komórka',
  dvhColValue: 'Wartość',
  dvhNamePlaceholder: 'Nazwa pola, np. Project.Name',
  dvhBindCell: 'Powiąż zaznaczoną komórkę',
  dvhRemove: 'Usuń',
  dvhEmpty: 'Brak pól. Zaznacz komórkę, nadaj nazwę i powiąż.',
  dvhBroken: 'Niepowiązane',
  dvhErrSelectCell: 'Najpierw zaznacz jedną komórkę.',
  dvhErrName: 'Wpisz nieużywaną nazwę pola.',
  dvhErrNoFile: 'Najpierw otwórz skoroszyt.',
  dvhUnboundToast: 'Pole „{name}” utraciło komórkę; powiąż je ponownie.',
  dvhClose: 'Zamknij',
  dvhBound: 'Pole „{name}” powiązano z {cell}.',
} satisfies Record<keyof typeof zh, string>
