import type { zh } from './zh'

export const it = {
  dvhSmartData: 'Smart Data',
  dvhPanelTitle: 'Campi Smart Data',
  dvhColName: 'Campo',
  dvhColCell: 'Cella',
  dvhColValue: 'Valore',
  dvhNamePlaceholder: 'Nome del campo, es. Project.Name',
  dvhBindCell: 'Collega cella selezionata',
  dvhRemove: 'Rimuovi',
  dvhEmpty: 'Nessun campo. Seleziona una cella, assegnale un nome e collegala.',
  dvhBroken: 'Non collegato',
  dvhErrSelectCell: 'Seleziona prima una singola cella.',
  dvhErrName: 'Inserisci un nome di campo non ancora usato.',
  dvhErrNoFile: 'Apri prima una cartella di lavoro.',
  dvhUnboundToast: 'Il campo «{name}» ha perso la sua cella; collegalo di nuovo.',
  dvhClose: 'Chiudi',
  dvhBound: 'Campo «{name}» collegato a {cell}.',
} satisfies Record<keyof typeof zh, string>
