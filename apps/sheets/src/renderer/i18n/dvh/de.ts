import type { zh } from './zh'

export const de = {
  dvhSmartData: 'Smart Data',
  dvhPanelTitle: 'Smart-Data-Felder',
  dvhColName: 'Feld',
  dvhColCell: 'Zelle',
  dvhColValue: 'Wert',
  dvhNamePlaceholder: 'Feldname, z. B. Project.Name',
  dvhBindCell: 'Ausgewählte Zelle binden',
  dvhRemove: 'Entfernen',
  dvhEmpty: 'Noch keine Felder. Zelle auswählen, benennen und binden.',
  dvhBroken: 'Nicht gebunden',
  dvhErrSelectCell: 'Wählen Sie zuerst eine einzelne Zelle aus.',
  dvhErrName: 'Geben Sie einen noch nicht verwendeten Feldnamen ein.',
  dvhErrNoFile: 'Öffnen Sie zuerst eine Arbeitsmappe.',
  dvhUnboundToast: 'Feld „{name}“ hat seine Zelle verloren; bitte neu binden.',
  dvhClose: 'Schließen',
  dvhBound: 'Feld „{name}“ an {cell} gebunden.',
} satisfies Record<keyof typeof zh, string>
