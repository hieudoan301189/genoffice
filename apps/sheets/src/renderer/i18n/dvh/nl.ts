import type { zh } from './zh'

export const nl = {
  dvhSmartData: 'Smart Data',
  dvhPanelTitle: 'Smart Data-velden',
  dvhColName: 'Veld',
  dvhColCell: 'Cel',
  dvhColValue: 'Waarde',
  dvhNamePlaceholder: 'Veldnaam, bijv. Project.Name',
  dvhBindCell: 'Geselecteerde cel koppelen',
  dvhRemove: 'Verwijderen',
  dvhEmpty: 'Nog geen velden. Selecteer een cel, geef een naam en koppel.',
  dvhBroken: 'Niet gekoppeld',
  dvhErrSelectCell: 'Selecteer eerst één cel.',
  dvhErrName: 'Voer een veldnaam in die nog niet wordt gebruikt.',
  dvhErrNoFile: 'Open eerst een werkmap.',
  dvhUnboundToast: 'Veld "{name}" heeft zijn cel verloren; koppel het opnieuw.',
  dvhClose: 'Sluiten',
  dvhBound: 'Veld "{name}" gekoppeld aan {cell}.',
} satisfies Record<keyof typeof zh, string>
