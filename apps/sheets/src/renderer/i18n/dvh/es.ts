import type { zh } from './zh'

export const es = {
  dvhSmartData: 'Smart Data',
  dvhPanelTitle: 'Campos de Smart Data',
  dvhColName: 'Campo',
  dvhColCell: 'Celda',
  dvhColValue: 'Valor',
  dvhNamePlaceholder: 'Nombre del campo, p. ej. Project.Name',
  dvhBindCell: 'Vincular celda seleccionada',
  dvhRemove: 'Quitar',
  dvhEmpty: 'Aún no hay campos. Seleccione una celda, asígnele un nombre y vincúlela.',
  dvhBroken: 'Sin vincular',
  dvhErrSelectCell: 'Seleccione primero una sola celda.',
  dvhErrName: 'Escriba un nombre de campo que no esté en uso.',
  dvhErrNoFile: 'Abra primero un libro.',
  dvhUnboundToast: 'El campo «{name}» perdió su celda; vincúlelo de nuevo.',
  dvhClose: 'Cerrar',
  dvhBound: 'Campo «{name}» vinculado a {cell}.',
} satisfies Record<keyof typeof zh, string>
