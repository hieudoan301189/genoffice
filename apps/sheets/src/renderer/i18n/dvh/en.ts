import type { zh } from './zh'

export const en = {
  dvhSmartData: 'Smart Data',
  dvhPanelTitle: 'Smart Data fields',
  dvhColName: 'Field',
  dvhColCell: 'Cell',
  dvhColValue: 'Value',
  dvhNamePlaceholder: 'Field name, e.g. Project.Name',
  dvhBindCell: 'Bind selected cell',
  dvhRemove: 'Remove',
  dvhEmpty: 'No fields yet. Select a cell, name it, and bind it.',
  dvhBroken: 'Unbound',
  dvhErrSelectCell: 'Select a single cell first.',
  dvhErrName: 'Enter a field name that is not used yet.',
  dvhErrNoFile: 'Open a workbook first.',
  dvhUnboundToast: 'Field "{name}" lost its cell; bind it again.',
  dvhClose: 'Close',
  dvhBound: 'Field "{name}" bound to {cell}.',
} satisfies Record<keyof typeof zh, string>
