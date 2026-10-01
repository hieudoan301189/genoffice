import type { zh } from './zh'

export const pt = {
  dvhSmartData: 'Smart Data',
  dvhPanelTitle: 'Campos do Smart Data',
  dvhColName: 'Campo',
  dvhColCell: 'Célula',
  dvhColValue: 'Valor',
  dvhNamePlaceholder: 'Nome do campo, ex.: Project.Name',
  dvhBindCell: 'Vincular célula selecionada',
  dvhRemove: 'Remover',
  dvhEmpty: 'Ainda não há campos. Selecione uma célula, dê um nome e vincule.',
  dvhBroken: 'Não vinculado',
  dvhErrSelectCell: 'Selecione primeiro uma única célula.',
  dvhErrName: 'Insira um nome de campo ainda não usado.',
  dvhErrNoFile: 'Abra primeiro uma pasta de trabalho.',
  dvhUnboundToast: 'O campo "{name}" perdeu sua célula; vincule-o novamente.',
  dvhClose: 'Fechar',
  dvhBound: 'Campo "{name}" vinculado a {cell}.',
} satisfies Record<keyof typeof zh, string>
