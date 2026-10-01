import type { zh } from './zh'

export const ru = {
  dvhSmartData: 'Smart Data',
  dvhPanelTitle: 'Поля Smart Data',
  dvhColName: 'Поле',
  dvhColCell: 'Ячейка',
  dvhColValue: 'Значение',
  dvhNamePlaceholder: 'Имя поля, напр. Project.Name',
  dvhBindCell: 'Привязать выбранную ячейку',
  dvhRemove: 'Удалить',
  dvhEmpty: 'Полей пока нет. Выберите ячейку, задайте имя и привяжите.',
  dvhBroken: 'Не привязано',
  dvhErrSelectCell: 'Сначала выберите одну ячейку.',
  dvhErrName: 'Введите неиспользуемое имя поля.',
  dvhErrNoFile: 'Сначала откройте книгу.',
  dvhUnboundToast: 'Поле «{name}» потеряло ячейку; привяжите его снова.',
  dvhClose: 'Закрыть',
  dvhBound: 'Поле «{name}» привязано к {cell}.',
} satisfies Record<keyof typeof zh, string>
