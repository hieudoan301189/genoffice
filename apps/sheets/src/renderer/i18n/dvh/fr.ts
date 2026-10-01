import type { zh } from './zh'

export const fr = {
  dvhSmartData: 'Smart Data',
  dvhPanelTitle: 'Champs Smart Data',
  dvhColName: 'Champ',
  dvhColCell: 'Cellule',
  dvhColValue: 'Valeur',
  dvhNamePlaceholder: 'Nom du champ, ex. Project.Name',
  dvhBindCell: 'Lier la cellule sélectionnée',
  dvhRemove: 'Supprimer',
  dvhEmpty: 'Aucun champ. Sélectionnez une cellule, nommez-la et liez-la.',
  dvhBroken: 'Non lié',
  dvhErrSelectCell: 'Sélectionnez d’abord une seule cellule.',
  dvhErrName: 'Saisissez un nom de champ non utilisé.',
  dvhErrNoFile: 'Ouvrez d’abord un classeur.',
  dvhUnboundToast: 'Le champ « {name} » a perdu sa cellule ; liez-le à nouveau.',
  dvhClose: 'Fermer',
  dvhBound: 'Champ « {name} » lié à {cell}.',
} satisfies Record<keyof typeof zh, string>
