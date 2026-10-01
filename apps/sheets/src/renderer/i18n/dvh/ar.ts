import type { zh } from './zh'

export const ar = {
  dvhSmartData: 'البيانات الذكية',
  dvhPanelTitle: 'حقول البيانات الذكية',
  dvhColName: 'الحقل',
  dvhColCell: 'الخلية',
  dvhColValue: 'القيمة',
  dvhNamePlaceholder: 'اسم الحقل، مثل Project.Name',
  dvhBindCell: 'ربط الخلية المحددة',
  dvhRemove: 'إزالة',
  dvhEmpty: 'لا توجد حقول بعد. حدد خلية وسمّها ثم اربطها.',
  dvhBroken: 'غير مرتبط',
  dvhErrSelectCell: 'حدد خلية واحدة أولاً.',
  dvhErrName: 'أدخل اسم حقل غير مستخدم.',
  dvhErrNoFile: 'افتح مصنفًا أولاً.',
  dvhUnboundToast: 'فقد الحقل "{name}" خليته؛ اربطه مرة أخرى.',
  dvhClose: 'إغلاق',
  dvhBound: 'تم ربط الحقل "{name}" بالخلية {cell}.',
} satisfies Record<keyof typeof zh, string>
