import type { zh } from './zh'

export const ms = {
  dvhSmartData: 'Smart Data',
  dvhPanelTitle: 'Medan Smart Data',
  dvhColName: 'Medan',
  dvhColCell: 'Sel',
  dvhColValue: 'Nilai',
  dvhNamePlaceholder: 'Nama medan, cth. Project.Name',
  dvhBindCell: 'Ikat sel dipilih',
  dvhRemove: 'Buang',
  dvhEmpty: 'Belum ada medan. Pilih sel, namakan dan ikat.',
  dvhBroken: 'Tidak diikat',
  dvhErrSelectCell: 'Pilih satu sel dahulu.',
  dvhErrName: 'Masukkan nama medan yang belum digunakan.',
  dvhErrNoFile: 'Buka buku kerja dahulu.',
  dvhUnboundToast: 'Medan "{name}" kehilangan selnya; ikat semula.',
  dvhClose: 'Tutup',
  dvhBound: 'Medan "{name}" diikat ke {cell}.',
} satisfies Record<keyof typeof zh, string>
