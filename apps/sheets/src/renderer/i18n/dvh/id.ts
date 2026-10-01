import type { zh } from './zh'

export const id = {
  dvhSmartData: 'Smart Data',
  dvhPanelTitle: 'Bidang Smart Data',
  dvhColName: 'Bidang',
  dvhColCell: 'Sel',
  dvhColValue: 'Nilai',
  dvhNamePlaceholder: 'Nama bidang, mis. Project.Name',
  dvhBindCell: 'Ikat sel terpilih',
  dvhRemove: 'Hapus',
  dvhEmpty: 'Belum ada bidang. Pilih sel, beri nama, lalu ikat.',
  dvhBroken: 'Tidak terikat',
  dvhErrSelectCell: 'Pilih satu sel terlebih dahulu.',
  dvhErrName: 'Masukkan nama bidang yang belum dipakai.',
  dvhErrNoFile: 'Buka buku kerja terlebih dahulu.',
  dvhUnboundToast: 'Bidang "{name}" kehilangan selnya; ikat lagi.',
  dvhClose: 'Tutup',
  dvhBound: 'Bidang "{name}" diikat ke {cell}.',
} satisfies Record<keyof typeof zh, string>
