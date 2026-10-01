import type { zh } from './zh'

export const ko = {
  dvhSmartData: '스마트 데이터',
  dvhPanelTitle: '스마트 데이터 필드',
  dvhColName: '필드',
  dvhColCell: '셀',
  dvhColValue: '값',
  dvhNamePlaceholder: '필드 이름(예: Project.Name)',
  dvhBindCell: '선택한 셀 바인딩',
  dvhRemove: '제거',
  dvhEmpty: '아직 필드가 없습니다. 셀을 선택하고 이름을 지정해 바인딩하세요.',
  dvhBroken: '바인딩 안 됨',
  dvhErrSelectCell: '먼저 셀 하나를 선택하세요.',
  dvhErrName: '사용하지 않은 필드 이름을 입력하세요.',
  dvhErrNoFile: '먼저 통합 문서를 여세요.',
  dvhUnboundToast: '필드 "{name}"의 셀이 삭제되었습니다. 다시 바인딩하세요.',
  dvhClose: '닫기',
  dvhBound: '필드 "{name}"을(를) {cell}에 바인딩했습니다.',
} satisfies Record<keyof typeof zh, string>
