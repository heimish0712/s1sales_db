function onEdit(e) {
  autoFillRegionOnAddressEdit_(e);      // 주소 → 지역구분 자동입력
  autoFillGradeOnAreaEdit_(e);
  autoFillContractDefaultsOnUnitEdit_(e);
  autoCalcFinalQuotePriceOnEdit_(e);    // 계약조건 → 최종 견적가 자동계산

  // 친구분이 만든 B열 동기화 함수가 프로젝트에 있을 때만 실행
  if (typeof B열_양방향_동기화_onEdit === 'function') {
    B열_양방향_동기화_onEdit(e);
  }
}
