function onEdit(e) {
  try {
    AUTOEDIT_handleSimpleOnEdit_(e);
  } catch (err) {
    console.error('[onEdit][AUTOEDIT] ' + (err && err.stack ? err.stack : err));
  }

  // 친구분이 만든 B열 동기화 함수가 프로젝트에 있을 때만 실행
  if (typeof B열_양방향_동기화_onEdit === 'function') {
    try {
      B열_양방향_동기화_onEdit(e);
    } catch (err) {
      console.error('[onEdit][B열_양방향_동기화_onEdit] ' + (err && err.stack ? err.stack : err));
    }
  }
}
