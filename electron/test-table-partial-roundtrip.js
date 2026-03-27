var bridge = require('./native/com_bridge_node.node');

(async () => {
  const h = bridge.comCreate('HWPFrame.HwpObject');
  console.log('HWP Version:', bridge.comGet(h, 'Version'));

  // 표 생성 + 채우기
  bridge.comCallWith(h, 'Run', ['SelectAll']);
  bridge.comCallWith(h, 'Run', ['Delete']);

  // 표 앞에 텍스트
  var act = bridge.comCallWith(h, 'CreateAction', ['InsertText']);
  var set = bridge.comCallWith(act, 'CreateSet', []);
  bridge.comCallWith(act, 'GetDefault', [set]);
  bridge.comCallWith(set, 'SetItem', ['Text', 'BEFORE TABLE']);
  bridge.comCallWith(act, 'Execute', [set]);
  bridge.comCallWith(h, 'Run', ['BreakPara']);

  // 표 생성
  var tact = bridge.comCallWith(h, 'CreateAction', ['TableCreate']);
  var tset = bridge.comCallWith(tact, 'CreateSet', []);
  bridge.comCallWith(tact, 'GetDefault', [tset]);
  bridge.comCallWith(tset, 'SetItem', ['Rows', 2]);
  bridge.comCallWith(tset, 'SetItem', ['Cols', 2]);
  bridge.comCallWith(tset, 'SetItem', ['WidthType', 2]);
  bridge.comCallWith(tact, 'Execute', [tset]);

  var cells = ['Original-A1', 'Original-B1', 'Original-A2', 'Original-B2'];
  for (var i = 0; i < cells.length; i++) {
    var a = bridge.comCallWith(h, 'CreateAction', ['InsertText']);
    var s = bridge.comCallWith(a, 'CreateSet', []);
    bridge.comCallWith(a, 'GetDefault', [s]);
    bridge.comCallWith(s, 'SetItem', ['Text', cells[i]]);
    bridge.comCallWith(a, 'Execute', [s]);
    if (i < cells.length - 1) bridge.comCallWith(h, 'Run', ['TableRightCell']);
  }

  // 표 뒤에 텍스트
  bridge.comCallWith(h, 'Run', ['MoveDocEnd']);
  bridge.comCallWith(h, 'Run', ['BreakPara']);
  act = bridge.comCallWith(h, 'CreateAction', ['InsertText']);
  set = bridge.comCallWith(act, 'CreateSet', []);
  bridge.comCallWith(act, 'GetDefault', [set]);
  bridge.comCallWith(set, 'SetItem', ['Text', 'AFTER TABLE']);
  bridge.comCallWith(act, 'Execute', [set]);

  console.log('문서: BEFORE TABLE / 표(2x2) / AFTER TABLE');

  // ── Step 1: 표 진입 → 셀 블록 선택 → HWPML2X 추출 ──
  bridge.comCallWith(h, 'Run', ['MoveDocBegin']);
  bridge.comCallWith(h, 'Run', ['MoveDown']);
  bridge.comCallWith(h, 'Run', ['TableCellBlock']);

  var origXml = bridge.comCallWith(h, 'GetTextFile', ['HWPML2X', 'saveblock']);
  console.log('\n원본 XML length:', origXml.length);
  console.log('Original-A1 포함:', origXml.indexOf('Original-A1') >= 0);

  // ── Step 2: XML 수정 ──
  var modXml = origXml
    .replace('Original-A1', 'MODIFIED-A1')
    .replace('Original-B1', 'MODIFIED-B1')
    .replace('Original-A2', 'MODIFIED-A2')
    .replace('Original-B2', 'MODIFIED-B2');
  console.log('MODIFIED-A1 포함:', modXml.indexOf('MODIFIED-A1') >= 0);

  // ── Step 3: 방법들 시도 ──

  // 방법 B: CellBlock 선택 상태에서 바로 SetTextFile (Delete 없이)
  console.log('\n=== 방법 B: CellBlock → SetTextFile (Delete 없이) ===');
  bridge.comCallWith(h, 'Run', ['Cancel']);
  bridge.comCallWith(h, 'Run', ['MoveDocBegin']);
  bridge.comCallWith(h, 'Run', ['MoveDown']);
  bridge.comCallWith(h, 'Run', ['TableCellBlock']);
  bridge.comCallWith(h, 'SetTextFile', [modXml, 'HWPML2X', '']);

  // 결과 확인
  var result = bridge.comCallWith(h, 'GetTextFile', ['UNICODE', '']);
  console.log('전체 텍스트:', JSON.stringify(result));
  console.log('MODIFIED-A1:', result.indexOf('MODIFIED-A1') >= 0);
  console.log('BEFORE TABLE:', result.indexOf('BEFORE TABLE') >= 0);
  console.log('AFTER TABLE:', result.indexOf('AFTER TABLE') >= 0);
  console.log('Original-A1:', result.indexOf('Original-A1') >= 0);

  // HWPML2X로 표 구조 확인
  bridge.comCallWith(h, 'Run', ['MoveDocBegin']);
  bridge.comCallWith(h, 'Run', ['MoveDown']);
  bridge.comCallWith(h, 'Run', ['TableCellBlock']);
  var newXml = bridge.comCallWith(h, 'GetTextFile', ['HWPML2X', 'saveblock']);
  if (newXml) {
    console.log('\n새 표 XML length:', newXml.length);
    console.log('MODIFIED-A1:', newXml.indexOf('MODIFIED-A1') >= 0);
    console.log('MODIFIED-B2:', newXml.indexOf('MODIFIED-B2') >= 0);
  } else {
    console.log('표 진입 실패 또는 XML null');
  }

  // 정리
  console.log('\n정리 중...');
  bridge.comCallWith(h, 'Clear', [1]);
  bridge.comCallWith(h, 'Quit', []);
})();
