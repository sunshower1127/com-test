var bridge = require('./native/com_bridge_node.node');

(async () => {
  const h = bridge.comCreate('HWPFrame.HwpObject');
  const ver = bridge.comGet(h, 'Version');
  console.log('HWP Version:', ver);

  // 표 생성 + 채우기
  bridge.comCallWith(h, 'Run', ['SelectAll']);
  bridge.comCallWith(h, 'Run', ['Delete']);

  var act = bridge.comCallWith(h, 'CreateAction', ['TableCreate']);
  var set = bridge.comCallWith(act, 'CreateSet', []);
  bridge.comCallWith(act, 'GetDefault', [set]);
  bridge.comCallWith(set, 'SetItem', ['Rows', 2]);
  bridge.comCallWith(set, 'SetItem', ['Cols', 3]);
  bridge.comCallWith(set, 'SetItem', ['WidthType', 2]);
  bridge.comCallWith(act, 'Execute', [set]);

  // 셀 채우기
  var cells = ['A1', 'B1', 'C1', 'A2', 'B2', 'C2'];
  for (var i = 0; i < cells.length; i++) {
    var act2 = bridge.comCallWith(h, 'CreateAction', ['InsertText']);
    var set2 = bridge.comCallWith(act2, 'CreateSet', []);
    bridge.comCallWith(act2, 'GetDefault', [set2]);
    bridge.comCallWith(set2, 'SetItem', ['Text', cells[i]]);
    bridge.comCallWith(act2, 'Execute', [set2]);
    if (i < cells.length - 1) bridge.comCallWith(h, 'Run', ['TableRightCell']);
  }

  // 표 밖으로
  bridge.comCallWith(h, 'Run', ['MoveDocEnd']);
  bridge.comCallWith(h, 'Run', ['BreakPara']);
  var act3 = bridge.comCallWith(h, 'CreateAction', ['InsertText']);
  var set3 = bridge.comCallWith(act3, 'CreateSet', []);
  bridge.comCallWith(act3, 'GetDefault', [set3]);
  bridge.comCallWith(set3, 'SetItem', ['Text', 'OUTSIDE']);
  bridge.comCallWith(act3, 'Execute', [set3]);

  console.log('\n표 생성 완료: A1~C2 + OUTSIDE');

  // ── 테스트 시작 ──

  // T1: MoveDown으로 진입 → TableCellBlock
  console.log('\n=== T1: MoveDown → TableCellBlock ===');
  bridge.comCallWith(h, 'Run', ['MoveDocBegin']);
  bridge.comCallWith(h, 'Run', ['MoveDown']);
  bridge.comCallWith(h, 'Run', ['TableCellBlock']);
  var t1 = bridge.comCallWith(h, 'GetTextFile', ['UNICODE', 'saveblock']);
  console.log('TableCellBlock:', JSON.stringify(t1));
  bridge.comCallWith(h, 'Run', ['Cancel']);

  // T2: MoveDown → TableCellBlockExtend (확장 선택)
  console.log('\n=== T2: TableCellBlockExtend ===');
  bridge.comCallWith(h, 'Run', ['MoveDocBegin']);
  bridge.comCallWith(h, 'Run', ['MoveDown']);
  bridge.comCallWith(h, 'Run', ['TableCellBlock']);
  bridge.comCallWith(h, 'Run', ['TableCellBlockExtend']);
  var t2 = bridge.comCallWith(h, 'GetTextFile', ['UNICODE', 'saveblock']);
  console.log('CellBlock+Extend:', JSON.stringify(t2));
  bridge.comCallWith(h, 'Run', ['Cancel']);

  // T3: SelectAll inside table
  console.log('\n=== T3: 표 안에서 SelectAll ===');
  bridge.comCallWith(h, 'Run', ['MoveDocBegin']);
  bridge.comCallWith(h, 'Run', ['MoveDown']);
  bridge.comCallWith(h, 'Run', ['SelectAll']);
  var t3 = bridge.comCallWith(h, 'GetTextFile', ['UNICODE', 'saveblock']);
  console.log('SelectAll in table:', JSON.stringify(t3));
  bridge.comCallWith(h, 'Run', ['Cancel']);

  // T4: TableSelBlock로 전체 표 선택
  console.log('\n=== T4: TableSelBlock ===');
  bridge.comCallWith(h, 'Run', ['MoveDocBegin']);
  bridge.comCallWith(h, 'Run', ['MoveDown']);
  bridge.comCallWith(h, 'Run', ['TableSelBlock']);
  var t4 = bridge.comCallWith(h, 'GetTextFile', ['UNICODE', 'saveblock']);
  console.log('TableSelBlock:', JSON.stringify(t4));
  bridge.comCallWith(h, 'Run', ['Cancel']);

  // T5: TableCellBlock → TableRightCell로 확장
  console.log('\n=== T5: CellBlock → TableSelCellRight ===');
  bridge.comCallWith(h, 'Run', ['MoveDocBegin']);
  bridge.comCallWith(h, 'Run', ['MoveDown']);
  bridge.comCallWith(h, 'Run', ['TableCellBlock']);
  bridge.comCallWith(h, 'Run', ['TableSelCellRight']);
  var t5 = bridge.comCallWith(h, 'GetTextFile', ['UNICODE', 'saveblock']);
  console.log('CellBlock+SelRight:', JSON.stringify(t5));
  bridge.comCallWith(h, 'Run', ['Cancel']);

  // T6: 전체 표 블록 선택 후 HWPML2X
  console.log('\n=== T6: 전체 표 블록 → HWPML2X ===');
  bridge.comCallWith(h, 'Run', ['MoveDocBegin']);
  bridge.comCallWith(h, 'Run', ['MoveDown']);
  bridge.comCallWith(h, 'Run', ['TableCellBlock']);
  // 오른쪽 끝까지
  bridge.comCallWith(h, 'Run', ['TableSelCellRight']);
  bridge.comCallWith(h, 'Run', ['TableSelCellRight']);
  // 아래쪽 끝까지
  bridge.comCallWith(h, 'Run', ['TableSelCellBelow']);
  var t6text = bridge.comCallWith(h, 'GetTextFile', ['UNICODE', 'saveblock']);
  var t6xml = bridge.comCallWith(h, 'GetTextFile', ['HWPML2X', 'saveblock']);
  console.log('UNICODE:', JSON.stringify(t6text));
  console.log('HWPML2X len:', t6xml ? t6xml.length : 0);
  if (t6xml) {
    console.log('TABLE:', t6xml.indexOf('TABLE') >= 0);
    console.log('A1:', t6xml.indexOf('A1') >= 0);
    console.log('C2:', t6xml.indexOf('C2') >= 0);
  }
  bridge.comCallWith(h, 'Run', ['Cancel']);

  // 정리
  console.log('\n정리 중...');
  bridge.comCallWith(h, 'Clear', [1]);
  bridge.comCallWith(h, 'Quit', []);
})();
