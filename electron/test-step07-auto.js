/**
 * 단계 7: 표 (Table) — 자동 테스트
 */

var bridge = require('./native/com_bridge_node.node');
var { initBridge, createComProxy } = require('./dist/worker/proxy');

bridge.comInit();
initBridge(bridge);

var h = bridge.comCreate('HWPFrame.HwpObject');
var hwp = createComProxy(h);
var win0 = bridge.comCallWith(bridge.comGet(h, 'XHwpWindows'), 'Item', [0]);
bridge.comPut(win0, 'Visible', true);

function insert(text) {
  hwp.HAction.GetDefault('InsertText', hwp.HParameterSet.HInsertText.HSet);
  hwp.HParameterSet.HInsertText.Text = text;
  hwp.HAction.Execute('InsertText', hwp.HParameterSet.HInsertText.HSet);
}
function run(label, fn) {
  process.stdout.write('── ' + label + ': ');
  try { fn(); } catch(e) { console.log('❌ ' + e.message); }
}

console.log('HWP Version: ' + String(hwp.Version) + '\n');

// ================================================================
// 7-1: 표 생성 (TableCreate)
// ================================================================
run('7-1 TableCreate 3x2', function() {
  // HParameterSet 패턴
  hwp.HAction.GetDefault('TableCreate', hwp.HParameterSet.HTableCreation.HSet);
  hwp.HParameterSet.HTableCreation.Rows = 3;
  hwp.HParameterSet.HTableCreation.Cols = 2;
  hwp.HParameterSet.HTableCreation.WidthType = 2; // 단에 맞춤
  hwp.HParameterSet.HTableCreation.HeightType = 0; // 자동
  var ret = hwp.HAction.Execute('TableCreate', hwp.HParameterSet.HTableCreation.HSet);
  console.log('ret=' + ret + ' ✅ (커서가 표 안으로 들어갔을 것)');
});

// ================================================================
// 7-2: 셀에 텍스트 삽입
// ================================================================
run('7-2 셀에 InsertText', function() {
  insert('A1');
  console.log('✅ 현재 셀에 "A1" 입력');
});

// ================================================================
// 7-3: TableRightCell (다음 셀 이동)
// ================================================================
run('7-3 TableRightCell', function() {
  bridge.comCallWith(h, 'Run', ['TableRightCell']);
  insert('B1');
  console.log('✅ 오른쪽 셀에 "B1"');
});

// ================================================================
// 7-4: TableLowerCell (아래 셀)
// ================================================================
run('7-4 TableLowerCell', function() {
  bridge.comCallWith(h, 'Run', ['TableLowerCell']);
  insert('B2');
  console.log('✅ 아래 셀에 "B2"');
});

// ================================================================
// 7-5: TableLeftCell (왼쪽 셀)
// ================================================================
run('7-5 TableLeftCell', function() {
  bridge.comCallWith(h, 'Run', ['TableLeftCell']);
  insert('A2');
  console.log('✅ 왼쪽 셀에 "A2"');
});

// ================================================================
// 7-6: TableUpperCell (위 셀)
// ================================================================
run('7-6 TableUpperCell', function() {
  bridge.comCallWith(h, 'Run', ['TableUpperCell']);
  insert(' (edited)');
  console.log('✅ 위 셀 "A1 (edited)"');
});

// ================================================================
// 7-7: 나머지 셀 채우기
// ================================================================
run('7-7 나머지 셀 채우기', function() {
  // A3로 이동
  bridge.comCallWith(h, 'Run', ['TableLowerCell']);
  bridge.comCallWith(h, 'Run', ['TableLowerCell']);
  insert('A3');
  bridge.comCallWith(h, 'Run', ['TableRightCell']);
  insert('B3');
  console.log('✅ A3, B3 입력');
});

// ================================================================
// 7-8: 표 밖으로 나가기
// ================================================================
run('7-8 표 밖으로 나가기', function() {
  bridge.comCallWith(h, 'Run', ['MoveDocEnd']);
  // 표 뒤에 텍스트 입력
  insert('\r\nText after table');
  var text = String(bridge.comCallWith(h, 'GetTextFile', ['UNICODE', '']));
  var hasAfter = text.indexOf('Text after table') >= 0;
  console.log(hasAfter ? '✅ 표 밖으로 나옴' : '❌ 표 안에 갇힘');
});

// ================================================================
// 7-9: GetTextFile로 표 내용 읽기
// ================================================================
run('7-9 GetTextFile 표 내용', function() {
  var text = String(bridge.comCallWith(h, 'GetTextFile', ['UNICODE', '']));
  console.log('전체 텍스트:');
  text.split(/\r?\n/).forEach(function(line, i) {
    console.log('  [' + i + '] "' + line + '"');
  });
});

// ================================================================
// 7-10: 표 안에서 CharShape 적용
// ================================================================
run('7-10 표 안에서 CharShape', function() {
  // 표 안으로 다시 들어가기 — MoveDocBegin → 표 첫 셀
  bridge.comCallWith(h, 'Run', ['MoveDocBegin']);
  // 표가 첫 문단이면 MoveDocBegin이 표 안으로 갈 수도
  // 확인
  bridge.comCallWith(h, 'Run', ['MoveSelRight']);
  bridge.comCallWith(h, 'Run', ['MoveSelRight']);
  var sel = String(bridge.comCallWith(h, 'GetTextFile', ['UNICODE', 'saveblock']));
  bridge.comCallWith(h, 'Run', ['Cancel']);
  console.log('표 첫 셀 확인: "' + sel + '"');

  if (sel.indexOf('A1') >= 0) {
    // A1 셀 텍스트 선택 후 볼드 적용
    bridge.comCallWith(h, 'Run', ['MoveDocBegin']);
    bridge.comCallWith(h, 'Run', ['MoveSelLineEnd']);
    hwp.HAction.GetDefault('CharShape', hwp.HParameterSet.HCharShape.HSet);
    hwp.HParameterSet.HCharShape.Bold = 1;
    hwp.HParameterSet.HCharShape.Height = 1600;
    hwp.HAction.Execute('CharShape', hwp.HParameterSet.HCharShape.HSet);
    bridge.comCallWith(h, 'Run', ['Cancel']);
    console.log('  A1에 16pt 볼드 적용');
  }
});

// ================================================================
// 7-11: CreateAction 패턴으로 표 생성
// ================================================================
run('7-11 CreateAction("TableCreate")', function() {
  bridge.comCallWith(h, 'Run', ['MoveDocEnd']);
  insert('\r\n');

  var act = hwp.CreateAction('TableCreate');
  var set = act.CreateSet();
  act.GetDefault(set);
  set.SetItem('Rows', 2);
  set.SetItem('Cols', 3);
  set.SetItem('WidthType', 2);
  var ret = act.Execute(set);
  console.log('ret=' + ret + ' (2x3 표)');

  // 셀 채우기
  var cells = ['X1', 'Y1', 'Z1', 'X2', 'Y2', 'Z2'];
  for (var i = 0; i < cells.length; i++) {
    insert(cells[i]);
    if (i < cells.length - 1) bridge.comCallWith(h, 'Run', ['TableRightCell']);
  }
  console.log('  6개 셀 채움');
});

// 닫지 않음
console.log('\n확인해주세요! (수동으로 닫으세요)');
