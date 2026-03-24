/**
 * 단계 7: 표 서식 데모 (닫지 않음)
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

function applyCharShape(props) {
  hwp.HAction.GetDefault('CharShape', hwp.HParameterSet.HCharShape.HSet);
  Object.keys(props).forEach(function(k) {
    hwp.HParameterSet.HCharShape[k] = props[k];
  });
  hwp.HAction.Execute('CharShape', hwp.HParameterSet.HCharShape.HSet);
}

function applyParaShape(props) {
  hwp.HAction.GetDefault('ParagraphShape', hwp.HParameterSet.HParaShape.HSet);
  Object.keys(props).forEach(function(k) {
    hwp.HParameterSet.HParaShape[k] = props[k];
  });
  hwp.HAction.Execute('ParagraphShape', hwp.HParameterSet.HParaShape.HSet);
}

// ── 표 1: 기본 3x3 + 셀별 서식 ──
console.log('표 1: 3x3 서식 데모 생성 중...');
hwp.HAction.GetDefault('TableCreate', hwp.HParameterSet.HTableCreation.HSet);
hwp.HParameterSet.HTableCreation.Rows = 3;
hwp.HParameterSet.HTableCreation.Cols = 3;
hwp.HParameterSet.HTableCreation.WidthType = 2;
hwp.HAction.Execute('TableCreate', hwp.HParameterSet.HTableCreation.HSet);

// 헤더 행 (첫 행) — 볼드 + 가운데 정렬
var headers = ['이름', '나이', '직업'];
for (var i = 0; i < 3; i++) {
  insert(headers[i]);
  bridge.comCallWith(h, 'Run', ['MoveSelLineBegin']); // 셀 전체 선택
  applyCharShape({ Bold: 1, Height: 1200 });
  applyParaShape({ AlignType: 3 }); // 가운데
  bridge.comCallWith(h, 'Run', ['Cancel']);
  if (i < 2) bridge.comCallWith(h, 'Run', ['TableRightCell']);
}

// 데이터 행 1
bridge.comCallWith(h, 'Run', ['TableLowerCell']);
bridge.comCallWith(h, 'Run', ['TableLeftCell']);
bridge.comCallWith(h, 'Run', ['TableLeftCell']);

var row1 = ['홍길동', '25', '개발자'];
for (var i = 0; i < 3; i++) {
  insert(row1[i]);
  if (i === 1) { // 나이: 오른쪽 정렬
    bridge.comCallWith(h, 'Run', ['MoveSelLineBegin']);
    applyParaShape({ AlignType: 2 });
    bridge.comCallWith(h, 'Run', ['Cancel']);
  }
  if (i === 2) { // 직업: 이탤릭
    bridge.comCallWith(h, 'Run', ['MoveSelLineBegin']);
    applyCharShape({ Italic: 1 });
    bridge.comCallWith(h, 'Run', ['Cancel']);
  }
  if (i < 2) bridge.comCallWith(h, 'Run', ['TableRightCell']);
}

// 데이터 행 2
bridge.comCallWith(h, 'Run', ['TableLowerCell']);
bridge.comCallWith(h, 'Run', ['TableLeftCell']);
bridge.comCallWith(h, 'Run', ['TableLeftCell']);

var row2 = ['김철수', '30', '디자이너'];
for (var i = 0; i < 3; i++) {
  insert(row2[i]);
  if (i === 0) { // 이름: 빨간색
    bridge.comCallWith(h, 'Run', ['MoveSelLineBegin']);
    applyCharShape({ TextColor: bridge.comCallWith(h, 'RGBColor', [255, 0, 0]) });
    bridge.comCallWith(h, 'Run', ['Cancel']);
  }
  if (i === 1) { // 나이: 오른쪽 정렬 + 볼드
    bridge.comCallWith(h, 'Run', ['MoveSelLineBegin']);
    applyCharShape({ Bold: 1 });
    applyParaShape({ AlignType: 2 });
    bridge.comCallWith(h, 'Run', ['Cancel']);
  }
  if (i < 2) bridge.comCallWith(h, 'Run', ['TableRightCell']);
}

// 표 밖으로
bridge.comCallWith(h, 'Run', ['MoveDocEnd']);

// ── 표 2: CreateAction 패턴 + 다른 크기 ──
console.log('표 2: 2x4 생성 중...');
insert('\r\n');

var act = hwp.CreateAction('TableCreate');
var set = act.CreateSet();
act.GetDefault(set);
set.SetItem('Rows', 2);
set.SetItem('Cols', 4);
set.SetItem('WidthType', 2);
act.Execute(set);

var data2 = ['Q1', 'Q2', 'Q3', 'Q4', '100', '200', '300', '400'];
for (var i = 0; i < 8; i++) {
  insert(data2[i]);
  // 첫 행은 볼드 가운데
  if (i < 4) {
    bridge.comCallWith(h, 'Run', ['MoveSelLineBegin']);
    applyCharShape({ Bold: 1 });
    applyParaShape({ AlignType: 3 });
    bridge.comCallWith(h, 'Run', ['Cancel']);
  }
  // 둘째 행은 오른쪽 정렬
  if (i >= 4) {
    bridge.comCallWith(h, 'Run', ['MoveSelLineBegin']);
    applyParaShape({ AlignType: 2 });
    bridge.comCallWith(h, 'Run', ['Cancel']);
  }
  if (i < 7) bridge.comCallWith(h, 'Run', ['TableRightCell']);
}

bridge.comCallWith(h, 'Run', ['MoveDocEnd']);

console.log('\n확인해주세요!');
console.log('표 1: 3x3 — 헤더 볼드+가운데, 나이 오른쪽, 직업 이탤릭, 김철수 빨간색');
console.log('표 2: 2x4 — Q1~Q4 볼드 가운데, 100~400 오른쪽 정렬');
