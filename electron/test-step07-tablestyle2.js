/**
 * 표 서식 2차: FillAttr 탐색 + 테두리/배경 적용
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

// 3x3 표 생성 + 채우기
hwp.HAction.GetDefault('TableCreate', hwp.HParameterSet.HTableCreation.HSet);
hwp.HParameterSet.HTableCreation.Rows = 3;
hwp.HParameterSet.HTableCreation.Cols = 3;
hwp.HParameterSet.HTableCreation.WidthType = 2;
hwp.HAction.Execute('TableCreate', hwp.HParameterSet.HTableCreation.HSet);

var cells = ['A1','B1','C1','A2','B2','C2','A3','B3','C3'];
for (var i = 0; i < 9; i++) {
  insert(cells[i]);
  if (i < 8) bridge.comCallWith(h, 'Run', ['TableRightCell']);
}
// A1으로 복귀
for (var i = 0; i < 2; i++) bridge.comCallWith(h, 'Run', ['TableUpperCell']);
for (var i = 0; i < 2; i++) bridge.comCallWith(h, 'Run', ['TableLeftCell']);
console.log('3x3 표 생성 완료\n');

// ── FillAttr 멤버 탐색 ──
run('T1 FillAttr 멤버', function() {
  hwp.HAction.GetDefault('CellBorderFill', hwp.HParameterSet.HCellBorderFill.HSet);
  var hps = bridge.comGet(h, 'HParameterSet');
  var hcbf = bridge.comGet(hps, 'HCellBorderFill');
  var fillAttr = bridge.comGet(hcbf, 'FillAttr');

  console.log('FillAttr type=' + typeof fillAttr);
  var members = bridge.comListMembers(fillAttr);
  var props = members.filter(function(m) {
    return ['QueryInterface','AddRef','Release','GetTypeInfoCount','GetTypeInfo','GetIDsOfNames','Invoke'].indexOf(m.name) === -1;
  });
  props.forEach(function(m) {
    var val = '';
    try {
      var v = bridge.comGet(fillAttr, m.name);
      val = (typeof v !== 'object') ? ' = ' + v : ' = [object]';
    } catch(e) { val = ' (읽기 실패)'; }
    console.log('\n  ' + m.name + ' (' + m.kind + ')' + val);
  });
});

// ── A1: 배경색 ──
run('\nT2 A1 배경색 (FillAttr.FaceColor)', function() {
  hwp.HAction.GetDefault('CellBorderFill', hwp.HParameterSet.HCellBorderFill.HSet);
  var hps = bridge.comGet(h, 'HParameterSet');
  var hcbf = bridge.comGet(hps, 'HCellBorderFill');
  var fillAttr = bridge.comGet(hcbf, 'FillAttr');

  // FaceColor 설정 시도
  var pink = bridge.comCallWith(h, 'RGBColor', [255, 200, 200]);
  try {
    bridge.comPut(fillAttr, 'FaceColor', pink);
    bridge.comPut(fillAttr, 'FillType', 1); // 단색 채우기?
    console.log('FaceColor=' + pink + ', FillType=1 설정');
  } catch(e) { console.log('FaceColor 실패: ' + e.message); }

  hwp.HAction.Execute('CellBorderFill', hwp.HParameterSet.HCellBorderFill.HSet);
  console.log('  Execute 완료');
});

// ── B1: 테두리 ──
run('\nT3 B1 테두리 (두꺼운 빨간선)', function() {
  bridge.comCallWith(h, 'Run', ['TableRightCell']); // B1으로

  hwp.HAction.GetDefault('CellBorderFill', hwp.HParameterSet.HCellBorderFill.HSet);
  var red = bridge.comCallWith(h, 'RGBColor', [255, 0, 0]);

  hwp.HParameterSet.HCellBorderFill.BorderTypeTop = 1;    // 실선
  hwp.HParameterSet.HCellBorderFill.BorderTypeBottom = 1;
  hwp.HParameterSet.HCellBorderFill.BorderTypeLeft = 1;
  hwp.HParameterSet.HCellBorderFill.BorderTypeRight = 1;

  hwp.HParameterSet.HCellBorderFill.BorderWidthTop = 3;   // 굵기
  hwp.HParameterSet.HCellBorderFill.BorderWidthBottom = 3;
  hwp.HParameterSet.HCellBorderFill.BorderWidthLeft = 3;
  hwp.HParameterSet.HCellBorderFill.BorderWidthRight = 3;

  hwp.HParameterSet.HCellBorderFill.BorderColorTop = red;
  hwp.HParameterSet.HCellBorderFill.BorderColorBottom = red;
  hwp.HParameterSet.HCellBorderFill.BorderCorlorLeft = red; // 오타 주의: Corlor
  hwp.HParameterSet.HCellBorderFill.BorderColorRight = red;

  hwp.HAction.Execute('CellBorderFill', hwp.HParameterSet.HCellBorderFill.HSet);
  console.log('✅ B1 빨간 테두리 적용');
});

// ── C1: 배경 + 테두리 동시 ──
run('\nT4 C1 파란 배경 + 두꺼운 테두리', function() {
  bridge.comCallWith(h, 'Run', ['TableRightCell']); // C1으로

  hwp.HAction.GetDefault('CellBorderFill', hwp.HParameterSet.HCellBorderFill.HSet);

  var hps = bridge.comGet(h, 'HParameterSet');
  var hcbf = bridge.comGet(hps, 'HCellBorderFill');
  var fillAttr = bridge.comGet(hcbf, 'FillAttr');

  var lightBlue = bridge.comCallWith(h, 'RGBColor', [200, 220, 255]);
  try {
    bridge.comPut(fillAttr, 'FaceColor', lightBlue);
    bridge.comPut(fillAttr, 'FillType', 1);
  } catch(e) {}

  hwp.HParameterSet.HCellBorderFill.BorderTypeTop = 1;
  hwp.HParameterSet.HCellBorderFill.BorderTypeBottom = 1;
  hwp.HParameterSet.HCellBorderFill.BorderTypeLeft = 1;
  hwp.HParameterSet.HCellBorderFill.BorderTypeRight = 1;
  hwp.HParameterSet.HCellBorderFill.BorderWidthTop = 5;
  hwp.HParameterSet.HCellBorderFill.BorderWidthBottom = 5;
  hwp.HParameterSet.HCellBorderFill.BorderWidthLeft = 5;
  hwp.HParameterSet.HCellBorderFill.BorderWidthRight = 5;

  hwp.HAction.Execute('CellBorderFill', hwp.HParameterSet.HCellBorderFill.HSet);
  console.log('✅');
});

// ── A2: 노란 배경 ──
run('\nT5 A2 노란 배경', function() {
  bridge.comCallWith(h, 'Run', ['TableLowerCell']);
  bridge.comCallWith(h, 'Run', ['TableLeftCell']);
  bridge.comCallWith(h, 'Run', ['TableLeftCell']);

  hwp.HAction.GetDefault('CellBorderFill', hwp.HParameterSet.HCellBorderFill.HSet);
  var hps = bridge.comGet(h, 'HParameterSet');
  var hcbf = bridge.comGet(hps, 'HCellBorderFill');
  var fillAttr = bridge.comGet(hcbf, 'FillAttr');

  var yellow = bridge.comCallWith(h, 'RGBColor', [255, 255, 150]);
  try {
    bridge.comPut(fillAttr, 'FaceColor', yellow);
    bridge.comPut(fillAttr, 'FillType', 1);
  } catch(e) {}

  hwp.HAction.Execute('CellBorderFill', hwp.HParameterSet.HCellBorderFill.HSet);
  console.log('✅');
});

// ── 전체 행 선택 후 배경색 ──
run('\nT6 3행 전체 초록 배경 (셀 하나씩)', function() {
  // A3 → B3 → C3 각각
  bridge.comCallWith(h, 'Run', ['TableLowerCell']); // A3

  for (var i = 0; i < 3; i++) {
    hwp.HAction.GetDefault('CellBorderFill', hwp.HParameterSet.HCellBorderFill.HSet);
    var hps = bridge.comGet(h, 'HParameterSet');
    var hcbf = bridge.comGet(hps, 'HCellBorderFill');
    var fillAttr = bridge.comGet(hcbf, 'FillAttr');
    var green = bridge.comCallWith(h, 'RGBColor', [200, 255, 200]);
    try {
      bridge.comPut(fillAttr, 'FaceColor', green);
      bridge.comPut(fillAttr, 'FillType', 1);
    } catch(e) {}
    hwp.HAction.Execute('CellBorderFill', hwp.HParameterSet.HCellBorderFill.HSet);
    if (i < 2) bridge.comCallWith(h, 'Run', ['TableRightCell']);
  }
  console.log('✅');
});

// 닫지 않음
console.log('\n확인해주세요!');
console.log('A1=분홍 배경, B1=빨간 테두리, C1=파란배경+두꺼운 테두리');
console.log('A2=노란 배경, A3/B3/C3=초록 배경');
