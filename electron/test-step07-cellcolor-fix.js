/**
 * 셀 배경색 — WinBrushFaceStyle=-1 로 수정
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

function setCellBg(r, g, b) {
  hwp.HAction.GetDefault('CellBorderFill', hwp.HParameterSet.HCellBorderFill.HSet);
  var hps = bridge.comGet(h, 'HParameterSet');
  var hcbf = bridge.comGet(hps, 'HCellBorderFill');
  var fa = bridge.comGet(hcbf, 'FillAttr');

  bridge.comPut(fa, 'type', 1);
  bridge.comPut(fa, 'WindowsBrush', 1);
  bridge.comPut(fa, 'WinBrushFaceColor', bridge.comCallWith(h, 'RGBColor', [r, g, b]));
  bridge.comPut(fa, 'WinBrushFaceStyle', -1);  // ← 핵심!

  hwp.HAction.Execute('CellBorderFill', hwp.HParameterSet.HCellBorderFill.HSet);
}

// 2x3 표 생성
hwp.HAction.GetDefault('TableCreate', hwp.HParameterSet.HTableCreation.HSet);
hwp.HParameterSet.HTableCreation.Rows = 2;
hwp.HParameterSet.HTableCreation.Cols = 3;
hwp.HParameterSet.HTableCreation.WidthType = 2;
hwp.HAction.Execute('TableCreate', hwp.HParameterSet.HTableCreation.HSet);

// A1: 빨강
insert('A1');
setCellBg(255, 0, 0);
console.log('A1: 빨강');

// B1: 파랑
bridge.comCallWith(h, 'Run', ['TableRightCell']);
insert('B1');
setCellBg(0, 0, 255);
console.log('B1: 파랑');

// C1: 노랑
bridge.comCallWith(h, 'Run', ['TableRightCell']);
insert('C1');
setCellBg(255, 255, 0);
console.log('C1: 노랑');

// A2: 초록
bridge.comCallWith(h, 'Run', ['TableLowerCell']);
bridge.comCallWith(h, 'Run', ['TableLeftCell']);
bridge.comCallWith(h, 'Run', ['TableLeftCell']);
insert('A2');
setCellBg(0, 200, 0);
console.log('A2: 초록');

// B2: 분홍
bridge.comCallWith(h, 'Run', ['TableRightCell']);
insert('B2');
setCellBg(255, 180, 200);
console.log('B2: 분홍');

// C2: 기본 (배경 안 넣음)
bridge.comCallWith(h, 'Run', ['TableRightCell']);
insert('C2 (기본)');
console.log('C2: 기본');

console.log('\n확인해주세요!');
