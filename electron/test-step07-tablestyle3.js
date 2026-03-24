/**
 * 표 서식 3차: WinBrushFaceColor로 배경색
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

// 2x2 표
hwp.HAction.GetDefault('TableCreate', hwp.HParameterSet.HTableCreation.HSet);
hwp.HParameterSet.HTableCreation.Rows = 2;
hwp.HParameterSet.HTableCreation.Cols = 2;
hwp.HParameterSet.HTableCreation.WidthType = 2;
hwp.HAction.Execute('TableCreate', hwp.HParameterSet.HTableCreation.HSet);

insert('A1'); bridge.comCallWith(h, 'Run', ['TableRightCell']);
insert('B1'); bridge.comCallWith(h, 'Run', ['TableLowerCell']);
insert('B2'); bridge.comCallWith(h, 'Run', ['TableLeftCell']);
insert('A2');

// A1으로
bridge.comCallWith(h, 'Run', ['TableUpperCell']);
console.log('2x2 표 준비 완료\n');

function setCellBg(colorRGB, label) {
  var color = bridge.comCallWith(h, 'RGBColor', colorRGB);

  hwp.HAction.GetDefault('CellBorderFill', hwp.HParameterSet.HCellBorderFill.HSet);
  var hps = bridge.comGet(h, 'HParameterSet');
  var hcbf = bridge.comGet(hps, 'HCellBorderFill');
  var fillAttr = bridge.comGet(hcbf, 'FillAttr');

  // 방법 1: WinBrushFaceColor + WindowsBrush
  try {
    bridge.comPut(fillAttr, 'WinBrushFaceColor', color);
    bridge.comPut(fillAttr, 'WindowsBrush', 1);
    console.log('  ' + label + ': WinBrushFaceColor=' + color + ', WindowsBrush=1');
  } catch(e) { console.log('  ' + label + ' WinBrush 실패: ' + e.message); }

  // 방법 2: type 설정
  try {
    bridge.comPut(fillAttr, 'type', 1);
  } catch(e) {}

  hwp.HAction.Execute('CellBorderFill', hwp.HParameterSet.HCellBorderFill.HSet);
}

// A1: 분홍
console.log('── A1 분홍');
setCellBg([255, 200, 200], 'A1');

// B1: 파랑
bridge.comCallWith(h, 'Run', ['TableRightCell']);
console.log('── B1 파랑');
setCellBg([200, 220, 255], 'B1');

// A2: 노랑
bridge.comCallWith(h, 'Run', ['TableLowerCell']);
bridge.comCallWith(h, 'Run', ['TableLeftCell']);
console.log('── A2 노랑');
setCellBg([255, 255, 150], 'A2');

// B2: 초록
bridge.comCallWith(h, 'Run', ['TableRightCell']);
console.log('── B2 초록');
setCellBg([200, 255, 200], 'B2');

// WinBrushFaceStyle도 시도
console.log('\n── 다시 A1에서 WinBrushFaceStyle 변경');
bridge.comCallWith(h, 'Run', ['TableUpperCell']);
bridge.comCallWith(h, 'Run', ['TableLeftCell']);

hwp.HAction.GetDefault('CellBorderFill', hwp.HParameterSet.HCellBorderFill.HSet);
var hps2 = bridge.comGet(h, 'HParameterSet');
var hcbf2 = bridge.comGet(hps2, 'HCellBorderFill');
var fa2 = bridge.comGet(hcbf2, 'FillAttr');

var pink = bridge.comCallWith(h, 'RGBColor', [255, 150, 150]);
bridge.comPut(fa2, 'WinBrushFaceColor', pink);
bridge.comPut(fa2, 'WinBrushFaceStyle', 1); // 스타일 1
bridge.comPut(fa2, 'WindowsBrush', 1);
bridge.comPut(fa2, 'type', 1);

hwp.HAction.Execute('CellBorderFill', hwp.HParameterSet.HCellBorderFill.HSet);
console.log('  WinBrushFaceStyle=1 설정');

// 현재 값 확인
hwp.HAction.GetDefault('CellBorderFill', hwp.HParameterSet.HCellBorderFill.HSet);
var hps3 = bridge.comGet(h, 'HParameterSet');
var hcbf3 = bridge.comGet(hps3, 'HCellBorderFill');
var fa3 = bridge.comGet(hcbf3, 'FillAttr');
console.log('  읽기: WinBrushFaceColor=' + bridge.comGet(fa3, 'WinBrushFaceColor'));
console.log('  읽기: WindowsBrush=' + bridge.comGet(fa3, 'WindowsBrush'));
console.log('  읽기: WinBrushFaceStyle=' + bridge.comGet(fa3, 'WinBrushFaceStyle'));
console.log('  읽기: type=' + bridge.comGet(fa3, 'type'));

console.log('\n확인해주세요!');
