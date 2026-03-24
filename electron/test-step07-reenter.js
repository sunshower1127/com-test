/**
 * 표 재진입: MoveDown으로 진입 테스트
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

function readPos() {
  var ps = bridge.comCallWith(h, 'GetPosBySet', []);
  return {
    list: bridge.comCallWith(ps, 'Item', ['List']),
    para: bridge.comCallWith(ps, 'Item', ['Para']),
    pos: bridge.comCallWith(ps, 'Item', ['Pos']),
  };
}

function readCellText() {
  bridge.comCallWith(h, 'Run', ['MoveLineBegin']);
  bridge.comCallWith(h, 'Run', ['MoveSelLineEnd']);
  var t = String(bridge.comCallWith(h, 'GetTextFile', ['UNICODE', 'saveblock']));
  bridge.comCallWith(h, 'Run', ['Cancel']);
  return t;
}

// 표 생성 + 채우기
hwp.HAction.GetDefault('TableCreate', hwp.HParameterSet.HTableCreation.HSet);
hwp.HParameterSet.HTableCreation.Rows = 2;
hwp.HParameterSet.HTableCreation.Cols = 2;
hwp.HParameterSet.HTableCreation.WidthType = 2;
hwp.HAction.Execute('TableCreate', hwp.HParameterSet.HTableCreation.HSet);

var cells = ['A1', 'B1', 'A2', 'B2'];
for (var i = 0; i < 4; i++) {
  insert(cells[i]);
  if (i < 3) bridge.comCallWith(h, 'Run', ['TableRightCell']);
}
bridge.comCallWith(h, 'Run', ['MoveDocEnd']);
console.log('표 생성 완료. 표 밖에 있음.\n');

// ── MoveDown으로 재진입 ──
console.log('=== MoveDocBegin → MoveDown ===');
bridge.comCallWith(h, 'Run', ['MoveDocBegin']);
bridge.comCallWith(h, 'Run', ['MoveDown']);
var p = readPos();
var t = readCellText();
console.log('  L=' + p.list + ', P=' + p.para + ', pos=' + p.pos + ', 셀="' + t + '"');

// ── 셀 순회 ──
console.log('\n=== 셀 순회 ===');
bridge.comCallWith(h, 'Run', ['MoveDocBegin']);
bridge.comCallWith(h, 'Run', ['MoveDown']);

// 현재 셀
t = readCellText();
p = readPos();
console.log('  진입 셀: L=' + p.list + ' "' + t + '"');

bridge.comCallWith(h, 'Run', ['TableRightCell']);
t = readCellText();
p = readPos();
console.log('  →Right: L=' + p.list + ' "' + t + '"');

bridge.comCallWith(h, 'Run', ['TableRightCell']);
t = readCellText();
p = readPos();
console.log('  →Right: L=' + p.list + ' "' + t + '"');

bridge.comCallWith(h, 'Run', ['TableRightCell']);
t = readCellText();
p = readPos();
console.log('  →Right: L=' + p.list + ' "' + t + '"');

// ── 텍스트 수정 ──
console.log('\n=== A1 텍스트 수정 ===');
bridge.comCallWith(h, 'Run', ['MoveDocEnd']);
bridge.comCallWith(h, 'Run', ['MoveDocBegin']);
bridge.comCallWith(h, 'Run', ['MoveDown']);
bridge.comCallWith(h, 'Run', ['MoveLineBegin']);
bridge.comCallWith(h, 'Run', ['MoveSelLineEnd']);
bridge.comCallWith(h, 'Run', ['Delete']);
insert('MODIFIED');
t = readCellText();
console.log('  수정 후: "' + t + '"');

// ── 배경색도 적용 ──
console.log('\n=== 재진입 후 배경색 적용 ===');
hwp.HAction.GetDefault('CellBorderFill', hwp.HParameterSet.HCellBorderFill.HSet);
var hcbf = hwp.HParameterSet.HCellBorderFill;
var fa = hcbf.FillAttr;
fa.type = 1;
fa.WindowsBrush = 1;
fa.WinBrushFaceColor = bridge.comCallWith(h, 'RGBColor', [255, 200, 200]);
fa.WinBrushFaceStyle = -1;
hwp.HAction.Execute('CellBorderFill', hwp.HParameterSet.HCellBorderFill.HSet);
console.log('  분홍 배경 적용');

// B1에도
bridge.comCallWith(h, 'Run', ['TableRightCell']);
hwp.HAction.GetDefault('CellBorderFill', hwp.HParameterSet.HCellBorderFill.HSet);
hcbf = hwp.HParameterSet.HCellBorderFill;
fa = hcbf.FillAttr;
fa.type = 1;
fa.WindowsBrush = 1;
fa.WinBrushFaceColor = bridge.comCallWith(h, 'RGBColor', [200, 200, 255]);
fa.WinBrushFaceStyle = -1;
hwp.HAction.Execute('CellBorderFill', hwp.HParameterSet.HCellBorderFill.HSet);
console.log('  파란 배경 적용');

console.log('\n확인해주세요!');
