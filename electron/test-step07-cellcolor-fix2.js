/**
 * 셀 배경색 — 성공한 매크로와 동일한 프로퍼티 조합
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

// 1x2 표
hwp.HAction.GetDefault('TableCreate', hwp.HParameterSet.HTableCreation.HSet);
hwp.HParameterSet.HTableCreation.Rows = 1;
hwp.HParameterSet.HTableCreation.Cols = 2;
hwp.HParameterSet.HTableCreation.WidthType = 2;
hwp.HAction.Execute('TableCreate', hwp.HParameterSet.HTableCreation.HSet);

// A1
insert('A1');

// A1 배경색 — 성공한 매크로와 동일하게
hwp.HAction.GetDefault('CellBorderFill', hwp.HParameterSet.HCellBorderFill.HSet);
var hcbf = hwp.HParameterSet.HCellBorderFill;
var fa = hcbf.FillAttr;
fa.type = 1;
fa.WindowsBrush = 1;
fa.WinBrushFaceColor = bridge.comCallWith(h, 'RGBColor', [255, 0, 0]);
fa.WinBrushFaceStyle = -1;
fa.WinBrushHatchColor = bridge.comCallWith(h, 'RGBColor', [0, 0, 0]);
var ret1 = hwp.HAction.Execute('CellBorderFill', hwp.HParameterSet.HCellBorderFill.HSet);
console.log('A1 ret=' + ret1);

// B1
bridge.comCallWith(h, 'Run', ['TableRightCell']);
insert('B1');

// B1 — comPut 대신 Proxy 프로퍼티 접근으로
hwp.HAction.GetDefault('CellBorderFill', hwp.HParameterSet.HCellBorderFill.HSet);
hcbf = hwp.HParameterSet.HCellBorderFill;
fa = hcbf.FillAttr;
fa.type = 1;
fa.WindowsBrush = 1;
fa.WinBrushFaceColor = bridge.comCallWith(h, 'RGBColor', [0, 0, 255]);
fa.WinBrushFaceStyle = -1;
fa.WinBrushHatchColor = bridge.comCallWith(h, 'RGBColor', [0, 0, 0]);
var ret2 = hwp.HAction.Execute('CellBorderFill', hwp.HParameterSet.HCellBorderFill.HSet);
console.log('B1 ret=' + ret2);

console.log('\n확인해주세요!');
