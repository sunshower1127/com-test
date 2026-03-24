/**
 * 셀 배경색 — 변수 하나씩 격리 테스트
 *
 * 성공 조건: fa.type=1, fa.WindowsBrush=1, fa.WinBrushFaceColor,
 *            fa.WinBrushFaceStyle=-1, fa.WinBrushHatchColor
 *            + Proxy 접근
 *
 * 테스트:
 * A1: 전부 다 (기준 — 이게 되어야 함)
 * B1: WinBrushHatchColor 빼기
 * C1: WinBrushFaceStyle 안 넣기 (기본값 0)
 * A2: WinBrushFaceStyle = 1 (이전 실패값)
 * B2: WinBrushFaceStyle = 0
 * C2: comPut으로 접근 (Proxy 대신)
 * A3: comPut + WinBrushHatchColor + FaceStyle=-1 (comPut만 다름)
 * B3: Proxy + HatchColor 없이 + FaceStyle=-1
 * C3: Proxy + HatchColor 있고 + FaceStyle=1
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

var red = function() { return bridge.comCallWith(h, 'RGBColor', [255, 0, 0]); };
var black = function() { return bridge.comCallWith(h, 'RGBColor', [0, 0, 0]); };

// 3x3 표
hwp.HAction.GetDefault('TableCreate', hwp.HParameterSet.HTableCreation.HSet);
hwp.HParameterSet.HTableCreation.Rows = 3;
hwp.HParameterSet.HTableCreation.Cols = 3;
hwp.HParameterSet.HTableCreation.WidthType = 2;
hwp.HAction.Execute('TableCreate', hwp.HParameterSet.HTableCreation.HSet);

function nextCell() {
  bridge.comCallWith(h, 'Run', ['TableRightCell']);
}

// ─── A1: 전부 다 (기준) ───
insert('A1:ALL');
hwp.HAction.GetDefault('CellBorderFill', hwp.HParameterSet.HCellBorderFill.HSet);
var hcbf = hwp.HParameterSet.HCellBorderFill;
var fa = hcbf.FillAttr;
fa.type = 1;
fa.WindowsBrush = 1;
fa.WinBrushFaceColor = red();
fa.WinBrushFaceStyle = -1;
fa.WinBrushHatchColor = black();
hwp.HAction.Execute('CellBorderFill', hwp.HParameterSet.HCellBorderFill.HSet);
console.log('A1: ALL (기준)');

// ─── B1: HatchColor 빼기 ───
nextCell();
insert('B1:no Hatch');
hwp.HAction.GetDefault('CellBorderFill', hwp.HParameterSet.HCellBorderFill.HSet);
hcbf = hwp.HParameterSet.HCellBorderFill;
fa = hcbf.FillAttr;
fa.type = 1;
fa.WindowsBrush = 1;
fa.WinBrushFaceColor = red();
fa.WinBrushFaceStyle = -1;
// WinBrushHatchColor 생략
hwp.HAction.Execute('CellBorderFill', hwp.HParameterSet.HCellBorderFill.HSet);
console.log('B1: no HatchColor');

// ─── C1: FaceStyle 안 넣기 ───
nextCell();
insert('C1:no Style');
hwp.HAction.GetDefault('CellBorderFill', hwp.HParameterSet.HCellBorderFill.HSet);
hcbf = hwp.HParameterSet.HCellBorderFill;
fa = hcbf.FillAttr;
fa.type = 1;
fa.WindowsBrush = 1;
fa.WinBrushFaceColor = red();
// WinBrushFaceStyle 생략 (기본값 0)
fa.WinBrushHatchColor = black();
hwp.HAction.Execute('CellBorderFill', hwp.HParameterSet.HCellBorderFill.HSet);
console.log('C1: no FaceStyle');

// ─── A2: FaceStyle = 1 ───
nextCell();
insert('A2:Style=1');
hwp.HAction.GetDefault('CellBorderFill', hwp.HParameterSet.HCellBorderFill.HSet);
hcbf = hwp.HParameterSet.HCellBorderFill;
fa = hcbf.FillAttr;
fa.type = 1;
fa.WindowsBrush = 1;
fa.WinBrushFaceColor = red();
fa.WinBrushFaceStyle = 1;
fa.WinBrushHatchColor = black();
hwp.HAction.Execute('CellBorderFill', hwp.HParameterSet.HCellBorderFill.HSet);
console.log('A2: FaceStyle=1');

// ─── B2: FaceStyle = 0 ───
nextCell();
insert('B2:Style=0');
hwp.HAction.GetDefault('CellBorderFill', hwp.HParameterSet.HCellBorderFill.HSet);
hcbf = hwp.HParameterSet.HCellBorderFill;
fa = hcbf.FillAttr;
fa.type = 1;
fa.WindowsBrush = 1;
fa.WinBrushFaceColor = red();
fa.WinBrushFaceStyle = 0;
fa.WinBrushHatchColor = black();
hwp.HAction.Execute('CellBorderFill', hwp.HParameterSet.HCellBorderFill.HSet);
console.log('B2: FaceStyle=0');

// ─── C2: comPut으로 접근 (전부 다 설정) ───
nextCell();
insert('C2:comPut');
hwp.HAction.GetDefault('CellBorderFill', hwp.HParameterSet.HCellBorderFill.HSet);
var hps = bridge.comGet(h, 'HParameterSet');
var hcbf2 = bridge.comGet(hps, 'HCellBorderFill');
var fa2 = bridge.comGet(hcbf2, 'FillAttr');
bridge.comPut(fa2, 'type', 1);
bridge.comPut(fa2, 'WindowsBrush', 1);
bridge.comPut(fa2, 'WinBrushFaceColor', red());
bridge.comPut(fa2, 'WinBrushFaceStyle', -1);
bridge.comPut(fa2, 'WinBrushHatchColor', black());
hwp.HAction.Execute('CellBorderFill', hwp.HParameterSet.HCellBorderFill.HSet);
console.log('C2: comPut (all props)');

// ─── A3: comPut + 전부 다 + Execute도 comCallWith ───
nextCell();
insert('A3:comPut+comExec');
hwp.HAction.GetDefault('CellBorderFill', hwp.HParameterSet.HCellBorderFill.HSet);
hps = bridge.comGet(h, 'HParameterSet');
hcbf2 = bridge.comGet(hps, 'HCellBorderFill');
fa2 = bridge.comGet(hcbf2, 'FillAttr');
bridge.comPut(fa2, 'type', 1);
bridge.comPut(fa2, 'WindowsBrush', 1);
bridge.comPut(fa2, 'WinBrushFaceColor', red());
bridge.comPut(fa2, 'WinBrushFaceStyle', -1);
bridge.comPut(fa2, 'WinBrushHatchColor', black());
// Execute도 bridge 직접
var ha = bridge.comGet(h, 'HAction');
var hset = bridge.comGet(hcbf2, 'HSet');
bridge.comCallWith(ha, 'Execute', ['CellBorderFill', hset]);
console.log('A3: comPut + comExec');

// ─── B3: Proxy + HatchColor 없이 + FaceStyle=-1 ───
nextCell();
insert('B3:no Hatch,-1');
hwp.HAction.GetDefault('CellBorderFill', hwp.HParameterSet.HCellBorderFill.HSet);
hcbf = hwp.HParameterSet.HCellBorderFill;
fa = hcbf.FillAttr;
fa.type = 1;
fa.WindowsBrush = 1;
fa.WinBrushFaceColor = red();
fa.WinBrushFaceStyle = -1;
hwp.HAction.Execute('CellBorderFill', hwp.HParameterSet.HCellBorderFill.HSet);
console.log('B3: Proxy, no Hatch, Style=-1');

// ─── C3: Proxy + HatchColor + FaceStyle=1 ───
nextCell();
insert('C3:Hatch,Style=1');
hwp.HAction.GetDefault('CellBorderFill', hwp.HParameterSet.HCellBorderFill.HSet);
hcbf = hwp.HParameterSet.HCellBorderFill;
fa = hcbf.FillAttr;
fa.type = 1;
fa.WindowsBrush = 1;
fa.WinBrushFaceColor = red();
fa.WinBrushFaceStyle = 1;
fa.WinBrushHatchColor = black();
hwp.HAction.Execute('CellBorderFill', hwp.HParameterSet.HCellBorderFill.HSet);
console.log('C3: Proxy, Hatch, Style=1');

console.log('\n확인해주세요!');
console.log('A1=ALL(기준) B1=noHatch C1=noStyle');
console.log('A2=Style=1   B2=Style=0  C2=comPut');
console.log('A3=comPut+Ex B3=noHatch  C3=Style=1+Hatch');
