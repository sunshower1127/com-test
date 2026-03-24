/**
 * 단계 5: 서식 적용 결과 보여주기 (닫지 않음)
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

function applyCharShape(propName, value) {
  hwp.HAction.GetDefault('CharShape', hwp.HParameterSet.HCharShape.HSet);
  hwp.HParameterSet.HCharShape[propName] = value;
  hwp.HAction.Execute('CharShape', hwp.HParameterSet.HCharShape.HSet);
}

function selectRange(start, count) {
  bridge.comCallWith(h, 'Run', ['MoveDocBegin']);
  for (var i = 0; i < start; i++) bridge.comCallWith(h, 'Run', ['MoveRight']);
  for (var i = 0; i < count; i++) bridge.comCallWith(h, 'Run', ['MoveSelRight']);
}

// 텍스트 준비
insert('ABC DEF GHI JKL MNO PQR');
insert('\r\n');
insert('STYLED TEXT BELOW');

// ABC — 24pt
selectRange(0, 3);
applyCharShape('Height', 2400);
bridge.comCallWith(h, 'Run', ['Cancel']);

// DEF — 볼드
selectRange(4, 3);
applyCharShape('Bold', 1);
bridge.comCallWith(h, 'Run', ['Cancel']);

// GHI — 이탤릭
selectRange(8, 3);
applyCharShape('Italic', 1);
bridge.comCallWith(h, 'Run', ['Cancel']);

// JKL — 빨간색
selectRange(12, 3);
applyCharShape('TextColor', bridge.comCallWith(h, 'RGBColor', [255, 0, 0]));
bridge.comCallWith(h, 'Run', ['Cancel']);

// MNO — 밑줄
selectRange(16, 3);
applyCharShape('UnderlineType', 1);
bridge.comCallWith(h, 'Run', ['Cancel']);

// PQR — 취소선
selectRange(20, 3);
applyCharShape('StrikeOutType', 1);
bridge.comCallWith(h, 'Run', ['Cancel']);

// 두번째 줄 — 32pt 볼드 파란색
bridge.comCallWith(h, 'SetPos', [0, 1, 0]);
for (var i = 0; i < 17; i++) bridge.comCallWith(h, 'Run', ['MoveSelRight']);
hwp.HAction.GetDefault('CharShape', hwp.HParameterSet.HCharShape.HSet);
hwp.HParameterSet.HCharShape.Height = 3200;
hwp.HParameterSet.HCharShape.Bold = 1;
hwp.HParameterSet.HCharShape.TextColor = bridge.comCallWith(h, 'RGBColor', [0, 0, 255]);
hwp.HAction.Execute('CharShape', hwp.HParameterSet.HCharShape.HSet);
bridge.comCallWith(h, 'Run', ['Cancel']);

console.log('서식 적용 완료! 한글 창에서 확인해주세요.');
console.log('');
console.log('ABC = 24pt');
console.log('DEF = 볼드');
console.log('GHI = 이탤릭');
console.log('JKL = 빨간색');
console.log('MNO = 밑줄');
console.log('PQR = 취소선');
console.log('STYLED TEXT BELOW = 32pt 볼드 파란색');
console.log('');
console.log('(한글 창 열려있음 — 수동으로 닫으세요)');
