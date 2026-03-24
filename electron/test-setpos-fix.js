/**
 * SetPos pos=16+N 공식 검증
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
function peekAfter(n) {
  for (var i = 0; i < n; i++) bridge.comCallWith(h, 'Run', ['MoveSelRight']);
  var t = bridge.comCallWith(h, 'GetTextFile', ['UNICODE', 'saveblock']);
  bridge.comCallWith(h, 'Run', ['Cancel']);
  return String(t);
}

insert('ABCDEFGHIJ');
console.log('텍스트: "ABCDEFGHIJ"\n');

console.log('=== SetPos(0, 0, 16+N) ===');
for (var i = 0; i <= 10; i++) {
  bridge.comCallWith(h, 'SetPos', [0, 0, 16 + i]);
  var ch = peekAfter(1);
  var expected = 'ABCDEFGHIJ'[i] || '(end)';
  var ok = ch === expected;
  console.log('  N=' + i + ' pos=' + (16+i) + ' → "' + ch + '" ' + (ok ? '✅' : '❌ expected "' + expected + '"'));
}

// 한글 테스트
console.log('\n=== 한글: SetPos(0, 0, 16+N) ===');
bridge.comCallWith(h, 'Run', ['SelectAll']);
bridge.comCallWith(h, 'Run', ['Delete']);
insert('가나다라마바사');

for (var i = 0; i <= 7; i++) {
  bridge.comCallWith(h, 'SetPos', [0, 0, 16 + i]);
  var ch = peekAfter(1);
  console.log('  N=' + i + ' pos=' + (16+i) + ' → "' + ch + '"');
}

// 여러 문단 테스트
console.log('\n=== 여러 문단: SetPos(0, para, 16+N) ===');
bridge.comCallWith(h, 'Run', ['SelectAll']);
bridge.comCallWith(h, 'Run', ['Delete']);
insert('ABC\r\nDEF\r\nGHI');

for (var para = 0; para < 3; para++) {
  for (var i = 0; i < 3; i++) {
    bridge.comCallWith(h, 'SetPos', [0, para, 16 + i]);
    var ch = peekAfter(1);
    console.log('  para=' + para + ' N=' + i + ' pos=' + (16+i) + ' → "' + ch + '"');
  }
}

// MovePos도 테스트
console.log('\n=== MovePos(1, 0, 16+N) ===');
bridge.comCallWith(h, 'Run', ['SelectAll']);
bridge.comCallWith(h, 'Run', ['Delete']);
insert('ABCDEFGHIJ');

for (var i = 0; i <= 5; i++) {
  bridge.comCallWith(h, 'MovePos', [1, 0, 16 + i]);
  var ch = peekAfter(1);
  console.log('  N=' + i + ' pos=' + (16+i) + ' → "' + ch + '"');
}

console.log('\n=== SelectText(0, 16+2, 0, 16+5) — "CDE" 선택 ===');
bridge.comCallWith(h, 'Run', ['MoveDocBegin']);
var sel = bridge.comCallWith(h, 'SelectText', [0, 16+2, 0, 16+5]);
console.log('SelectText ret=' + sel);
var text = bridge.comCallWith(h, 'GetTextFile', ['UNICODE', 'saveblock']);
console.log('saveblock: "' + String(text) + '" (기대: "CDE")');
bridge.comCallWith(h, 'Run', ['Cancel']);

// 종료
console.log('\n완료');
bridge.comCallWith(h, 'Clear', [1]);
bridge.comCallWith(h, 'Quit', []);
