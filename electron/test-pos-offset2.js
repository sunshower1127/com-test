/**
 * pos 오프셋: para 0 = 16, para 1+ = 0 검증
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

// BreakPara로 3문단
insert('ABC');
bridge.comCallWith(h, 'Run', ['BreakPara']);
insert('DEF');
bridge.comCallWith(h, 'Run', ['BreakPara']);
insert('GHI');

console.log('텍스트: ABC | DEF | GHI\n');

// Para 0: offset 16
console.log('=== Para 0 (offset 16) ===');
for (var i = 0; i < 3; i++) {
  bridge.comCallWith(h, 'SetPos', [0, 0, 16 + i]);
  var ch = peekAfter(1);
  var expected = 'ABC'[i];
  console.log('  SetPos(0,0,' + (16+i) + ') → "' + ch + '" ' + (ch === expected ? '✅' : '❌'));
}

// Para 1: offset 0
console.log('\n=== Para 1 (offset 0) ===');
for (var i = 0; i < 3; i++) {
  bridge.comCallWith(h, 'SetPos', [0, 1, i]);
  var ch = peekAfter(1);
  var expected = 'DEF'[i];
  console.log('  SetPos(0,1,' + i + ') → "' + ch + '" ' + (ch === expected ? '✅' : '❌'));
}

// Para 2: offset 0
console.log('\n=== Para 2 (offset 0) ===');
for (var i = 0; i < 3; i++) {
  bridge.comCallWith(h, 'SetPos', [0, 2, i]);
  var ch = peekAfter(1);
  var expected = 'GHI'[i];
  console.log('  SetPos(0,2,' + i + ') → "' + ch + '" ' + (ch === expected ? '✅' : '❌'));
}

// MovePos도 동일한지
console.log('\n=== MovePos 검증 ===');
bridge.comCallWith(h, 'MovePos', [1, 0, 17]);
console.log('  MovePos(1,0,17) → "' + peekAfter(1) + '" (기대: B)');
bridge.comCallWith(h, 'MovePos', [1, 1, 1]);
console.log('  MovePos(1,1,1) → "' + peekAfter(1) + '" (기대: E)');
bridge.comCallWith(h, 'MovePos', [1, 2, 2]);
console.log('  MovePos(1,2,2) → "' + peekAfter(1) + '" (기대: I)');

// SelectText도 검증
console.log('\n=== SelectText 여러 문단 ===');
bridge.comCallWith(h, 'SelectText', [0, 16+1, 0, 16+2]);
var sel = String(bridge.comCallWith(h, 'GetTextFile', ['UNICODE', 'saveblock']));
bridge.comCallWith(h, 'Run', ['Cancel']);
console.log('  Para0 "B": "' + sel + '" ' + (sel === 'B' ? '✅' : '❌'));

bridge.comCallWith(h, 'SelectText', [1, 0, 1, 2]);
sel = String(bridge.comCallWith(h, 'GetTextFile', ['UNICODE', 'saveblock']));
bridge.comCallWith(h, 'Run', ['Cancel']);
console.log('  Para1 "DE": "' + sel + '" ' + (sel === 'DE' ? '✅' : '❌'));

bridge.comCallWith(h, 'SelectText', [2, 1, 2, 3]);
sel = String(bridge.comCallWith(h, 'GetTextFile', ['UNICODE', 'saveblock']));
bridge.comCallWith(h, 'Run', ['Cancel']);
console.log('  Para2 "HI": "' + sel + '" ' + (sel === 'HI' ? '✅' : '❌'));

// 종료
console.log('\n완료');
bridge.comCallWith(h, 'Clear', [1]);
bridge.comCallWith(h, 'Quit', []);
