/**
 * pos 오프셋 16 검증: 언제나 16인가?
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
function readPosAt() {
  // HWP 매크로에서는 GetPosBySet().Item("Pos")가 되지만
  // 브릿지에서는 null이므로, MoveSelRight로 위치 확인
}

// ================================================================
// T1: BreakPara로 만든 문단에서 각 문단의 pos 시작값
// ================================================================
console.log('=== T1: BreakPara로 만든 3개 문단 ===');
insert('ABC');
bridge.comCallWith(h, 'Run', ['BreakPara']);
insert('DEF');
bridge.comCallWith(h, 'Run', ['BreakPara']);
insert('GHI');

// 각 문단에서 SetPos(0, para, 16+0) 테스트
for (var para = 0; para < 3; para++) {
  bridge.comCallWith(h, 'SetPos', [0, para, 16]);
  var ch = peekAfter(1);
  var expected = ['A', 'D', 'G'][para];
  console.log('  para=' + para + ' SetPos(0,' + para + ',16) → "' + ch + '" ' +
    (ch === expected ? '✅' : '❌ expected "' + expected + '"'));

  // pos=17도 확인
  bridge.comCallWith(h, 'SetPos', [0, para, 17]);
  var ch2 = peekAfter(1);
  var expected2 = ['B', 'E', 'H'][para];
  console.log('  para=' + para + ' SetPos(0,' + para + ',17) → "' + ch2 + '" ' +
    (ch2 === expected2 ? '✅' : '❌ expected "' + expected2 + '"'));
}

// ================================================================
// T2: \r\n으로 만든 문단에서 동일한지
// ================================================================
console.log('\n=== T2: \\r\\n으로 만든 3개 문단 ===');
bridge.comCallWith(h, 'Run', ['SelectAll']);
bridge.comCallWith(h, 'Run', ['Delete']);
insert('ABC\r\nDEF\r\nGHI');

for (var para = 0; para < 3; para++) {
  bridge.comCallWith(h, 'SetPos', [0, para, 16]);
  var ch = peekAfter(1);
  var expected = ['A', 'D', 'G'][para];
  console.log('  para=' + para + ' SetPos(0,' + para + ',16) → "' + ch + '" ' +
    (ch === expected ? '✅' : '❌ expected "' + expected + '"'));
}

// ================================================================
// T3: 빈 문서 → FileNew → 새 문단에서도 16인지
// ================================================================
console.log('\n=== T3: FileNew 후 새 문서 ===');
bridge.comCallWith(h, 'Clear', [1]);
bridge.comCallWith(h, 'Run', ['FileNew']);
insert('XYZ123');
bridge.comCallWith(h, 'SetPos', [0, 0, 16]);
var ch = peekAfter(1);
console.log('  SetPos(0,0,16) → "' + ch + '" ' + (ch === 'X' ? '✅' : '❌'));
bridge.comCallWith(h, 'SetPos', [0, 0, 19]);
ch = peekAfter(1);
console.log('  SetPos(0,0,19) → "' + ch + '" ' + (ch === '1' ? '✅' : '❌'));

// ================================================================
// T4: 긴 텍스트 (100글자 이상)에서도 동일한지
// ================================================================
console.log('\n=== T4: 긴 텍스트 (알파벳 반복) ===');
bridge.comCallWith(h, 'Run', ['SelectAll']);
bridge.comCallWith(h, 'Run', ['Delete']);
var longText = '';
for (var i = 0; i < 100; i++) longText += String.fromCharCode(65 + (i % 26));
insert(longText); // ABCDEFG...ABCDEFG... 100글자

bridge.comCallWith(h, 'SetPos', [0, 0, 16 + 0]);
console.log('  pos=16 → "' + peekAfter(1) + '" (기대: A)');
bridge.comCallWith(h, 'SetPos', [0, 0, 16 + 50]);
console.log('  pos=66 → "' + peekAfter(1) + '" (기대: ' + String.fromCharCode(65 + 50%26) + ')');
bridge.comCallWith(h, 'SetPos', [0, 0, 16 + 99]);
console.log('  pos=115 → "' + peekAfter(1) + '" (기대: ' + String.fromCharCode(65 + 99%26) + ')');

// ================================================================
// T5: 특수문자/이모지/혼합에서도 1:1인지
// ================================================================
console.log('\n=== T5: 특수문자/혼합 ===');
bridge.comCallWith(h, 'Run', ['SelectAll']);
bridge.comCallWith(h, 'Run', ['Delete']);
insert('A가1!@한B');

var expected5 = ['A', '가', '1', '!', '@', '한', 'B'];
for (var i = 0; i < expected5.length; i++) {
  bridge.comCallWith(h, 'SetPos', [0, 0, 16 + i]);
  var ch = peekAfter(1);
  console.log('  pos=' + (16+i) + ' → "' + ch + '" ' +
    (ch === expected5[i] ? '✅' : '❌ expected "' + expected5[i] + '"'));
}

// ================================================================
// T6: SelectText + Delete가 진짜 되는지
// ================================================================
console.log('\n=== T6: SelectText(16+N) → Delete ===');
bridge.comCallWith(h, 'Run', ['SelectAll']);
bridge.comCallWith(h, 'Run', ['Delete']);
insert('ABCDEFGHIJ');

bridge.comCallWith(h, 'SelectText', [0, 16+3, 0, 16+6]); // DEF 선택
bridge.comCallWith(h, 'Run', ['Delete']);
var result = String(bridge.comCallWith(h, 'GetTextFile', ['UNICODE', '']));
console.log('  삭제 후: "' + result + '" ' + (result === 'ABCGHIJ' ? '✅' : '❌'));

// 종료
console.log('\n완료');
bridge.comCallWith(h, 'Clear', [1]);
bridge.comCallWith(h, 'Quit', []);
