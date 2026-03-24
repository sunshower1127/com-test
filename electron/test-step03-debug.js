/**
 * 단계 3 디버그: MovePos pos 해석 + GetSelectedPosBySet
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
console.log('텍스트: "ABCDEFGHIJ" (10글자)\n');

// ── MovePos(1, 0, pos) 에서 pos 값별로 커서 위치 확인 ──
console.log('── MovePos(1, 0, pos) — pos별 커서 위치');
for (var pos = 0; pos <= 12; pos++) {
  bridge.comCallWith(h, 'MovePos', [1, 0, pos]);
  var p = peekAfter(1);
  console.log('  pos=' + pos + ' → 뒤 1글자: "' + p + '"');
}

// ── MovePos(0, 0, pos) — moveMain ──
console.log('\n── MovePos(0, 0, pos) — moveMain');
for (var pos = 0; pos <= 12; pos++) {
  try {
    bridge.comCallWith(h, 'MovePos', [0, 0, pos]);
    var p = peekAfter(1);
    console.log('  pos=' + pos + ' → 뒤 1글자: "' + p + '"');
  } catch(e) {
    console.log('  pos=' + pos + ' → ❌ ' + e.message);
  }
}

// ── 한글 텍스트로 pos 확인 (글자당 바이트 다름) ──
console.log('\n── 한글 텍스트 pos 확인');
bridge.comCallWith(h, 'Run', ['SelectAll']);
bridge.comCallWith(h, 'Run', ['Delete']);
insert('가나다라마');
console.log('텍스트: "가나다라마" (5글자)');
for (var pos = 0; pos <= 12; pos++) {
  bridge.comCallWith(h, 'MovePos', [1, 0, pos]);
  var p = peekAfter(1);
  console.log('  pos=' + pos + ' → 뒤 1글자: "' + p + '"');
}

// ── GetSelectedPosBySet — GetPosBySet으로 위치 가져와서 비교 ──
console.log('\n── GetSelectedPosBySet 대안');
bridge.comCallWith(h, 'Run', ['SelectAll']);
bridge.comCallWith(h, 'Run', ['Delete']);
insert('ABCDEFGHIJ');
bridge.comCallWith(h, 'Run', ['MoveDocBegin']);
for (var i = 0; i < 3; i++) bridge.comCallWith(h, 'Run', ['MoveSelRight']);

// 방법: GetPosBySet으로 현재 위치 확인
var posSet = bridge.comCallWith(h, 'GetPosBySet', []);
var members = bridge.comListMembers(posSet);
var names = members.map(function(m) { return m.name; }).filter(function(n) {
  return ['QueryInterface','AddRef','Release','GetTypeInfoCount','GetTypeInfo','GetIDsOfNames','Invoke'].indexOf(n) === -1;
});
console.log('GetPosBySet 멤버: ' + names.join(', '));
names.forEach(function(n) {
  try {
    var v = bridge.comGet(posSet, n);
    console.log('  .' + n + ' = ' + String(v) + ' (' + typeof v + ')');
  } catch(e) {}
});

bridge.comCallWith(h, 'Run', ['Cancel']);

// 종료
console.log('\n── 완료');
bridge.comCallWith(h, 'Clear', [1]);
bridge.comCallWith(h, 'Quit', []);
