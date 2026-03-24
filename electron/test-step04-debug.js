/**
 * 단계 4 디버그: Delete 문제 + 빈 문자열 AllReplace
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
function getText() {
  return String(bridge.comCallWith(h, 'GetTextFile', ['UNICODE', ''])).replace(/\r?\n/g, '|');
}
function reset(text) {
  bridge.comCallWith(h, 'Run', ['SelectAll']);
  bridge.comCallWith(h, 'Run', ['Delete']);
  insert(text);
}

console.log('HWP Version: ' + String(hwp.Version) + '\n');

// ── D1: Delete 단독 (선택 없이) ──
console.log('── D1: Delete 단독 (커서 앞 글자 삭제)');
reset('ABCDEF');
bridge.comCallWith(h, 'Run', ['MoveDocEnd']);
bridge.comCallWith(h, 'Run', ['Delete']); // 커서 뒤 글자 삭제? 아니면 앞?
console.log('  MoveDocEnd + Delete: "' + getText() + '"');

reset('ABCDEF');
bridge.comCallWith(h, 'Run', ['MoveDocBegin']);
bridge.comCallWith(h, 'Run', ['Delete']);
console.log('  MoveDocBegin + Delete: "' + getText() + '" (기대: "BCDEF")');

// ── D2: BackSpace vs Delete ──
console.log('\n── D2: BackSpace');
reset('ABCDEF');
bridge.comCallWith(h, 'Run', ['MoveDocEnd']);
bridge.comCallWith(h, 'Run', ['BackSpace']);
console.log('  MoveDocEnd + BackSpace: "' + getText() + '" (기대: "ABCDE")');

// ── D3: 선택 후 Delete 재시도 (reset 직후) ──
console.log('\n── D3: reset 직후 선택+Delete');
reset('ABCDEFGHIJ');
// reset이 MoveDocBegin을 안 해서 커서가 끝에 있을 수 있음
bridge.comCallWith(h, 'Run', ['MoveDocBegin']);
for (var i = 0; i < 3; i++) bridge.comCallWith(h, 'Run', ['MoveSelRight']);
// saveblock 확인
var sel = bridge.comCallWith(h, 'GetTextFile', ['UNICODE', 'saveblock']);
console.log('  선택된 텍스트: "' + String(sel) + '"');
bridge.comCallWith(h, 'Run', ['Delete']);
console.log('  Delete 후: "' + getText() + '" (기대: "DEFGHIJ")');

// ── D4: 선택 후 Cut (잘라내기) ──
console.log('\n── D4: 선택 후 Cut');
reset('ABCDEFGHIJ');
bridge.comCallWith(h, 'Run', ['MoveDocBegin']);
for (var i = 0; i < 3; i++) bridge.comCallWith(h, 'Run', ['MoveSelRight']);
bridge.comCallWith(h, 'Run', ['Cut']);
console.log('  Cut 후: "' + getText() + '" (기대: "DEFGHIJ")');

// ── D5: SelectAll → Delete ──
console.log('\n── D5: SelectAll → Delete');
reset('ABCDEFGHIJ');
bridge.comCallWith(h, 'Run', ['SelectAll']);
bridge.comCallWith(h, 'Run', ['Delete']);
console.log('  결과: "' + getText() + '" (기대: 빈 문자열)');

// ── D6: AllReplace 빈 문자열 — ReplaceString=" " (공백) ──
console.log('\n── D6: AllReplace 공백으로 치환');
reset('AAA---BBB---CCC');
bridge.comCallWith(h, 'Run', ['MoveDocBegin']);
try {
  hwp.HAction.GetDefault('AllReplace', hwp.HParameterSet.HFindReplace.HSet);
  hwp.HParameterSet.HFindReplace.FindString = '---';
  hwp.HParameterSet.HFindReplace.ReplaceString = ' ';
  hwp.HParameterSet.HFindReplace.IgnoreMessage = 1;
  hwp.HParameterSet.HFindReplace.ReplaceMode = 1;
  hwp.HAction.Execute('AllReplace', hwp.HParameterSet.HFindReplace.HSet);
  console.log('  결과: "' + getText() + '" (기대: "AAA BBB CCC")');
} catch(e) { console.log('  ❌ ' + e.message); }

// ── D7: AllReplace 빈 문자열 — CreateAction 패턴 ──
console.log('\n── D7: AllReplace 빈 문자열 (CreateAction)');
reset('AAA---BBB');
bridge.comCallWith(h, 'Run', ['MoveDocBegin']);
try {
  var act = hwp.CreateAction('AllReplace');
  var set = act.CreateSet();
  act.GetDefault(set);
  set.SetItem('FindString', '---');
  set.SetItem('ReplaceString', '');
  set.SetItem('IgnoreMessage', 1);
  set.SetItem('ReplaceMode', 1);
  var ret = act.Execute(set);
  console.log('  ret=' + ret + ', 결과: "' + getText() + '"');
} catch(e) { console.log('  ❌ ' + e.message); }

// 종료
console.log('\n── 완료');
bridge.comCallWith(h, 'Clear', [1]);
bridge.comCallWith(h, 'Quit', []);
