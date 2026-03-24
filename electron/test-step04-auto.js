/**
 * 단계 4: 텍스트 수정 — 자동 테스트
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
function run(label, fn) {
  process.stdout.write('── ' + label + ': ');
  try { fn(); } catch(e) { console.log('❌ ' + e.message); }
}

console.log('HWP Version: ' + String(hwp.Version) + '\n');

// ================================================================
// 4-1: Run 기반 선택 → Delete
// ================================================================
run('4-1 MoveSelRight x3 → Delete', function() {
  reset('ABCDEFGHIJ');
  bridge.comCallWith(h, 'Run', ['MoveDocBegin']);
  for (var i = 0; i < 3; i++) bridge.comCallWith(h, 'Run', ['MoveSelRight']);
  bridge.comCallWith(h, 'Run', ['Delete']);
  var t = getText();
  console.log(t === 'DEFGHIJ' ? '✅' : '❌ "' + t + '"', '(기대: "DEFGHIJ")');
});

// ================================================================
// 4-2: Run 기반 선택 → InsertText 덮어쓰기
// ================================================================
run('4-2 MoveSelRight x3 → InsertText', function() {
  reset('ABCDEFGHIJ');
  bridge.comCallWith(h, 'Run', ['MoveDocBegin']);
  for (var i = 0; i < 3; i++) bridge.comCallWith(h, 'Run', ['MoveSelRight']);
  insert('XXX');
  var t = getText();
  console.log(t === 'XXXDEFGHIJ' ? '✅' : '❌ "' + t + '"', '(기대: "XXXDEFGHIJ")');
});

// ================================================================
// 4-3: 특정 위치 텍스트 교체 (MoveRight + MoveSelRight + InsertText)
// ================================================================
run('4-3 위치 지정 교체 (4~6번째 글자)', function() {
  reset('ABCDEFGHIJ');
  bridge.comCallWith(h, 'Run', ['MoveDocBegin']);
  for (var i = 0; i < 3; i++) bridge.comCallWith(h, 'Run', ['MoveRight']);
  for (var i = 0; i < 3; i++) bridge.comCallWith(h, 'Run', ['MoveSelRight']);
  insert('xxx');
  var t = getText();
  console.log(t === 'ABCxxxGHIJ' ? '✅' : '❌ "' + t + '"', '(기대: "ABCxxxGHIJ")');
});

// ================================================================
// 4-4: AllReplace (찾기/바꾸기) — HParameterSet 패턴
// ================================================================
run('4-4 AllReplace (HParameterSet)', function() {
  reset('AABBCCAABBCC');
  bridge.comCallWith(h, 'Run', ['MoveDocBegin']);

  hwp.HAction.GetDefault('AllReplace', hwp.HParameterSet.HFindReplace.HSet);
  hwp.HParameterSet.HFindReplace.FindString = 'BB';
  hwp.HParameterSet.HFindReplace.ReplaceString = 'XX';
  hwp.HParameterSet.HFindReplace.IgnoreMessage = 1;
  hwp.HParameterSet.HFindReplace.ReplaceMode = 1;
  var ret = hwp.HAction.Execute('AllReplace', hwp.HParameterSet.HFindReplace.HSet);
  var t = getText();
  console.log('ret=' + ret + ', ' + (t === 'AAXXCCAAXXCC' ? '✅' : '❌ "' + t + '"'), '(기대: "AAXXCCAAXXCC")');
});

// ================================================================
// 4-5: AllReplace — CreateAction 패턴
// ================================================================
run('4-5 AllReplace (CreateAction)', function() {
  reset('HelloWorldHello');
  bridge.comCallWith(h, 'Run', ['MoveDocBegin']);

  var act = hwp.CreateAction('AllReplace');
  var set = act.CreateSet();
  act.GetDefault(set);
  set.SetItem('FindString', 'Hello');
  set.SetItem('ReplaceString', 'Hi');
  set.SetItem('IgnoreMessage', 1);
  set.SetItem('ReplaceMode', 1);
  var ret = act.Execute(set);
  var t = getText();
  console.log('ret=' + ret + ', ' + (t === 'HiWorldHi' ? '✅' : '❌ "' + t + '"'), '(기대: "HiWorldHi")');
});

// ================================================================
// 4-6: AllReplace — 한글 텍스트
// ================================================================
run('4-6 AllReplace 한글', function() {
  reset('가나다라마가나다');
  bridge.comCallWith(h, 'Run', ['MoveDocBegin']);

  hwp.HAction.GetDefault('AllReplace', hwp.HParameterSet.HFindReplace.HSet);
  hwp.HParameterSet.HFindReplace.FindString = '가나다';
  hwp.HParameterSet.HFindReplace.ReplaceString = 'XYZ';
  hwp.HParameterSet.HFindReplace.IgnoreMessage = 1;
  hwp.HParameterSet.HFindReplace.ReplaceMode = 1;
  hwp.HAction.Execute('AllReplace', hwp.HParameterSet.HFindReplace.HSet);
  var t = getText();
  console.log(t === 'XYZ라마XYZ' ? '✅' : '❌ "' + t + '"', '(기대: "XYZ라마XYZ")');
});

// ================================================================
// 4-7: AllReplace — 여러 문단에 걸쳐
// ================================================================
run('4-7 AllReplace 여러 문단', function() {
  reset('AAA111BBB\r\nCCC111DDD\r\nEEE111FFF');
  bridge.comCallWith(h, 'Run', ['MoveDocBegin']);

  hwp.HAction.GetDefault('AllReplace', hwp.HParameterSet.HFindReplace.HSet);
  hwp.HParameterSet.HFindReplace.FindString = '111';
  hwp.HParameterSet.HFindReplace.ReplaceString = '___';
  hwp.HParameterSet.HFindReplace.IgnoreMessage = 1;
  hwp.HParameterSet.HFindReplace.ReplaceMode = 1;
  hwp.HAction.Execute('AllReplace', hwp.HParameterSet.HFindReplace.HSet);
  var t = getText();
  console.log(t === 'AAA___BBB|CCC___DDD|EEE___FFF' ? '✅' : '❌ "' + t + '"');
});

// ================================================================
// 4-8: SetTextFile 전체 교체 패턴 안정성
// ================================================================
run('4-8 SelectAll+Delete+SetTextFile 전체 교체', function() {
  reset('원본 텍스트');
  bridge.comCallWith(h, 'Run', ['SelectAll']);
  bridge.comCallWith(h, 'Run', ['Delete']);
  bridge.comCallWith(h, 'SetTextFile', ['완전히 새로운 내용', 'UNICODE', '']);
  var t = getText();
  console.log(t === '완전히 새로운 내용' ? '✅' : '❌ "' + t + '"');
});

// ================================================================
// 4-9: Undo
// ================================================================
run('4-9 Undo', function() {
  reset('BEFORE');
  bridge.comCallWith(h, 'Run', ['SelectAll']);
  bridge.comCallWith(h, 'Run', ['Delete']);
  insert('AFTER');
  console.log('삭제 후: "' + getText() + '"');
  bridge.comCallWith(h, 'Run', ['Undo']); // InsertText 취소
  bridge.comCallWith(h, 'Run', ['Undo']); // Delete 취소
  var t = getText();
  process.stdout.write('  Undo 후: "' + t + '" ');
  console.log(t === 'BEFORE' ? '✅' : '❌');
});

// ================================================================
// 4-10: Redo
// ================================================================
run('4-10 Redo', function() {
  reset('ABCDEF');
  bridge.comCallWith(h, 'Run', ['MoveDocBegin']);
  for (var i = 0; i < 3; i++) bridge.comCallWith(h, 'Run', ['MoveSelRight']);
  bridge.comCallWith(h, 'Run', ['Delete']);
  var after = getText(); // "DEF"
  bridge.comCallWith(h, 'Run', ['Undo']);
  var undone = getText(); // "ABCDEF"
  bridge.comCallWith(h, 'Run', ['Redo']);
  var redone = getText(); // "DEF"
  console.log('삭제="' + after + '", Undo="' + undone + '", Redo="' + redone + '" ' +
    (redone === after ? '✅' : '❌'));
});

// ================================================================
// 4-11: 빈 문자열로 바꾸기 (삭제 효과)
// ================================================================
run('4-11 AllReplace 빈 문자열 (삭제)', function() {
  reset('AAA---BBB---CCC');
  bridge.comCallWith(h, 'Run', ['MoveDocBegin']);

  hwp.HAction.GetDefault('AllReplace', hwp.HParameterSet.HFindReplace.HSet);
  hwp.HParameterSet.HFindReplace.FindString = '---';
  hwp.HParameterSet.HFindReplace.ReplaceString = '';
  hwp.HParameterSet.HFindReplace.IgnoreMessage = 1;
  hwp.HParameterSet.HFindReplace.ReplaceMode = 1;
  hwp.HAction.Execute('AllReplace', hwp.HParameterSet.HFindReplace.HSet);
  var t = getText();
  console.log(t === 'AAABBBCCC' ? '✅' : '❌ "' + t + '"');
});

// 종료
console.log('\n── 완료');
bridge.comCallWith(h, 'Clear', [1]);
bridge.comCallWith(h, 'Quit', []);
