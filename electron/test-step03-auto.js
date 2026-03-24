/**
 * 단계 3: 커서 이동 + 선택 — 자동 테스트 (영문 텍스트, 자동교정 회피)
 */

var bridge = require('./native/com_bridge_node.node');
var { initBridge, createComProxy } = require('./dist/worker/proxy');

bridge.comInit();
initBridge(bridge);

var h = bridge.comCreate('HWPFrame.HwpObject');
var hwp = createComProxy(h);
var win0 = bridge.comCallWith(bridge.comGet(h, 'XHwpWindows'), 'Item', [0]);
bridge.comPut(win0, 'Visible', true);

console.log('HWP Version: ' + String(hwp.Version) + '\n');

function insert(text) {
  hwp.HAction.GetDefault('InsertText', hwp.HParameterSet.HInsertText.HSet);
  hwp.HParameterSet.HInsertText.Text = text;
  hwp.HAction.Execute('InsertText', hwp.HParameterSet.HInsertText.HSet);
}

// 텍스트 준비: \r\n 사용 (자동교정 회피)
insert('AAABBBCCC\r\nDDDEEEFFF\r\nGGGHHHIII');
var full = String(bridge.comCallWith(h, 'GetTextFile', ['UNICODE', ''])).replace(/\r?\n/g, '|');
console.log('준비: "' + full + '"\n');

function run(label, fn) {
  process.stdout.write('── ' + label + ': ');
  try { fn(); } catch(e) { console.log('❌ ' + e.message); }
}

// 커서 뒤 N글자
function peekAfter(n) {
  for (var i = 0; i < n; i++) bridge.comCallWith(h, 'Run', ['MoveSelRight']);
  var t = bridge.comCallWith(h, 'GetTextFile', ['UNICODE', 'saveblock']);
  bridge.comCallWith(h, 'Run', ['Cancel']);
  return String(t);
}
// 커서 앞 N글자
function peekBefore(n) {
  for (var i = 0; i < n; i++) bridge.comCallWith(h, 'Run', ['MoveSelLeft']);
  var t = bridge.comCallWith(h, 'GetTextFile', ['UNICODE', 'saveblock']);
  bridge.comCallWith(h, 'Run', ['Cancel']);
  return String(t);
}

// ================================================================
// Run 기반 커서 이동
// ================================================================

run('3-1 MoveDocBegin', function() {
  bridge.comCallWith(h, 'Run', ['MoveDocBegin']);
  var p = peekAfter(3);
  console.log(p === 'AAA' ? '✅' : '❌ got "' + p + '"', '(기대: "AAA")');
});

run('3-2 MoveDocEnd', function() {
  bridge.comCallWith(h, 'Run', ['MoveDocEnd']);
  var p = peekBefore(3);
  console.log(p === 'III' ? '✅' : '❌ got "' + p + '"', '(기대: "III")');
});

run('3-3 MoveLineBegin', function() {
  bridge.comCallWith(h, 'Run', ['MoveDocEnd']);
  bridge.comCallWith(h, 'Run', ['MoveLineBegin']);
  var p = peekAfter(3);
  console.log(p === 'GGG' ? '✅' : '❌ got "' + p + '"', '(기대: "GGG")');
});

run('3-4 MoveLineEnd', function() {
  bridge.comCallWith(h, 'Run', ['MoveDocBegin']);
  bridge.comCallWith(h, 'Run', ['MoveLineEnd']);
  var p = peekBefore(3);
  console.log(p === 'CCC' ? '✅' : '❌ got "' + p + '"', '(기대: "CCC")');
});

run('3-5 MoveNextParaBegin', function() {
  bridge.comCallWith(h, 'Run', ['MoveDocBegin']);
  bridge.comCallWith(h, 'Run', ['MoveNextParaBegin']);
  var p = peekAfter(3);
  console.log(p === 'DDD' ? '✅' : '❌ got "' + p + '"', '(기대: "DDD")');
});

run('3-6 MovePrevParaBegin', function() {
  bridge.comCallWith(h, 'Run', ['MoveDocEnd']);
  bridge.comCallWith(h, 'Run', ['MovePrevParaBegin']);
  var p = peekAfter(3);
  console.log(p === 'GGG' ? '✅' : '❌ got "' + p + '"', '(기대: "GGG")');
});

run('3-7 MoveParaBegin', function() {
  bridge.comCallWith(h, 'MovePos', [1, 1, 5]); // 두번째문단 중간
  bridge.comCallWith(h, 'Run', ['MoveParaBegin']);
  var p = peekAfter(3);
  console.log(p === 'DDD' ? '✅' : '❌ got "' + p + '"', '(기대: "DDD")');
});

run('3-8 MoveParaEnd', function() {
  bridge.comCallWith(h, 'MovePos', [1, 1, 0]); // 두번째문단 처음
  bridge.comCallWith(h, 'Run', ['MoveParaEnd']);
  var p = peekBefore(3);
  console.log(p === 'FFF' ? '✅' : '❌ got "' + p + '"', '(기대: "FFF")');
});

// ================================================================
// Run 기반 선택
// ================================================================

run('3-9 MoveSelDocEnd', function() {
  bridge.comCallWith(h, 'Run', ['MoveDocBegin']);
  bridge.comCallWith(h, 'Run', ['MoveSelDocEnd']);
  var t = bridge.comCallWith(h, 'GetTextFile', ['UNICODE', 'saveblock']);
  bridge.comCallWith(h, 'Run', ['Cancel']);
  var expected = 'AAABBBCCC\r\nDDDEEEFFF\r\nGGGHHHIII';
  console.log(String(t) === expected ? '✅' : '❌', '"' + String(t).replace(/\r?\n/g, '|') + '"');
});

run('3-10 MoveSelLineEnd', function() {
  bridge.comCallWith(h, 'Run', ['MoveDocBegin']);
  bridge.comCallWith(h, 'Run', ['MoveSelLineEnd']);
  var t = String(bridge.comCallWith(h, 'GetTextFile', ['UNICODE', 'saveblock']));
  bridge.comCallWith(h, 'Run', ['Cancel']);
  console.log(t === 'AAABBBCCC' ? '✅' : '❌ got "' + t + '"', '(기대: "AAABBBCCC")');
});

run('3-11 MoveSelNextWord', function() {
  bridge.comCallWith(h, 'Run', ['MoveDocBegin']);
  bridge.comCallWith(h, 'Run', ['MoveSelNextWord']);
  var t = String(bridge.comCallWith(h, 'GetTextFile', ['UNICODE', 'saveblock']));
  bridge.comCallWith(h, 'Run', ['Cancel']);
  console.log('선택: "' + t + '"');
});

run('3-12 MoveSelRight x3', function() {
  bridge.comCallWith(h, 'Run', ['MoveDocBegin']);
  var t = peekAfter(3); // peekAfter does MoveSelRight + saveblock + Cancel
  // but peekAfter already cancels... redo properly
  bridge.comCallWith(h, 'Run', ['MoveDocBegin']);
  for (var i = 0; i < 3; i++) bridge.comCallWith(h, 'Run', ['MoveSelRight']);
  var t2 = String(bridge.comCallWith(h, 'GetTextFile', ['UNICODE', 'saveblock']));
  bridge.comCallWith(h, 'Run', ['Cancel']);
  console.log(t2 === 'AAA' ? '✅' : '❌ got "' + t2 + '"', '(기대: "AAA")');
});

// ================================================================
// MovePos API
// ================================================================

run('3-13 MovePos(2) = 문서 처음', function() {
  var ret = bridge.comCallWith(h, 'MovePos', [2, 0, 0]);
  var p = peekAfter(3);
  console.log('ret=' + ret + ', ' + (p === 'AAA' ? '✅' : '❌ "' + p + '"'));
});

run('3-14 MovePos(3) = 문서 끝', function() {
  var ret = bridge.comCallWith(h, 'MovePos', [3, 0, 0]);
  var p = peekBefore(3);
  console.log('ret=' + ret + ', ' + (p === 'III' ? '✅' : '❌ "' + p + '"'));
});

run('3-15 MovePos(1,1,0) = 2번째 문단 처음', function() {
  var ret = bridge.comCallWith(h, 'MovePos', [1, 1, 0]);
  var p = peekAfter(3);
  console.log('ret=' + ret + ', ' + (p === 'DDD' ? '✅' : '❌ "' + p + '"'));
});

run('3-16 MovePos(1,0,3) = 1번째 문단 4번째', function() {
  var ret = bridge.comCallWith(h, 'MovePos', [1, 0, 3]);
  var p = peekAfter(3);
  console.log('ret=' + ret + ', ' + (p === 'BBB' ? '✅' : '❌ "' + p + '"') + ' (기대: "BBB")');
});

run('3-17 MovePos(6) = 현재 문단 처음', function() {
  bridge.comCallWith(h, 'MovePos', [1, 1, 5]);
  var ret = bridge.comCallWith(h, 'MovePos', [6, 0, 0]);
  var p = peekAfter(3);
  console.log('ret=' + ret + ', ' + (p === 'DDD' ? '✅' : '❌ "' + p + '"'));
});

run('3-18 MovePos(7) = 현재 문단 끝', function() {
  bridge.comCallWith(h, 'MovePos', [1, 1, 0]);
  var ret = bridge.comCallWith(h, 'MovePos', [7, 0, 0]);
  var p = peekBefore(3);
  console.log('ret=' + ret + ', ' + (p === 'FFF' ? '✅' : '❌ "' + p + '"'));
});

// ================================================================
// SetPos / GetPosBySet / SetPosBySet
// ================================================================

run('3-19 SetPos(0,2,0) = 3번째 문단', function() {
  var ret = bridge.comCallWith(h, 'SetPos', [0, 2, 0]);
  var p = peekAfter(3);
  console.log('ret=' + ret + ', ' + (p === 'GGG' ? '✅' : '❌ "' + p + '"'));
});

run('3-20 GetPosBySet → SetPosBySet', function() {
  bridge.comCallWith(h, 'MovePos', [1, 1, 3]); // 2번째 문단 4번째
  var posSet = bridge.comCallWith(h, 'GetPosBySet', []);
  console.log('type=' + typeof posSet);
  bridge.comCallWith(h, 'Run', ['MoveDocBegin']); // 다른 곳으로
  var ret = bridge.comCallWith(h, 'SetPosBySet', [posSet]);
  var p = peekAfter(3);
  console.log('         ret=' + ret + ', ' + (p === 'EEE' ? '✅' : '❌ "' + p + '"') + ' (기대: "EEE")');
});

// ================================================================
// SelectText → 조작
// ================================================================

run('3-21 SelectText(0,0,0,3) → Delete', function() {
  bridge.comCallWith(h, 'Run', ['MoveDocBegin']);
  var sel = bridge.comCallWith(h, 'SelectText', [0, 0, 0, 3]);
  bridge.comCallWith(h, 'Run', ['Delete']);
  var t = String(bridge.comCallWith(h, 'GetTextFile', ['UNICODE', ''])).replace(/\r?\n/g, '|');
  var ok = t.indexOf('BBBCCC') === 0;
  console.log('sel=' + sel + ', ' + (ok ? '✅' : '❌') + ' "' + t.substring(0, 30) + '..."');
  bridge.comCallWith(h, 'Run', ['Undo']);
});

run('3-22 SelectText(0,0,0,3) → InsertText 덮어쓰기', function() {
  bridge.comCallWith(h, 'Run', ['MoveDocBegin']);
  bridge.comCallWith(h, 'SelectText', [0, 0, 0, 3]);
  insert('XXX');
  var t = String(bridge.comCallWith(h, 'GetTextFile', ['UNICODE', ''])).replace(/\r?\n/g, '|');
  var ok = t.indexOf('XXXBBBCCC') === 0;
  console.log(ok ? '✅' : '❌', '"' + t.substring(0, 30) + '..."');
  bridge.comCallWith(h, 'Run', ['Undo']);
  bridge.comCallWith(h, 'Run', ['Undo']); // InsertText + Delete 두 번
});

// ================================================================
// GetSelectedPosBySet
// ================================================================

run('3-23 GetSelectedPosBySet', function() {
  bridge.comCallWith(h, 'Run', ['MoveDocBegin']);
  for (var i = 0; i < 3; i++) bridge.comCallWith(h, 'Run', ['MoveSelRight']);

  // 방법 1: CreateSet("SePos")
  try {
    var sset = bridge.comCallWith(h, 'CreateSet', ['SePos']);
    var eset = bridge.comCallWith(h, 'CreateSet', ['SePos']);
    var ret = bridge.comCallWith(h, 'GetSelectedPosBySet', [sset, eset]);
    console.log('ret=' + ret);
    // Item으로 읽기
    try {
      var sp = bridge.comCallWith(sset, 'Item', ['Para']);
      var ss = bridge.comCallWith(sset, 'Item', ['Pos']);
      var ep = bridge.comCallWith(eset, 'Item', ['Para']);
      var es = bridge.comCallWith(eset, 'Item', ['Pos']);
      console.log('         start=(' + sp + ',' + ss + ') end=(' + ep + ',' + es + ')');
    } catch(e2) {
      console.log('         Item 실패: ' + e2.message);
      // SetID 확인
      try {
        var sid = bridge.comGet(sset, 'SetID');
        console.log('         SetID=' + sid);
        // listMembers
        var members = bridge.comListMembers(sset);
        var names = members.map(function(m) { return m.name; }).filter(function(n) {
          return ['QueryInterface','AddRef','Release','GetTypeInfoCount','GetTypeInfo','GetIDsOfNames','Invoke'].indexOf(n) === -1;
        });
        console.log('         멤버: ' + names.join(', '));
      } catch(e3) {}
    }
  } catch(e) { console.log('❌ ' + e.message); }
  bridge.comCallWith(h, 'Run', ['Cancel']);
});

// ================================================================
// Undo / Redo
// ================================================================

run('3-24 Undo', function() {
  var before = String(bridge.comCallWith(h, 'GetTextFile', ['UNICODE', ''])).replace(/\r?\n/g, '|');
  bridge.comCallWith(h, 'Run', ['MoveDocBegin']);
  for (var i = 0; i < 3; i++) bridge.comCallWith(h, 'Run', ['MoveSelRight']);
  bridge.comCallWith(h, 'Run', ['Delete']);
  var after = String(bridge.comCallWith(h, 'GetTextFile', ['UNICODE', ''])).replace(/\r?\n/g, '|');
  bridge.comCallWith(h, 'Run', ['Undo']);
  var restored = String(bridge.comCallWith(h, 'GetTextFile', ['UNICODE', ''])).replace(/\r?\n/g, '|');
  console.log(restored === before ? '✅ Undo 정상' : '❌ 복구 안 됨');
});

// 종료
console.log('\n── 완료');
bridge.comCallWith(h, 'Clear', [1]);
bridge.comCallWith(h, 'Quit', []);
