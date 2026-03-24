/**
 * 단계 3: 커서 이동 + 선택 — 대화형 (한 단계씩)
 *
 * 사용법: cd electron && node test-step03.js
 */

var bridge = require('./native/com_bridge_node.node');
var { initBridge, createComProxy } = require('./dist/worker/proxy');
var readline = require('readline');

bridge.comInit();
initBridge(bridge);

var h = bridge.comCreate('HWPFrame.HwpObject');
var hwp = createComProxy(h);
var win0 = bridge.comCallWith(bridge.comGet(h, 'XHwpWindows'), 'Item', [0]);
bridge.comPut(win0, 'Visible', true);

console.log('HWP Version: ' + String(hwp.Version));

var rl = readline.createInterface({ input: process.stdin, output: process.stdout });
function ask(msg) {
  return new Promise(function(resolve) {
    rl.question('\n' + msg + ' → Enter ', function() { resolve(); });
  });
}

// 커서 뒤 N글자 확인 (선택 후 saveblock)
function peekAtCursor(count) {
  count = count || 5;
  for (var i = 0; i < count; i++) {
    bridge.comCallWith(h, 'Run', ['MoveSelRight']);
  }
  var text = bridge.comCallWith(h, 'GetTextFile', ['UNICODE', 'saveblock']);
  bridge.comCallWith(h, 'Run', ['Cancel']);
  return String(text);
}

// 커서 앞 N글자 확인
function peekBefore(count) {
  count = count || 3;
  for (var i = 0; i < count; i++) {
    bridge.comCallWith(h, 'Run', ['MoveSelLeft']);
  }
  var text = bridge.comCallWith(h, 'GetTextFile', ['UNICODE', 'saveblock']);
  bridge.comCallWith(h, 'Run', ['Cancel']);
  return String(text);
}

(async function() {

  // 텍스트 준비
  hwp.HAction.GetDefault('InsertText', hwp.HParameterSet.HInsertText.HSet);
  hwp.HParameterSet.HInsertText.Text = '첫번째문단ABCDE';
  hwp.HAction.Execute('InsertText', hwp.HParameterSet.HInsertText.HSet);
  bridge.comCallWith(h, 'Run', ['BreakPara']);
  hwp.HAction.GetDefault('InsertText', hwp.HParameterSet.HInsertText.HSet);
  hwp.HParameterSet.HInsertText.Text = '두번째문단12345';
  hwp.HAction.Execute('InsertText', hwp.HParameterSet.HInsertText.HSet);
  bridge.comCallWith(h, 'Run', ['BreakPara']);
  hwp.HAction.GetDefault('InsertText', hwp.HParameterSet.HInsertText.HSet);
  hwp.HParameterSet.HInsertText.Text = '세번째문단FGHIJ';
  hwp.HAction.Execute('InsertText', hwp.HParameterSet.HInsertText.HSet);

  console.log('\n준비: "첫번째문단ABCDE | 두번째문단12345 | 세번째문단FGHIJ"\n');

  // ── Run 기반 이동 ──

  await ask('── 3-1: MoveDocBegin');
  try {
    bridge.comCallWith(h, 'Run', ['MoveDocBegin']);
    console.log('  ✅ 커서 뒤 3글자: "' + peekAtCursor(3) + '" (기대: "첫번째")');
  } catch(e) { console.log('  ❌ ' + e.message); }

  await ask('── 3-2: MoveDocEnd');
  try {
    bridge.comCallWith(h, 'Run', ['MoveDocEnd']);
    console.log('  ✅ 커서 앞 3글자: "' + peekBefore(3) + '" (기대: "HIJ")');
  } catch(e) { console.log('  ❌ ' + e.message); }

  await ask('── 3-3: MoveLineBegin (마지막 줄에서)');
  try {
    bridge.comCallWith(h, 'Run', ['MoveDocEnd']);
    bridge.comCallWith(h, 'Run', ['MoveLineBegin']);
    console.log('  ✅ 커서 뒤 3글자: "' + peekAtCursor(3) + '" (기대: "세번째")');
  } catch(e) { console.log('  ❌ ' + e.message); }

  await ask('── 3-4: MoveLineEnd (첫 줄에서)');
  try {
    bridge.comCallWith(h, 'Run', ['MoveDocBegin']);
    bridge.comCallWith(h, 'Run', ['MoveLineEnd']);
    console.log('  ✅ 커서 앞 3글자: "' + peekBefore(3) + '" (기대: "CDE")');
  } catch(e) { console.log('  ❌ ' + e.message); }

  await ask('── 3-5: MoveNextParaBegin (문서 처음에서)');
  try {
    bridge.comCallWith(h, 'Run', ['MoveDocBegin']);
    bridge.comCallWith(h, 'Run', ['MoveNextParaBegin']);
    console.log('  ✅ 커서 뒤 3글자: "' + peekAtCursor(3) + '" (기대: "두번째")');
  } catch(e) { console.log('  ❌ ' + e.message); }

  await ask('── 3-6: MovePrevParaBegin (문서 끝에서)');
  try {
    bridge.comCallWith(h, 'Run', ['MoveDocEnd']);
    bridge.comCallWith(h, 'Run', ['MovePrevParaBegin']);
    console.log('  ✅ 커서 뒤 3글자: "' + peekAtCursor(3) + '" (기대: "세번째")');
  } catch(e) { console.log('  ❌ ' + e.message); }

  // ── Run 기반 선택 ──

  await ask('── 3-7: MoveSelDocEnd (전체 선택)');
  try {
    bridge.comCallWith(h, 'Run', ['MoveDocBegin']);
    bridge.comCallWith(h, 'Run', ['MoveSelDocEnd']);
    var text = bridge.comCallWith(h, 'GetTextFile', ['UNICODE', 'saveblock']);
    bridge.comCallWith(h, 'Run', ['Cancel']);
    console.log('  ✅ 선택: "' + String(text) + '"');
  } catch(e) { console.log('  ❌ ' + e.message); }

  await ask('── 3-8: MoveSelLineEnd (첫줄 선택)');
  try {
    bridge.comCallWith(h, 'Run', ['MoveDocBegin']);
    bridge.comCallWith(h, 'Run', ['MoveSelLineEnd']);
    var text = bridge.comCallWith(h, 'GetTextFile', ['UNICODE', 'saveblock']);
    bridge.comCallWith(h, 'Run', ['Cancel']);
    console.log('  ✅ 선택: "' + String(text) + '" (기대: "첫번째문단ABCDE")');
  } catch(e) { console.log('  ❌ ' + e.message); }

  await ask('── 3-9: MoveSelNextWord (첫 단어)');
  try {
    bridge.comCallWith(h, 'Run', ['MoveDocBegin']);
    bridge.comCallWith(h, 'Run', ['MoveSelNextWord']);
    var text = bridge.comCallWith(h, 'GetTextFile', ['UNICODE', 'saveblock']);
    bridge.comCallWith(h, 'Run', ['Cancel']);
    console.log('  ✅ 선택: "' + String(text) + '"');
  } catch(e) { console.log('  ❌ ' + e.message); }

  // ── MovePos API ──

  await ask('── 3-10: MovePos(2) = 문서 처음');
  try {
    var ret = bridge.comCallWith(h, 'MovePos', [2, 0, 0]);
    console.log('  ret=' + String(ret) + ', 커서 뒤: "' + peekAtCursor(3) + '" (기대: "첫번째")');
  } catch(e) { console.log('  ❌ ' + e.message); }

  await ask('── 3-11: MovePos(3) = 문서 끝');
  try {
    var ret = bridge.comCallWith(h, 'MovePos', [3, 0, 0]);
    console.log('  ret=' + String(ret) + ', 커서 앞: "' + peekBefore(3) + '" (기대: "HIJ")');
  } catch(e) { console.log('  ❌ ' + e.message); }

  await ask('── 3-12: MovePos(1, 1, 0) = 두번째 문단 처음');
  try {
    var ret = bridge.comCallWith(h, 'MovePos', [1, 1, 0]);
    console.log('  ret=' + String(ret) + ', 커서 뒤: "' + peekAtCursor(3) + '" (기대: "두번째")');
  } catch(e) { console.log('  ❌ ' + e.message); }

  await ask('── 3-13: MovePos(1, 0, 3) = 첫문단 4번째 글자');
  try {
    var ret = bridge.comCallWith(h, 'MovePos', [1, 0, 3]);
    console.log('  ret=' + String(ret) + ', 커서 뒤: "' + peekAtCursor(3) + '" (기대: "째문단")');
  } catch(e) { console.log('  ❌ ' + e.message); }

  await ask('── 3-14: MovePos(6) = 현재 문단 처음');
  try {
    bridge.comCallWith(h, 'MovePos', [1, 1, 5]);
    var ret = bridge.comCallWith(h, 'MovePos', [6, 0, 0]);
    console.log('  ret=' + String(ret) + ', 커서 뒤: "' + peekAtCursor(3) + '" (기대: "두번째")');
  } catch(e) { console.log('  ❌ ' + e.message); }

  await ask('── 3-15: MovePos(7) = 현재 문단 끝');
  try {
    bridge.comCallWith(h, 'MovePos', [1, 1, 0]);
    var ret = bridge.comCallWith(h, 'MovePos', [7, 0, 0]);
    console.log('  ret=' + String(ret) + ', 커서 앞: "' + peekBefore(3) + '" (기대: "345")');
  } catch(e) { console.log('  ❌ ' + e.message); }

  // ── SetPos / GetPosBySet ──

  await ask('── 3-16: SetPos(0, 2, 0) = 세번째 문단');
  try {
    var ret = bridge.comCallWith(h, 'SetPos', [0, 2, 0]);
    console.log('  ret=' + String(ret) + ', 커서 뒤: "' + peekAtCursor(3) + '" (기대: "세번째")');
  } catch(e) { console.log('  ❌ ' + e.message); }

  await ask('── 3-17: GetPosBySet → SetPosBySet');
  try {
    bridge.comCallWith(h, 'MovePos', [1, 1, 3]);
    var posSet = bridge.comCallWith(h, 'GetPosBySet', []);
    console.log('  GetPosBySet type=' + typeof posSet);
    bridge.comCallWith(h, 'Run', ['MoveDocBegin']);
    var ret = bridge.comCallWith(h, 'SetPosBySet', [posSet]);
    console.log('  SetPosBySet ret=' + String(ret) + ', 커서 뒤: "' + peekAtCursor(3) + '" (기대: "째문단")');
  } catch(e) { console.log('  ❌ ' + e.message); }

  // ── SelectText + Delete ──

  await ask('── 3-18: SelectText(0,0,0,5) → Delete');
  try {
    bridge.comCallWith(h, 'Run', ['MoveDocBegin']);
    var sel = bridge.comCallWith(h, 'SelectText', [0, 0, 0, 5]);
    console.log('  SelectText ret=' + String(sel));
    bridge.comCallWith(h, 'Run', ['Delete']);
    var text = bridge.comCallWith(h, 'GetTextFile', ['UNICODE', '']);
    console.log('  삭제 후: "' + String(text).substring(0, 30) + '..." (기대: "문단ABCDE...")');
    bridge.comCallWith(h, 'Run', ['Undo']);
    console.log('  Undo 완료');
  } catch(e) { console.log('  ❌ ' + e.message); }

  // ── GetSelectedPosBySet ──

  await ask('── 3-19: GetSelectedPosBySet');
  try {
    bridge.comCallWith(h, 'Run', ['MoveDocBegin']);
    for (var i = 0; i < 3; i++) bridge.comCallWith(h, 'Run', ['MoveSelRight']);
    var sset = bridge.comCallWith(h, 'CreateSet', ['SePos']);
    var eset = bridge.comCallWith(h, 'CreateSet', ['SePos']);
    var ret = bridge.comCallWith(h, 'GetSelectedPosBySet', [sset, eset]);
    console.log('  ret=' + String(ret));
    if (ret) {
      try {
        var sp = bridge.comCallWith(sset, 'Item', ['Para']);
        var ss = bridge.comCallWith(sset, 'Item', ['Pos']);
        var ep = bridge.comCallWith(eset, 'Item', ['Para']);
        var es = bridge.comCallWith(eset, 'Item', ['Pos']);
        console.log('  start=(' + sp + ',' + ss + ') end=(' + ep + ',' + es + ') (기대: (0,0)→(0,3))');
      } catch(e2) {
        console.log('  Item 접근 실패: ' + e2.message);
      }
    }
    bridge.comCallWith(h, 'Run', ['Cancel']);
  } catch(e) { console.log('  ❌ ' + e.message); }

  // ── 종료 ──
  await ask('── 테스트 완료! 종료');
  bridge.comCallWith(h, 'Clear', [1]);
  bridge.comCallWith(h, 'Quit', []);
  rl.close();
  process.exit(0);

})();
