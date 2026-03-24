/**
 * 자동교정 디버그 2차: 트리거 조건 + 끄는 방법
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

var rl = readline.createInterface({ input: process.stdin, output: process.stdout });
function ask(msg) {
  return new Promise(function(resolve) {
    rl.question('\n' + msg + ' → Enter ', function() { resolve(); });
  });
}

function clear() {
  bridge.comCallWith(h, 'Run', ['SelectAll']);
  bridge.comCallWith(h, 'Run', ['Delete']);
}

function getText() {
  return String(bridge.comCallWith(h, 'GetTextFile', ['UNICODE', '']));
}

function insert(text) {
  hwp.HAction.GetDefault('InsertText', hwp.HParameterSet.HInsertText.HSet);
  hwp.HParameterSet.HInsertText.Text = text;
  hwp.HAction.Execute('InsertText', hwp.HParameterSet.HInsertText.HSet);
}

(async function() {

  console.log('HWP Version: ' + String(hwp.Version) + '\n');

  // ── T1: BreakPara 없이 → GetTextFile로 읽기 → BreakPara → 다시 읽기 ──
  await ask('── T1: BreakPara 전후 비교');
  clear();
  insert('첫번째');
  console.log('  BreakPara 전: "' + getText() + '"');
  bridge.comCallWith(h, 'Run', ['BreakPara']);
  console.log('  BreakPara 후: "' + getText() + '"');

  // ── T2: 다른 트리거들 — MoveDocBegin, MovePos 등 ──
  await ask('── T2: MoveDocBegin이 교정을 트리거하는지');
  clear();
  insert('첫번째');
  console.log('  삽입 직후: "' + getText() + '"');
  bridge.comCallWith(h, 'Run', ['MoveDocBegin']);
  console.log('  MoveDocBegin 후: "' + getText() + '"');

  await ask('── T3: MoveDocEnd가 트리거하는지');
  clear();
  insert('첫번째');
  bridge.comCallWith(h, 'Run', ['MoveDocEnd']);
  console.log('  MoveDocEnd 후: "' + getText() + '"');

  await ask('── T4: MovePos가 트리거하는지');
  clear();
  insert('첫번째');
  bridge.comCallWith(h, 'MovePos', [2, 0, 0]); // 문서 처음
  console.log('  MovePos(2) 후: "' + getText() + '"');

  await ask('── T5: Run("Cancel")이 트리거하는지');
  clear();
  insert('첫번째');
  bridge.comCallWith(h, 'Run', ['Cancel']);
  console.log('  Cancel 후: "' + getText() + '"');

  // ── T6: 자동교정 끄기 — AutoSpellCheck 토글 후 BreakPara ──
  await ask('── T6: AutoSpellCheck 끈 상태에서 BreakPara');
  // AutoSpellCheck 토글 (T6 이전에 한 번 켰으니 다시 꺼야 할 수도)
  bridge.comCallWith(h, 'Run', ['AutoSpellCheck']);
  console.log('  AutoSpellCheck 토글됨');
  clear();
  insert('첫번째');
  bridge.comCallWith(h, 'Run', ['BreakPara']);
  insert('두번째');
  bridge.comCallWith(h, 'Run', ['BreakPara']);
  insert('세번째');
  console.log('  결과: "' + getText() + '"');

  // ── T7: 다시 켜고 확인 ──
  await ask('── T7: AutoSpellCheck 다시 켜고 BreakPara');
  bridge.comCallWith(h, 'Run', ['AutoSpellCheck']);
  console.log('  AutoSpellCheck 토글됨');
  clear();
  insert('첫번째');
  bridge.comCallWith(h, 'Run', ['BreakPara']);
  insert('두번째');
  bridge.comCallWith(h, 'Run', ['BreakPara']);
  insert('세번째');
  console.log('  결과: "' + getText() + '"');

  // ── T8: ToggleAutoCorrect 끈 상태에서 BreakPara ──
  await ask('── T8: ToggleAutoCorrect 끈 상태에서 BreakPara');
  bridge.comCallWith(h, 'Run', ['ToggleAutoCorrect']);
  console.log('  ToggleAutoCorrect 토글됨');
  clear();
  insert('첫번째');
  bridge.comCallWith(h, 'Run', ['BreakPara']);
  insert('두번째');
  bridge.comCallWith(h, 'Run', ['BreakPara']);
  insert('세번째');
  console.log('  결과: "' + getText() + '"');

  // ── T9: 둘 다 끈 상태 ──
  await ask('── T9: AutoSpellCheck + ToggleAutoCorrect 둘 다 끈 상태');
  // 현재 상태 모름 — 일단 둘 다 토글
  bridge.comCallWith(h, 'Run', ['AutoSpellCheck']);
  bridge.comCallWith(h, 'Run', ['ToggleAutoCorrect']);
  console.log('  둘 다 토글됨');
  clear();
  insert('첫번째');
  bridge.comCallWith(h, 'Run', ['BreakPara']);
  insert('두번째');
  bridge.comCallWith(h, 'Run', ['BreakPara']);
  insert('세번째');
  var r = getText();
  console.log('  결과: "' + r + '"');
  if (r.indexOf(' ') === -1) {
    console.log('  ✅ 교정 안 됨!');
  } else {
    console.log('  ⚠️ 여전히 교정됨');
  }

  // ── T10: BreakPara 대신 \r\n으로 줄바꿈 ──
  await ask('── T10: InsertText에 \\r\\n 넣기 (BreakPara 없이)');
  clear();
  insert('첫번째\r\n두번째\r\n세번째');
  console.log('  결과: "' + getText() + '"');

  // ── 종료 ──
  await ask('── 완료! 종료');
  bridge.comCallWith(h, 'Clear', [1]);
  bridge.comCallWith(h, 'Quit', []);
  rl.close();
  process.exit(0);

})();
