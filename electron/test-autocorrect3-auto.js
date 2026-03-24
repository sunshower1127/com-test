/**
 * 자동교정 디버그 3차: BreakPara가 어느 범위를 교정하는지
 */

var bridge = require('./native/com_bridge_node.node');
var { initBridge, createComProxy } = require('./dist/worker/proxy');

bridge.comInit();
initBridge(bridge);

var h = bridge.comCreate('HWPFrame.HwpObject');
var hwp = createComProxy(h);
var win0 = bridge.comCallWith(bridge.comGet(h, 'XHwpWindows'), 'Item', [0]);
bridge.comPut(win0, 'Visible', true);

function clear() {
  bridge.comCallWith(h, 'Run', ['SelectAll']);
  bridge.comCallWith(h, 'Run', ['Delete']);
}

function getText() {
  return String(bridge.comCallWith(h, 'GetTextFile', ['UNICODE', ''])).replace(/\r?\n/g, '|');
}

function insert(text) {
  hwp.HAction.GetDefault('InsertText', hwp.HParameterSet.HInsertText.HSet);
  hwp.HParameterSet.HInsertText.Text = text;
  hwp.HAction.Execute('InsertText', hwp.HParameterSet.HInsertText.HSet);
}

// ── T1: 여러 문단 만들어놓고, 마지막에서 BreakPara ──
// 이전 문단들은 이미 교정된 상태 vs 안 된 상태?
console.log('── T1: \\r\\n으로 3문단 만들고 → BreakPara');
clear();
insert('첫번째\r\n두번째\r\n세번째');
console.log('  BreakPara 전: "' + getText() + '"');
bridge.comCallWith(h, 'Run', ['BreakPara']);
console.log('  BreakPara 후: "' + getText() + '"');
console.log('  → 세번째(커서 있던 문단)만 교정? 전부 교정?\n');

// ── T2: 문서 중간에서 BreakPara ──
console.log('── T2: 커서를 두번째 문단 끝으로 → BreakPara');
clear();
insert('첫번째\r\n두번째\r\n세번째');
console.log('  전: "' + getText() + '"');
// 두번째 문단 끝으로 이동
bridge.comCallWith(h, 'MovePos', [1, 1, 0]); // 두번째 문단 처음
bridge.comCallWith(h, 'MovePos', [7, 0, 0]); // 문단 끝
bridge.comCallWith(h, 'Run', ['BreakPara']);
console.log('  후: "' + getText() + '"');
console.log('  → 두번째만? 전부?\n');

// ── T3: 문서 처음에서 BreakPara ──
console.log('── T3: 커서를 첫번째 문단 끝으로 → BreakPara');
clear();
insert('첫번째\r\n두번째\r\n세번째');
console.log('  전: "' + getText() + '"');
bridge.comCallWith(h, 'MovePos', [1, 0, 0]); // 첫번째 문단 처음
bridge.comCallWith(h, 'MovePos', [7, 0, 0]); // 문단 끝
bridge.comCallWith(h, 'Run', ['BreakPara']);
console.log('  후: "' + getText() + '"');
console.log('  → 첫번째만? 전부?\n');

// ── T4: 이미 교정된 문단에서 BreakPara하면 다시 교정하는지 ──
console.log('── T4: 이미 교정된 텍스트에서 BreakPara');
clear();
insert('첫번째');
bridge.comCallWith(h, 'Run', ['BreakPara']); // 교정 발동
insert('두번째');
// 이 시점: "첫 번째|두번째"
console.log('  현재: "' + getText() + '"');
// 첫번째 문단으로 가서 다시 BreakPara
bridge.comCallWith(h, 'MovePos', [1, 0, 0]);
bridge.comCallWith(h, 'MovePos', [7, 0, 0]);
bridge.comCallWith(h, 'Run', ['BreakPara']); // 첫 번째 문단을 또 쪼갬
console.log('  첫문단에서 BreakPara 후: "' + getText() + '"');
console.log('  → "첫 번째"가 더 바뀌는지?\n');

// ── T5: BreakPara 연속 2번 ──
console.log('── T5: BreakPara 연속 2번 (빈 문단 생성)');
clear();
insert('첫번째');
bridge.comCallWith(h, 'Run', ['BreakPara']);
bridge.comCallWith(h, 'Run', ['BreakPara']);
insert('두번째');
console.log('  결과: "' + getText() + '"\n');

// 종료
console.log('── 완료');
bridge.comCallWith(h, 'Clear', [1]);
bridge.comCallWith(h, 'Quit', []);
