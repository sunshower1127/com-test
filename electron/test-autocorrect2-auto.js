/**
 * 자동교정 디버그 2차 — 자동 (비대화형)
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

function check(result) {
  return result.indexOf(' ') === -1 ? '✅ 교정 안 됨' : '⚠️ 교정됨';
}

// ── T1: BreakPara 전후 ──
console.log('── T1: BreakPara 전후 비교');
clear();
insert('첫번째');
console.log('  BreakPara 전: "' + getText() + '"');
bridge.comCallWith(h, 'Run', ['BreakPara']);
console.log('  BreakPara 후: "' + getText().replace(/\r?\n/g, '|') + '"');

// ── T2: MoveDocBegin ──
console.log('\n── T2: MoveDocBegin');
clear();
insert('첫번째');
bridge.comCallWith(h, 'Run', ['MoveDocBegin']);
var r = getText();
console.log('  결과: "' + r + '" ' + check(r));

// ── T3: MoveDocEnd ──
console.log('\n── T3: MoveDocEnd');
clear();
insert('첫번째');
bridge.comCallWith(h, 'Run', ['MoveDocEnd']);
r = getText();
console.log('  결과: "' + r + '" ' + check(r));

// ── T4: MovePos(2) ──
console.log('\n── T4: MovePos(2)');
clear();
insert('첫번째');
bridge.comCallWith(h, 'MovePos', [2, 0, 0]);
r = getText();
console.log('  결과: "' + r + '" ' + check(r));

// ── T5: Cancel ──
console.log('\n── T5: Cancel');
clear();
insert('첫번째');
bridge.comCallWith(h, 'Run', ['Cancel']);
r = getText();
console.log('  결과: "' + r + '" ' + check(r));

// ── T6: MoveLeft ──
console.log('\n── T6: MoveLeft');
clear();
insert('첫번째');
bridge.comCallWith(h, 'Run', ['MoveLeft']);
r = getText();
console.log('  결과: "' + r + '" ' + check(r));

// ── T7: MoveRight ──
console.log('\n── T7: MoveRight');
clear();
insert('첫번째');
bridge.comCallWith(h, 'Run', ['MoveRight']);
r = getText();
console.log('  결과: "' + r + '" ' + check(r));

// ── T8: \r\n으로 줄바꿈 (BreakPara 없이) ──
console.log('\n── T8: InsertText에 \\r\\n');
clear();
insert('첫번째\r\n두번째\r\n세번째');
r = getText();
console.log('  결과: "' + r.replace(/\r?\n/g, '|') + '" ' + check(r));

// ── T9: AutoSpellCheck 토글 후 BreakPara ──
console.log('\n── T9: AutoSpellCheck 토글 → BreakPara');
bridge.comCallWith(h, 'Run', ['AutoSpellCheck']);
clear();
insert('첫번째');
bridge.comCallWith(h, 'Run', ['BreakPara']);
insert('두번째');
r = getText();
console.log('  결과: "' + r.replace(/\r?\n/g, '|') + '" ' + check(r));

// ── T10: 한 번 더 토글 (원래대로) → BreakPara ──
console.log('\n── T10: AutoSpellCheck 다시 토글 → BreakPara');
bridge.comCallWith(h, 'Run', ['AutoSpellCheck']);
clear();
insert('첫번째');
bridge.comCallWith(h, 'Run', ['BreakPara']);
insert('두번째');
r = getText();
console.log('  결과: "' + r.replace(/\r?\n/g, '|') + '" ' + check(r));

// ── T11: ToggleAutoCorrect 토글 → BreakPara ──
console.log('\n── T11: ToggleAutoCorrect 토글 → BreakPara');
bridge.comCallWith(h, 'Run', ['ToggleAutoCorrect']);
clear();
insert('첫번째');
bridge.comCallWith(h, 'Run', ['BreakPara']);
insert('두번째');
r = getText();
console.log('  결과: "' + r.replace(/\r?\n/g, '|') + '" ' + check(r));

// ── T12: 다시 토글 (원래대로) ──
console.log('\n── T12: ToggleAutoCorrect 다시 토글 → BreakPara');
bridge.comCallWith(h, 'Run', ['ToggleAutoCorrect']);
clear();
insert('첫번째');
bridge.comCallWith(h, 'Run', ['BreakPara']);
insert('두번째');
r = getText();
console.log('  결과: "' + r.replace(/\r?\n/g, '|') + '" ' + check(r));

// 종료
console.log('\n── 완료');
bridge.comCallWith(h, 'Clear', [1]);
bridge.comCallWith(h, 'Quit', []);
