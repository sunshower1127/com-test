/**
 * 자동교정 6차: 레지스트리 00000277 플래그 조작 + \r\n 우회 확인
 */

var bridge = require('./native/com_bridge_node.node');
var { initBridge, createComProxy } = require('./dist/worker/proxy');
var { execSync } = require('child_process');

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

function testBreakPara(label) {
  clear();
  insert('첫번째');
  bridge.comCallWith(h, 'Run', ['BreakPara']);
  insert('끝');
  var r = getText();
  var ok = r.indexOf('첫번째') >= 0;
  console.log('  ' + label + ': "' + r + '" ' + (ok ? '✅' : '⚠️'));
  return ok;
}

console.log('HWP Version: ' + String(hwp.Version) + '\n');

// ── T1: 현재 레지스트리 값 읽기 ──
console.log('── T1: Parameters\\00000277 현재 값');
var regPath = 'HKCU\\SOFTWARE\\HNC\\Hwp\\10.0\\Parameters\\00000277';
try {
  var out = execSync('reg query "' + regPath + '"', { encoding: 'utf8' });
  console.log(out);
} catch(e) { console.log('  ❌ ' + e.message); }

// ── T2: 모든 플래그를 0으로 → HWP 재시작 없이 테스트 ──
console.log('── T2: 플래그 0으로 변경 (HWP 이미 실행 중 — 효과 없을 수 있음)');
for (var i = 0; i <= 6; i++) {
  var valName = '0000700' + i;
  try {
    execSync('reg add "' + regPath + '" /v ' + valName + ' /t REG_DWORD /d 0 /f', { encoding: 'utf8' });
  } catch(e) {}
}
console.log('  레지스트리 변경 완료');
testBreakPara('실행 중 변경');

// ── T3: 플래그 복구 ──
console.log('\n── T3: 플래그 원래 값으로 복구');
var originals = { '00007000': 1, '00007001': 1, '00007002': 0, '00007003': 0, '00007004': 1, '00007005': 1, '00007006': 1 };
Object.keys(originals).forEach(function(k) {
  try {
    execSync('reg add "' + regPath + '" /v ' + k + ' /t REG_DWORD /d ' + originals[k] + ' /f', { encoding: 'utf8' });
  } catch(e) {}
});
console.log('  복구 완료');

// ── T4: \r\n 우회로 여러 문단 만들기 ──
console.log('\n── T4: \\r\\n 우회 (BreakPara 없이 여러 문단)');
clear();
insert('첫번째\r\n두번째\r\n세번째\r\n네번째\r\n다섯번째');
var r = getText();
console.log('  결과: "' + r + '"');
console.log('  ' + (r.indexOf(' ') === -1 ? '✅ 교정 안 됨' : '⚠️ 교정됨'));

// ── T5: \r\n 문단 후 추가 삽입 ──
console.log('\n── T5: \\r\\n 문단 후 MoveDocEnd → 추가 InsertText');
bridge.comCallWith(h, 'Run', ['MoveDocEnd']);
insert('\r\n여섯번째');
r = getText();
console.log('  결과: "' + r + '"');
console.log('  ' + (r.indexOf(' ') === -1 ? '✅ 교정 안 됨' : '⚠️ 교정됨'));

// ── T6: \r\n으로 만든 문단에서 BreakPara하면? ──
console.log('\n── T6: \\r\\n으로 만든 문단에서 BreakPara');
clear();
insert('첫번째\r\n두번째');
console.log('  BreakPara 전: "' + getText() + '"');
// 첫번째 문단 끝으로 이동 후 BreakPara
bridge.comCallWith(h, 'MovePos', [1, 0, 0]);
bridge.comCallWith(h, 'MovePos', [7, 0, 0]);
bridge.comCallWith(h, 'Run', ['BreakPara']);
console.log('  첫문단에서 BreakPara: "' + getText() + '"');

// 종료
console.log('\n── 완료');
bridge.comCallWith(h, 'Clear', [1]);
bridge.comCallWith(h, 'Quit', []);
