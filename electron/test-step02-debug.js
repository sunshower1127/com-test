/**
 * 단계 2 디버그: 안 되는 것들 집중 테스트
 *
 * 사용법: cd electron && node test-step02-debug.js
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

console.log('HWP Version: ' + String(hwp.Version) + '\n');

var rl = readline.createInterface({ input: process.stdin, output: process.stdout });
function ask(msg) {
  return new Promise(function(resolve) {
    rl.question(msg, function() { resolve(); });
  });
}

(async function() {

  // 먼저 텍스트를 넣어둠
  hwp.HAction.GetDefault('InsertText', hwp.HParameterSet.HInsertText.HSet);
  hwp.HParameterSet.HInsertText.Text = '원본 텍스트입니다.';
  hwp.HAction.Execute('InsertText', hwp.HParameterSet.HInsertText.HSet);
  var before = bridge.comCallWith(h, 'GetTextFile', ['UNICODE', '']);
  console.log('원본: "' + String(before) + '"\n');

  // ──────────────────────────────────────
  // 디버그 1: SetTextFile 전체 교체
  // ──────────────────────────────────────
  await ask('── D1: SetTextFile("UNICODE", "") 전체 교체 시도 → Enter');
  try {
    var ret = bridge.comCallWith(h, 'SetTextFile', ['교체된 내용!', 'UNICODE', '']);
    console.log('  ret=' + String(ret));
    var after = bridge.comCallWith(h, 'GetTextFile', ['UNICODE', '']);
    console.log('  결과: "' + String(after) + '"');
    if (String(after).indexOf('원본') >= 0) {
      console.log('  ⚠️ 원본이 남아있음 → 교체가 아닌 삽입\n');
    } else {
      console.log('  ✅ 원본 사라짐 → 정상 교체\n');
    }
  } catch(e) { console.log('  ❌ ' + e.message + '\n'); }

  // ──────────────────────────────────────
  // 디버그 2: SetTextFile("TEXT", "") 로 시도
  // ──────────────────────────────────────
  await ask('── D2: SetTextFile("TEXT", "") 전체 교체 시도 → Enter');
  try {
    // 먼저 초기화
    hwp.Clear(1);
    bridge.comCallWith(h, 'Run', ['FileNew']);
    hwp.HAction.GetDefault('InsertText', hwp.HParameterSet.HInsertText.HSet);
    hwp.HParameterSet.HInsertText.Text = '원본2';
    hwp.HAction.Execute('InsertText', hwp.HParameterSet.HInsertText.HSet);
    console.log('  교체 전: "' + String(bridge.comCallWith(h, 'GetTextFile', ['UNICODE', ''])) + '"');

    var ret = bridge.comCallWith(h, 'SetTextFile', ['TEXT로 교체!', 'TEXT', '']);
    console.log('  ret=' + String(ret));
    var after = bridge.comCallWith(h, 'GetTextFile', ['UNICODE', '']);
    console.log('  결과: "' + String(after) + '"\n');
  } catch(e) { console.log('  ❌ ' + e.message + '\n'); }

  // ──────────────────────────────────────
  // 디버그 3: SelectAll + Delete 후 SetTextFile
  // ──────────────────────────────────────
  await ask('── D3: SelectAll → Delete → SetTextFile (우회) → Enter');
  try {
    // 초기화
    hwp.Clear(1);
    bridge.comCallWith(h, 'Run', ['FileNew']);
    hwp.HAction.GetDefault('InsertText', hwp.HParameterSet.HInsertText.HSet);
    hwp.HParameterSet.HInsertText.Text = '원본3';
    hwp.HAction.Execute('InsertText', hwp.HParameterSet.HInsertText.HSet);
    console.log('  교체 전: "' + String(bridge.comCallWith(h, 'GetTextFile', ['UNICODE', ''])) + '"');

    bridge.comCallWith(h, 'Run', ['SelectAll']);
    bridge.comCallWith(h, 'Run', ['Delete']);
    var ret = bridge.comCallWith(h, 'SetTextFile', ['우회로 교체 성공!', 'UNICODE', '']);
    console.log('  ret=' + String(ret));
    var after = bridge.comCallWith(h, 'GetTextFile', ['UNICODE', '']);
    console.log('  결과: "' + String(after) + '"');
    if (String(after).indexOf('원본') >= 0) {
      console.log('  ⚠️ 여전히 원본 남음\n');
    } else {
      console.log('  ✅ 깨끗하게 교체됨\n');
    }
  } catch(e) { console.log('  ❌ ' + e.message + '\n'); }

  // ──────────────────────────────────────
  // 디버그 4: saveblock 재테스트
  // ──────────────────────────────────────
  await ask('── D4: saveblock — SelectAll 후 GetTextFile → Enter');
  try {
    bridge.comCallWith(h, 'Run', ['SelectAll']);
    var text = bridge.comCallWith(h, 'GetTextFile', ['UNICODE', 'saveblock']);
    if (text === null || text === undefined || text === '' || text === false) {
      console.log('  ⚠️ 반환="' + String(text) + '"');
    } else {
      console.log('  ✅ "' + String(text).substring(0, 100) + '"');
    }
    console.log('');
  } catch(e) { console.log('  ❌ ' + e.message + '\n'); }

  // ──────────────────────────────────────
  // 디버그 5: saveblock — 마우스 선택 대신 키보드 선택 시뮬레이션
  // ──────────────────────────────────────
  await ask('── D5: MoveDocBegin → SelectAll → saveblock (다시) → Enter');
  try {
    bridge.comCallWith(h, 'Run', ['MoveDocBegin']);
    bridge.comCallWith(h, 'Run', ['SelectAll']);
    // SelectionMode 확인
    var mode = bridge.comGet(h, 'SelectionMode');
    console.log('  SelectionMode=' + String(mode));
    var text = bridge.comCallWith(h, 'GetTextFile', ['UNICODE', 'saveblock']);
    if (text === null || text === undefined || text === '' || text === false) {
      console.log('  ⚠️ 반환="' + String(text) + '"');
    } else {
      console.log('  ✅ "' + String(text).substring(0, 100) + '"');
    }
    console.log('');
  } catch(e) { console.log('  ❌ ' + e.message + '\n'); }

  // ──────────────────────────────────────
  // 디버그 6: GetTextFile("UTF8") 재확인 + "HWP" 포맷도 테스트
  // ──────────────────────────────────────
  await ask('── D6: GetTextFile 포맷별 재확인 (UTF8, HWP) → Enter');
  try {
    var utf8 = bridge.comCallWith(h, 'GetTextFile', ['UTF8', '']);
    console.log('  UTF8: type=' + typeof utf8 + ', val=' + String(utf8));

    var hwpFmt = bridge.comCallWith(h, 'GetTextFile', ['HWP', '']);
    if (hwpFmt === null || hwpFmt === undefined || hwpFmt === false) {
      console.log('  HWP: ⚠️ ' + String(hwpFmt));
    } else {
      console.log('  HWP: type=' + typeof hwpFmt + ', len=' + String(hwpFmt).length);
      console.log('  HWP 처음 80자: "' + String(hwpFmt).substring(0, 80) + '"');
    }
    console.log('');
  } catch(e) { console.log('  ❌ ' + e.message + '\n'); }

  // ── 종료 ──
  await ask('── 디버그 완료! Enter로 종료');
  bridge.comCallWith(h, 'Clear', [1]);
  bridge.comCallWith(h, 'Quit', []);
  rl.close();
  process.exit(0);

})();
