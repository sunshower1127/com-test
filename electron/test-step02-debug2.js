/**
 * 디버그 2차: SelectText vs Run("SelectAll") saveblock 차이
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
    rl.question(msg, function() { resolve(); });
  });
}

(async function() {

  // 텍스트 준비
  hwp.HAction.GetDefault('InsertText', hwp.HParameterSet.HInsertText.HSet);
  hwp.HParameterSet.HInsertText.Text = 'ABCDE12345';
  hwp.HAction.Execute('InsertText', hwp.HParameterSet.HInsertText.HSet);
  console.log('준비 완료: "ABCDE12345"\n');

  // ── T1: SelectText로 부분 선택 → saveblock ──
  await ask('── T1: SelectText(0,0,0,5) → saveblock → Enter');
  try {
    bridge.comCallWith(h, 'Run', ['MoveDocBegin']);
    var sel = bridge.comCallWith(h, 'SelectText', [0, 0, 0, 5]);
    var mode = bridge.comGet(h, 'SelectionMode');
    console.log('  SelectText ret=' + String(sel) + ', SelectionMode=' + String(mode));
    var text = bridge.comCallWith(h, 'GetTextFile', ['UNICODE', 'saveblock']);
    console.log('  saveblock: "' + String(text) + '" (기대: "ABCDE")\n');
  } catch(e) { console.log('  ❌ ' + e.message + '\n'); }

  // ── T2: 키보드 시뮬로 선택 (MoveSelDocEnd) ──
  await ask('── T2: MoveDocBegin → MoveSelRight x5 → saveblock → Enter');
  try {
    bridge.comCallWith(h, 'Run', ['MoveDocBegin']);
    for (var i = 0; i < 5; i++) {
      bridge.comCallWith(h, 'Run', ['MoveSelRight']);
    }
    var mode = bridge.comGet(h, 'SelectionMode');
    console.log('  SelectionMode=' + String(mode));
    var text = bridge.comCallWith(h, 'GetTextFile', ['UNICODE', 'saveblock']);
    console.log('  saveblock: "' + String(text) + '" (기대: "ABCDE")\n');
  } catch(e) { console.log('  ❌ ' + e.message + '\n'); }

  // ── T3: SelectAll → saveblock ──
  await ask('── T3: SelectAll → saveblock → Enter');
  try {
    bridge.comCallWith(h, 'Run', ['SelectAll']);
    var mode = bridge.comGet(h, 'SelectionMode');
    console.log('  SelectionMode=' + String(mode));
    var text = bridge.comCallWith(h, 'GetTextFile', ['UNICODE', 'saveblock']);
    console.log('  saveblock: "' + String(text) + '" (기대: "ABCDE12345")\n');
  } catch(e) { console.log('  ❌ ' + e.message + '\n'); }

  // ── T4: SelectText 범위를 좀 더 크게 ──
  await ask('── T4: SelectText(0,0,0,10) 전체 선택 → saveblock → Enter');
  try {
    bridge.comCallWith(h, 'Run', ['MoveDocBegin']);
    var sel = bridge.comCallWith(h, 'SelectText', [0, 0, 0, 10]);
    var mode = bridge.comGet(h, 'SelectionMode');
    console.log('  SelectText ret=' + String(sel) + ', SelectionMode=' + String(mode));
    var text = bridge.comCallWith(h, 'GetTextFile', ['UNICODE', 'saveblock']);
    console.log('  saveblock: "' + String(text) + '" (기대: "ABCDE12345")\n');
  } catch(e) { console.log('  ❌ ' + e.message + '\n'); }

  // ── 종료 ──
  await ask('── 완료! Enter로 종료');
  bridge.comCallWith(h, 'Clear', [1]);
  bridge.comCallWith(h, 'Quit', []);
  rl.close();
  process.exit(0);

})();
