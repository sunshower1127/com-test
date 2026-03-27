/**
 * HTML vs HWPML2X 라운드트립 비교
 *
 * 서식 있는 문서를 꺼냈다가 다시 넣으면 서식이 유지되는가?
 */

var bridge = require('./native/com_bridge_node.node');
var { initBridge, createComProxy } = require('./dist/worker/proxy');
var fs = require('fs');
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

(async function() {

  // 이력서 열기
  var filePath = 'C:/Users/seng1/Documents/Projects/com-test/data/이력서_자기소개서_기본.hwp';
  bridge.comCallWith(h, 'Open', [filePath, 'HWP', '']);

  await ask('── 원본 이력서가 열려있습니다. 현재 상태를 확인하세요');

  // 두 포맷 추출
  var hwpml = String(bridge.comCallWith(h, 'GetTextFile', ['HWPML2X', '']));
  var html = String(bridge.comCallWith(h, 'GetTextFile', ['HTML', '']));
  console.log('  HWPML2X: ' + hwpml.length + '자');
  console.log('  HTML: ' + html.length + '자');

  // ── 테스트 A: HWPML2X 라운드트립 ──
  await ask('── HWPML2X 라운드트립 (꺼냈다가 그대로 넣기)');
  bridge.comCallWith(h, 'Run', ['SelectAll']);
  bridge.comCallWith(h, 'Run', ['Delete']);
  bridge.comCallWith(h, 'SetTextFile', [hwpml, 'HWPML2X', '']);
  console.log('  HWPML2X 넣기 완료 — 서식이 원본과 동일한지 확인하세요');

  // ── 테스트 B: HTML 라운드트립 ──
  await ask('── HTML 라운드트립 (새 문서에 HTML 넣기)');
  bridge.comCallWith(h, 'Clear', [1]);
  bridge.comCallWith(h, 'Run', ['FileNew']);
  bridge.comCallWith(h, 'SetTextFile', [html, 'HTML', '']);
  console.log('  HTML 넣기 완료 — 서식이 원본과 비교해서 어떤지 확인하세요');

  // ── 테스트 C: HTML에서 텍스트 수정 후 넣기 ──
  await ask('── HTML 수정 후 넣기');
  var modifiedHtml = html.replace(/지원분야/g, '지원분야: 개발팀');
  bridge.comCallWith(h, 'Clear', [1]);
  bridge.comCallWith(h, 'Run', ['FileNew']);
  bridge.comCallWith(h, 'SetTextFile', [modifiedHtml, 'HTML', '']);
  console.log('  HTML 수정본 넣기 완료');

  await ask('── 완료! 종료');
  bridge.comCallWith(h, 'Clear', [1]);
  bridge.comCallWith(h, 'Quit', []);
  rl.close();
  process.exit(0);

})();
