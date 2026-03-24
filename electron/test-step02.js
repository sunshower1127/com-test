/**
 * 단계 2: 텍스트 입출력 — 대화형 테스트 (한 단계씩)
 *
 * 사용법: cd electron && node test-step02.js
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

  // ── 2-1 ──
  await ask('── 2-1: InsertText (HParameterSet) → Enter');
  try {
    hwp.HAction.GetDefault('InsertText', hwp.HParameterSet.HInsertText.HSet);
    hwp.HParameterSet.HInsertText.Text = '안녕하세요 텍스트 테스트입니다.';
    var ret = hwp.HAction.Execute('InsertText', hwp.HParameterSet.HInsertText.HSet);
    console.log('  ✅ ret=' + String(ret) + '  → 화면 확인\n');
  } catch(e) { console.log('  ❌ ' + e.message + '\n'); }

  // ── 2-2 ──
  await ask('── 2-2: InsertText (CreateAction) → Enter');
  try {
    bridge.comCallWith(h, 'Run', ['BreakPara']);
    var act = hwp.CreateAction('InsertText');
    var set = act.CreateSet();
    act.GetDefault(set);
    set.SetItem('Text', 'CreateAction 패턴으로 삽입!');
    var ret = act.Execute(set);
    console.log('  ✅ ret=' + String(ret) + '  → 두 번째 줄 확인\n');
  } catch(e) { console.log('  ❌ ' + e.message + '\n'); }

  // ── 2-3 ──
  await ask('── 2-3: 여러 줄 삽입 (\\r\\n) → Enter');
  try {
    bridge.comCallWith(h, 'Run', ['BreakPara']);
    hwp.HAction.GetDefault('InsertText', hwp.HParameterSet.HInsertText.HSet);
    hwp.HParameterSet.HInsertText.Text = '첫째줄\r\n둘째줄\r\n셋째줄';
    var ret = hwp.HAction.Execute('InsertText', hwp.HParameterSet.HInsertText.HSet);
    console.log('  ✅ ret=' + String(ret) + '  → 3줄 확인\n');
  } catch(e) { console.log('  ❌ ' + e.message + '\n'); }

  // ── 2-4 ──
  await ask('── 2-4: GetTextFile("UNICODE", "") → Enter');
  try {
    var text = bridge.comCallWith(h, 'GetTextFile', ['UNICODE', '']);
    console.log('  ✅ len=' + String(text).length);
    console.log('  내용: "' + String(text).substring(0, 100) + '"\n');
  } catch(e) { console.log('  ❌ ' + e.message + '\n'); }

  // ── 2-5 ──
  await ask('── 2-5: GetTextFile("TEXT", "") → Enter');
  try {
    var text = bridge.comCallWith(h, 'GetTextFile', ['TEXT', '']);
    if (text === null || text === undefined || text === false) {
      console.log('  ⚠️ 반환=' + String(text) + '\n');
    } else {
      console.log('  len=' + String(text).length);
      console.log('  내용: "' + String(text).substring(0, 100) + '"\n');
    }
  } catch(e) { console.log('  ❌ ' + e.message + '\n'); }

  // ── 2-6 ──
  await ask('── 2-6: GetTextFile("UTF8", "") → Enter');
  try {
    var text = bridge.comCallWith(h, 'GetTextFile', ['UTF8', '']);
    if (text === null || text === undefined || text === false) {
      console.log('  ⚠️ 반환=' + String(text) + ' (미지원)\n');
    } else {
      console.log('  ✅ len=' + String(text).length);
      console.log('  내용: "' + String(text).substring(0, 100) + '"\n');
    }
  } catch(e) { console.log('  ❌ ' + e.message + '\n'); }

  // ── 2-7 ──
  await ask('── 2-7: GetTextFile("HTML", "") → Enter');
  try {
    var text = bridge.comCallWith(h, 'GetTextFile', ['HTML', '']);
    if (text === null || text === undefined || text === false) {
      console.log('  ⚠️ 반환=' + String(text) + '\n');
    } else {
      console.log('  ✅ len=' + String(text).length);
      console.log('  처음 200자: "' + String(text).substring(0, 200) + '"\n');
    }
  } catch(e) { console.log('  ❌ ' + e.message + '\n'); }

  // ── 2-8 ──
  await ask('── 2-8: GetTextFile("HWPML2X", "") → Enter');
  try {
    var text = bridge.comCallWith(h, 'GetTextFile', ['HWPML2X', '']);
    if (text === null || text === undefined || text === false) {
      console.log('  ⚠️ 반환=' + String(text) + '\n');
    } else {
      console.log('  ✅ len=' + String(text).length);
      console.log('  처음 200자: "' + String(text).substring(0, 200) + '"\n');
    }
  } catch(e) { console.log('  ❌ ' + e.message + '\n'); }

  // ── 2-9 ──
  await ask('── 2-9: SetTextFile 전체 교체 ("UNICODE", "") → Enter');
  try {
    var ret = bridge.comCallWith(h, 'SetTextFile', ['SetTextFile로 교체된 내용입니다.', 'UNICODE', '']);
    console.log('  ret=' + String(ret));
    var text = bridge.comCallWith(h, 'GetTextFile', ['UNICODE', '']);
    console.log('  교체 후: "' + String(text).substring(0, 100) + '"');
    console.log('  → 기존 텍스트가 사라지고 새 텍스트만 보이는지 확인\n');
  } catch(e) { console.log('  ❌ ' + e.message + '\n'); }

  // ── 2-10 ──
  await ask('── 2-10: SetTextFile insertfile → Enter');
  try {
    var ret = bridge.comCallWith(h, 'SetTextFile', ['\r\n추가 삽입된 텍스트!', 'UNICODE', 'insertfile']);
    console.log('  ret=' + String(ret));
    var text = bridge.comCallWith(h, 'GetTextFile', ['UNICODE', '']);
    console.log('  삽입 후: "' + String(text).substring(0, 150) + '"');
    console.log('  → 기존 텍스트 뒤에 추가되었는지 확인\n');
  } catch(e) { console.log('  ❌ ' + e.message + '\n'); }

  // ── 2-11 ──
  await ask('── 2-11: GetPageText(0, "") → Enter');
  try {
    var text = bridge.comCallWith(h, 'GetPageText', [0, '']);
    if (text === null || text === undefined || text === false) {
      console.log('  ⚠️ pgno=0 반환=' + String(text));
      try {
        var text2 = bridge.comCallWith(h, 'GetPageText', [1, '']);
        console.log('  pgno=1: "' + String(text2).substring(0, 80) + '"\n');
      } catch(e2) { console.log('  pgno=1도 실패: ' + e2.message + '\n'); }
    } else {
      console.log('  ✅ "' + String(text).substring(0, 100) + '"\n');
    }
  } catch(e) { console.log('  ❌ ' + e.message + '\n'); }

  // ── 2-12 ──
  await ask('── 2-12: GetHeadingString() → Enter');
  try {
    var heading = bridge.comCallWith(h, 'GetHeadingString', []);
    console.log('  ✅ "' + String(heading) + '" (빈 문자열 = 제목 없음)\n');
  } catch(e) { console.log('  ❌ ' + e.message + '\n'); }

  // ── 2-13 ──
  await ask('── 2-13: BreakPara x2 + 텍스트 → Enter');
  try {
    bridge.comCallWith(h, 'Run', ['BreakPara']);
    bridge.comCallWith(h, 'Run', ['BreakPara']);
    hwp.HAction.GetDefault('InsertText', hwp.HParameterSet.HInsertText.HSet);
    hwp.HParameterSet.HInsertText.Text = '← 위에 빈 줄 2개';
    hwp.HAction.Execute('InsertText', hwp.HParameterSet.HInsertText.HSet);
    var text = bridge.comCallWith(h, 'GetTextFile', ['UNICODE', '']);
    console.log('  ✅ 전체: "' + String(text).substring(0, 150) + '"\n');
  } catch(e) { console.log('  ❌ ' + e.message + '\n'); }

  // ── 2-14 ──
  await ask('── 2-14: SelectText + GetTextFile saveblock → Enter');
  try {
    bridge.comCallWith(h, 'Run', ['MoveDocBegin']);
    var sel = bridge.comCallWith(h, 'SelectText', [0, 0, 0, 5]);
    console.log('  SelectText ret=' + String(sel));
    var text = bridge.comCallWith(h, 'GetTextFile', ['UNICODE', 'saveblock']);
    if (text === null || text === undefined || text === '' || text === false) {
      console.log('  ⚠️ saveblock 반환="' + String(text) + '"\n');
    } else {
      console.log('  ✅ 선택 텍스트: "' + String(text) + '"\n');
    }
  } catch(e) { console.log('  ❌ ' + e.message + '\n'); }

  // ── 종료 ──
  await ask('── 테스트 완료! Enter로 종료');
  bridge.comCallWith(h, 'Clear', [1]);
  bridge.comCallWith(h, 'Quit', []);
  rl.close();
  process.exit(0);

})();
