/**
 * 자동교정 테스트
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

function insertAndCheck(input) {
  // 새 문서로 초기화
  bridge.comCallWith(h, 'Run', ['SelectAll']);
  bridge.comCallWith(h, 'Run', ['Delete']);
  // 삽입
  hwp.HAction.GetDefault('InsertText', hwp.HParameterSet.HInsertText.HSet);
  hwp.HParameterSet.HInsertText.Text = input;
  hwp.HAction.Execute('InsertText', hwp.HParameterSet.HInsertText.HSet);
  // 읽기
  var output = String(bridge.comCallWith(h, 'GetTextFile', ['UNICODE', '']));
  var changed = (input !== output);
  console.log('  입력: "' + input + '"');
  console.log('  결과: "' + output + '"');
  console.log('  ' + (changed ? '⚠️ 변경됨!' : '✅ 동일'));
  return output;
}

(async function() {

  console.log('HWP Version: ' + String(hwp.Version) + '\n');

  // ── T1: 띄어쓰기 교정 패턴들 ──
  await ask('── T1: 띄어쓰기 교정 패턴');
  insertAndCheck('첫번째');
  insertAndCheck('두번째');
  insertAndCheck('세번째');
  insertAndCheck('네번째');
  insertAndCheck('다섯번째');

  await ask('── T2: 더 많은 띄어쓰기 패턴');
  insertAndCheck('안녕하세요반갑습니다');
  insertAndCheck('대한민국서울시');
  insertAndCheck('오늘날씨가좋습니다');
  insertAndCheck('나는학생입니다');

  await ask('── T3: 영문 + 숫자는 교정 안 되는지');
  insertAndCheck('HelloWorld');
  insertAndCheck('ABCDE12345');
  insertAndCheck('test123');

  await ask('── T4: 한영 혼합');
  insertAndCheck('첫번째test');
  insertAndCheck('hello첫번째');
  insertAndCheck('ABC첫번째DEF');

  // ── T5: SetTextFile로 넣으면 교정되는지 ──
  await ask('── T5: SetTextFile로 삽입 시 교정 여부');
  try {
    bridge.comCallWith(h, 'Run', ['SelectAll']);
    bridge.comCallWith(h, 'Run', ['Delete']);
    bridge.comCallWith(h, 'SetTextFile', ['첫번째두번째세번째', 'UNICODE', '']);
    var text = String(bridge.comCallWith(h, 'GetTextFile', ['UNICODE', '']));
    console.log('  입력: "첫번째두번째세번째"');
    console.log('  결과: "' + text + '"');
    console.log('  ' + (text !== '첫번째두번째세번째' ? '⚠️ 변경됨!' : '✅ 동일'));
  } catch(e) { console.log('  ❌ ' + e.message); }

  // ── T6: 자동교정 끄기 시도 — SetAutoSpell ──
  await ask('── T6: 자동교정 끄기 시도');

  // 방법 1: Run("AutoSpellCheck") 토글
  try {
    bridge.comCallWith(h, 'Run', ['AutoSpellCheck']);
    console.log('  Run("AutoSpellCheck") 실행됨');
    insertAndCheck('첫번째');
  } catch(e) { console.log('  Run("AutoSpellCheck") ❌ ' + e.message); }

  // 방법 2: SetAutoSpell(false) / SpellCheck 프로퍼티
  try {
    bridge.comCallWith(h, 'SetAutoSpell', [0]);
    console.log('  SetAutoSpell(0) 실행됨');
    insertAndCheck('첫번째');
  } catch(e) { console.log('  SetAutoSpell ❌ ' + e.message); }

  // 방법 3: AutoCorrect 관련
  try {
    bridge.comCallWith(h, 'Run', ['ToggleAutoCorrect']);
    console.log('  Run("ToggleAutoCorrect") 실행됨');
    insertAndCheck('첫번째');
  } catch(e) { console.log('  ToggleAutoCorrect ❌ ' + e.message); }

  // 방법 4: 옵션으로 끄기
  try {
    bridge.comCallWith(h, 'Run', ['AutoCorrectOption']);
    console.log('  Run("AutoCorrectOption") 실행됨 (팝업 뜰 수 있음)');
  } catch(e) { console.log('  AutoCorrectOption ❌ ' + e.message); }

  // ── T7: 이미 자동교정이 꺼졌다면 재테스트 ──
  await ask('── T7: 자동교정 끄기 후 재테스트');
  insertAndCheck('첫번째');
  insertAndCheck('두번째');
  insertAndCheck('오늘날씨가좋습니다');

  // ── T8: BreakPara 후 삽입도 교정되는지 ──
  await ask('── T8: 문단 나눠서 각각 삽입');
  bridge.comCallWith(h, 'Run', ['SelectAll']);
  bridge.comCallWith(h, 'Run', ['Delete']);
  var words = ['첫번째', '두번째', '세번째', '네번째', '다섯번째'];
  for (var i = 0; i < words.length; i++) {
    if (i > 0) bridge.comCallWith(h, 'Run', ['BreakPara']);
    hwp.HAction.GetDefault('InsertText', hwp.HParameterSet.HInsertText.HSet);
    hwp.HParameterSet.HInsertText.Text = words[i];
    hwp.HAction.Execute('InsertText', hwp.HParameterSet.HInsertText.HSet);
  }
  var text = String(bridge.comCallWith(h, 'GetTextFile', ['UNICODE', '']));
  console.log('  입력: ' + words.join(' | '));
  console.log('  결과: "' + text + '"');
  var lines = text.split('\r\n');
  for (var j = 0; j < lines.length; j++) {
    var changed = (lines[j] !== words[j]);
    console.log('  [' + j + '] "' + words[j] + '" → "' + lines[j] + '" ' + (changed ? '⚠️' : '✅'));
  }

  // ── 종료 ──
  await ask('── 완료! 종료');
  bridge.comCallWith(h, 'Clear', [1]);
  bridge.comCallWith(h, 'Quit', []);
  rl.close();
  process.exit(0);

})();
