/**
 * 동적 오프셋 감지: MoveParaBegin 후 위치로 오프셋 계산
 *
 * 문제: 브릿지에서 GetPosBySet().Item("Pos")가 null이므로
 * 다른 방법으로 오프셋을 알아내야 함
 */

var bridge = require('./native/com_bridge_node.node');
var { initBridge, createComProxy } = require('./dist/worker/proxy');

bridge.comInit();
initBridge(bridge);

var h = bridge.comCreate('HWPFrame.HwpObject');
var hwp = createComProxy(h);
var win0 = bridge.comCallWith(bridge.comGet(h, 'XHwpWindows'), 'Item', [0]);
bridge.comPut(win0, 'Visible', true);

function insert(text) {
  hwp.HAction.GetDefault('InsertText', hwp.HParameterSet.HInsertText.HSet);
  hwp.HParameterSet.HInsertText.Text = text;
  hwp.HAction.Execute('InsertText', hwp.HParameterSet.HInsertText.HSet);
}
function peekAfter(n) {
  for (var i = 0; i < n; i++) bridge.comCallWith(h, 'Run', ['MoveSelRight']);
  var t = bridge.comCallWith(h, 'GetTextFile', ['UNICODE', 'saveblock']);
  bridge.comCallWith(h, 'Run', ['Cancel']);
  return String(t);
}

console.log('HWP Version: ' + String(hwp.Version) + '\n');

// ================================================================
// 방법 1: 이진 탐색으로 오프셋 찾기
// SetPos(0, para, pos) → MoveSelRight → 첫 글자가 나오면 그게 오프셋
// ================================================================
function findOffset(para) {
  // 해당 문단에 텍스트가 있는지 먼저 확인
  bridge.comCallWith(h, 'SetPos', [0, para, 0]);

  for (var pos = 0; pos <= 32; pos++) {
    bridge.comCallWith(h, 'SetPos', [0, para, pos]);
    // MoveSelRight로 1글자 선택
    bridge.comCallWith(h, 'Run', ['MoveSelRight']);
    var t = bridge.comCallWith(h, 'GetTextFile', ['UNICODE', 'saveblock']);
    bridge.comCallWith(h, 'Run', ['Cancel']);
    var ch = String(t);
    // 실제 텍스트 글자인지 확인 (null, 빈문자열, 줄바꿈이 아닌)
    if (ch && ch !== 'null' && ch !== '' && ch !== '\r\n' && ch !== '\n' && ch.charCodeAt(0) > 31) {
      return { offset: pos, firstChar: ch };
    }
  }
  return { offset: -1, firstChar: null };
}

// ================================================================
// T1: 기본 문서 (ABC / DEF / GHI)
// ================================================================
console.log('=== T1: 기본 3문단 ===');
insert('ABC');
bridge.comCallWith(h, 'Run', ['BreakPara']);
insert('DEF');
bridge.comCallWith(h, 'Run', ['BreakPara']);
insert('GHI');

for (var para = 0; para < 3; para++) {
  var result = findOffset(para);
  console.log('  para=' + para + ' → offset=' + result.offset + ', firstChar="' + result.firstChar + '"');
}

// ================================================================
// T2: 머리말 추가 후 오프셋 변화
// ================================================================
console.log('\n=== T2: 머리말(Header) 추가 후 ===');
// 머리말 삽입 시도
try {
  hwp.HAction.GetDefault('InsertHeader', hwp.HParameterSet.HHeaderFooter.HSet);
  hwp.HParameterSet.HHeaderFooter.HeaderType = 0; // 양쪽
  hwp.HAction.Execute('InsertHeader', hwp.HParameterSet.HHeaderFooter.HSet);
  // 머리말에 텍스트 삽입
  insert('HEADER TEXT');
  // 본문으로 돌아가기
  bridge.comCallWith(h, 'Run', ['CloseEx']);
  console.log('  머리말 추가 성공');
} catch(e) {
  console.log('  머리말 추가 실패: ' + e.message);
  // Run으로 시도
  try {
    bridge.comCallWith(h, 'Run', ['MoveDocBegin']);
    bridge.comCallWith(h, 'Run', ['InsertHeader']);
    insert('HEADER');
    bridge.comCallWith(h, 'Run', ['CloseEx']);
    console.log('  머리말 추가 성공 (Run)');
  } catch(e2) {
    console.log('  머리말 추가 실패 (Run): ' + e2.message);
  }
}

for (var para = 0; para < 3; para++) {
  var result = findOffset(para);
  console.log('  para=' + para + ' → offset=' + result.offset + ', firstChar="' + result.firstChar + '"');
}

// ================================================================
// T3: 꼬리말도 추가
// ================================================================
console.log('\n=== T3: 꼬리말(Footer) 추가 후 ===');
try {
  bridge.comCallWith(h, 'Run', ['MoveDocBegin']);
  hwp.HAction.GetDefault('InsertFooter', hwp.HParameterSet.HHeaderFooter.HSet);
  hwp.HParameterSet.HHeaderFooter.FooterType = 0;
  hwp.HAction.Execute('InsertFooter', hwp.HParameterSet.HHeaderFooter.HSet);
  insert('FOOTER TEXT');
  bridge.comCallWith(h, 'Run', ['CloseEx']);
  console.log('  꼬리말 추가 성공');
} catch(e) {
  console.log('  꼬리말 추가 실패: ' + e.message);
}

for (var para = 0; para < 3; para++) {
  var result = findOffset(para);
  console.log('  para=' + para + ' → offset=' + result.offset + ', firstChar="' + result.firstChar + '"');
}

// ================================================================
// T4: 새 문서에서 오프셋 확인
// ================================================================
console.log('\n=== T4: 새 문서 ===');
bridge.comCallWith(h, 'Clear', [1]);
bridge.comCallWith(h, 'Run', ['FileNew']);
insert('NEWDOC');
var result = findOffset(0);
console.log('  para=0 → offset=' + result.offset + ', firstChar="' + result.firstChar + '"');

// ================================================================
// T5: findOffset을 활용한 안전한 SetPos 래퍼
// ================================================================
console.log('\n=== T5: 안전한 SetPos 래퍼 테스트 ===');
bridge.comCallWith(h, 'Run', ['SelectAll']);
bridge.comCallWith(h, 'Run', ['Delete']);
insert('ABCDEFGHIJ');
bridge.comCallWith(h, 'Run', ['BreakPara']);
insert('KLMNOPQRST');

function safeSetPos(list, para, charIndex) {
  var info = findOffset(para);
  if (info.offset < 0) return false;
  return bridge.comCallWith(h, 'SetPos', [list, para, info.offset + charIndex]);
}

safeSetPos(0, 0, 3);
console.log('  para=0 char=3 → "' + peekAfter(1) + '" (기대: D)');
safeSetPos(0, 1, 3);
console.log('  para=1 char=3 → "' + peekAfter(1) + '" (기대: N)');
safeSetPos(0, 0, 9);
console.log('  para=0 char=9 → "' + peekAfter(1) + '" (기대: J)');

// 종료
console.log('\n완료');
bridge.comCallWith(h, 'Clear', [1]);
bridge.comCallWith(h, 'Quit', []);
