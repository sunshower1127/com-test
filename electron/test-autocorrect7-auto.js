/**
 * insert('\r\n')이 BreakPara와 동일한 결과를 만드는지 확인
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

// ── T1: insert('\r\n')으로 커서 위치에서 줄바꿈 ──
console.log('── T1: insert("\\r\\n")으로 줄바꿈');
clear();
insert('첫번째');
insert('\r\n');
insert('두번째');
var r = getText();
console.log('  결과: "' + r + '" ' + (r === '첫번째|두번째' ? '✅' : '⚠️ ' + r));

// ── T2: 중간에 끼워넣기 ──
console.log('\n── T2: 문서 중간에서 insert("\\r\\n")');
clear();
insert('AAABBB');
bridge.comCallWith(h, 'MovePos', [1, 0, 3]); // AAA|BBB 사이
insert('\r\n');
r = getText();
console.log('  결과: "' + r + '" (기대: "AAA|BBB")');

// ── T3: 빈 줄 여러개 ──
console.log('\n── T3: insert("\\r\\n\\r\\n\\r\\n") 빈 줄 3개');
clear();
insert('위');
insert('\r\n\r\n\r\n');
insert('아래');
r = getText();
console.log('  결과: "' + r + '" (기대: "위|||아래")');

// ── T4: 자동교정 확인 ──
console.log('\n── T4: 자동교정 안 되는지 최종 확인');
clear();
insert('첫번째');
insert('\r\n');
insert('두번째');
insert('\r\n');
insert('세번째');
r = getText();
console.log('  결과: "' + r + '"');
console.log('  ' + (r.indexOf(' ') === -1 ? '✅ 교정 안 됨' : '⚠️ 교정됨'));

// ── T5: PageCount 비교 (BreakPara vs \r\n이 같은 문단 구조인지) ──
console.log('\n── T5: 문단 구조 비교');
clear();
insert('A\r\nB\r\nC');
var rn_text = getText();
// GetPageText로 확인
var p0 = String(bridge.comCallWith(h, 'GetPageText', [0, ''])).replace(/\r?\n/g, '|');
console.log('  \\r\\n 결과: text="' + rn_text + '", page0="' + p0 + '"');

clear();
insert('A');
bridge.comCallWith(h, 'Run', ['BreakPara']);
insert('B');
bridge.comCallWith(h, 'Run', ['BreakPara']);
insert('C');
var bp_text = getText();
var p0b = String(bridge.comCallWith(h, 'GetPageText', [0, ''])).replace(/\r?\n/g, '|');
console.log('  BreakPara 결과: text="' + bp_text + '", page0="' + p0b + '"');

// ── T6: HWPML2X로 구조 비교 ──
console.log('\n── T6: HWPML2X 문단 구조 비교');
clear();
insert('가\r\n나');
var xml1 = String(bridge.comCallWith(h, 'GetTextFile', ['HWPML2X', '']));
var paraCount1 = (xml1.match(/<P /g) || []).length;
console.log('  \\r\\n: <P> 태그 수 = ' + paraCount1);

clear();
insert('가');
bridge.comCallWith(h, 'Run', ['BreakPara']);
insert('나');
var xml2 = String(bridge.comCallWith(h, 'GetTextFile', ['HWPML2X', '']));
var paraCount2 = (xml2.match(/<P /g) || []).length;
console.log('  BreakPara: <P> 태그 수 = ' + paraCount2);
console.log('  ' + (paraCount1 === paraCount2 ? '✅ 동일한 문단 구조' : '⚠️ 구조 다름 (' + paraCount1 + ' vs ' + paraCount2 + ')'));

// 종료
console.log('\n── 완료');
bridge.comCallWith(h, 'Clear', [1]);
bridge.comCallWith(h, 'Quit', []);
