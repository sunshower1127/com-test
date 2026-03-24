/**
 * 필드 디버그: CreateField 동작 상세 확인
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
function run(label, fn) {
  process.stdout.write('── ' + label + ': ');
  try { fn(); } catch(e) { console.log('❌ ' + e.message); }
}

// ── CreateField 시그니처 확인 ──
console.log('=== CreateField 시그니처 ===');
var members = bridge.comListMembers(h);
var cf = members.filter(function(m) { return m.name === 'CreateField'; });
cf.forEach(function(m) { console.log(JSON.stringify(m)); });

// ── 간단한 테스트: 필드 하나만 ──
console.log('\n=== T1: 필드 하나만 ===');
var ret = bridge.comCallWith(h, 'CreateField', ['SET', '', 'field1']);
console.log('CreateField ret=' + ret);
var list = bridge.comCallWith(h, 'GetFieldList', [0, 0]);
console.log('GetFieldList: "' + String(list) + '"');
var text = String(bridge.comCallWith(h, 'GetTextFile', ['UNICODE', '']));
console.log('GetTextFile: "' + text + '"');

// ── T2: PutFieldText 후 확인 ──
bridge.comCallWith(h, 'PutFieldText', ['field1', 'VALUE1']);
var v = bridge.comCallWith(h, 'GetFieldText', ['field1']);
console.log('GetFieldText: "' + v + '"');

// ── T3: 텍스트 넣고 필드 만들기 ──
console.log('\n=== T3: 텍스트 + 필드 ===');
bridge.comCallWith(h, 'Run', ['SelectAll']);
bridge.comCallWith(h, 'Run', ['Delete']);

insert('Name: ');
bridge.comCallWith(h, 'CreateField', ['SET', '', 'f_name']);
insert(' / Age: ');
bridge.comCallWith(h, 'CreateField', ['SET', '', 'f_age']);

list = bridge.comCallWith(h, 'GetFieldList', [0, 0]);
console.log('필드 목록: "' + String(list) + '"');

bridge.comCallWith(h, 'PutFieldText', ['f_name', 'John']);
bridge.comCallWith(h, 'PutFieldText', ['f_age', '25']);
var n = bridge.comCallWith(h, 'GetFieldText', ['f_name']);
var a = bridge.comCallWith(h, 'GetFieldText', ['f_age']);
console.log('f_name="' + n + '", f_age="' + a + '"');
text = String(bridge.comCallWith(h, 'GetTextFile', ['UNICODE', '']));
console.log('전체: "' + text + '"');

// ── T4: BreakPara 후 필드 ──
console.log('\n=== T4: BreakPara 후 필드 ===');
bridge.comCallWith(h, 'Run', ['SelectAll']);
bridge.comCallWith(h, 'Run', ['Delete']);

insert('Line1: ');
bridge.comCallWith(h, 'CreateField', ['SET', '', 'line1']);
bridge.comCallWith(h, 'Run', ['BreakPara']);
insert('Line2: ');
bridge.comCallWith(h, 'CreateField', ['SET', '', 'line2']);

list = bridge.comCallWith(h, 'GetFieldList', [0, 0]);
console.log('필드 목록: "' + String(list) + '"');

bridge.comCallWith(h, 'PutFieldText', ['line1', 'AAA']);
bridge.comCallWith(h, 'PutFieldText', ['line2', 'BBB']);
var v1 = bridge.comCallWith(h, 'GetFieldText', ['line1']);
var v2 = bridge.comCallWith(h, 'GetFieldText', ['line2']);
console.log('line1="' + v1 + '", line2="' + v2 + '"');

// ── T5: \r\n 후 필드 ──
console.log('\n=== T5: \\r\\n 후 필드 ===');
bridge.comCallWith(h, 'Run', ['SelectAll']);
bridge.comCallWith(h, 'Run', ['Delete']);

insert('A: ');
bridge.comCallWith(h, 'CreateField', ['SET', '', 'fA']);
insert('\r\nB: ');
bridge.comCallWith(h, 'CreateField', ['SET', '', 'fB']);

list = bridge.comCallWith(h, 'GetFieldList', [0, 0]);
console.log('필드 목록: "' + String(list) + '"');

bridge.comCallWith(h, 'PutFieldText', ['fA', '111']);
bridge.comCallWith(h, 'PutFieldText', ['fB', '222']);
v1 = bridge.comCallWith(h, 'GetFieldText', ['fA']);
v2 = bridge.comCallWith(h, 'GetFieldText', ['fB']);
console.log('fA="' + v1 + '", fB="' + v2 + '"');
text = String(bridge.comCallWith(h, 'GetTextFile', ['UNICODE', '']));
console.log('전체: "' + text + '"');

// ── T6: 다른 필드 타입 ──
console.log('\n=== T6: 필드 타입 비교 ===');
bridge.comCallWith(h, 'Run', ['SelectAll']);
bridge.comCallWith(h, 'Run', ['Delete']);

var types = ['SET', 'GET', 'CLICK', 'REVISION', 'DATE', 'DOCINFO', 'USERINFO'];
types.forEach(function(t) {
  try {
    var ret = bridge.comCallWith(h, 'CreateField', [t, '', 'test_' + t]);
    insert(' ');
    console.log('  ' + t + ': ret=' + ret);
  } catch(e) {
    console.log('  ' + t + ': ❌ ' + e.message);
  }
});

list = bridge.comCallWith(h, 'GetFieldList', [0, 0]);
console.log('필드 목록: "' + String(list) + '"');

console.log('\n확인해주세요!');
