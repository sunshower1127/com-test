/**
 * 필드 디버그 3: Direction 파라미터 + 두번째 필드 접근 문제
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

// ── T1: Direction 다양한 값 ──
console.log('=== T1: Direction 값별 필드 생성 ===');
var directions = ['SET', '', 'NONE', 'set'];
directions.forEach(function(dir) {
  bridge.comCallWith(h, 'Run', ['SelectAll']);
  bridge.comCallWith(h, 'Run', ['Delete']);

  insert('A: ');
  var r1 = bridge.comCallWith(h, 'CreateField', [dir, '', 'fieldA']);
  insert(' B: ');
  var r2 = bridge.comCallWith(h, 'CreateField', [dir, '', 'fieldB']);

  var list = String(bridge.comCallWith(h, 'GetFieldList', [0, 0]));
  var fields = list.split('\x02');

  bridge.comCallWith(h, 'PutFieldText', ['fieldA', 'AAA']);
  bridge.comCallWith(h, 'PutFieldText', ['fieldB', 'BBB']);
  var vA = String(bridge.comCallWith(h, 'GetFieldText', ['fieldA']));
  var vB = String(bridge.comCallWith(h, 'GetFieldText', ['fieldB']));

  console.log('  dir="' + dir + '": r1=' + r1 + ', r2=' + r2 +
    ', fields=' + JSON.stringify(fields) +
    ', A="' + vA + '", B="' + vB + '"');
});

// ── T2: 같은 줄에 필드 2개 (공백 없이) ──
console.log('\n=== T2: 공백 없이 연속 필드 ===');
bridge.comCallWith(h, 'Run', ['SelectAll']);
bridge.comCallWith(h, 'Run', ['Delete']);
bridge.comCallWith(h, 'CreateField', ['SET', '', 'x1']);
bridge.comCallWith(h, 'CreateField', ['SET', '', 'x2']);

bridge.comCallWith(h, 'PutFieldText', ['x1', '111']);
bridge.comCallWith(h, 'PutFieldText', ['x2', '222']);
console.log('  x1="' + bridge.comCallWith(h, 'GetFieldText', ['x1']) + '"');
console.log('  x2="' + bridge.comCallWith(h, 'GetFieldText', ['x2']) + '"');

// ── T3: BreakPara로 문단 나눠서 ──
console.log('\n=== T3: BreakPara로 문단 나누기 ===');
bridge.comCallWith(h, 'Run', ['SelectAll']);
bridge.comCallWith(h, 'Run', ['Delete']);
bridge.comCallWith(h, 'CreateField', ['SET', '', 'p1']);
bridge.comCallWith(h, 'Run', ['BreakPara']);
bridge.comCallWith(h, 'CreateField', ['SET', '', 'p2']);
bridge.comCallWith(h, 'Run', ['BreakPara']);
bridge.comCallWith(h, 'CreateField', ['SET', '', 'p3']);

bridge.comCallWith(h, 'PutFieldText', ['p1', 'AAA']);
bridge.comCallWith(h, 'PutFieldText', ['p2', 'BBB']);
bridge.comCallWith(h, 'PutFieldText', ['p3', 'CCC']);
console.log('  p1="' + bridge.comCallWith(h, 'GetFieldText', ['p1']) + '"');
console.log('  p2="' + bridge.comCallWith(h, 'GetFieldText', ['p2']) + '"');
console.log('  p3="' + bridge.comCallWith(h, 'GetFieldText', ['p3']) + '"');

var text = String(bridge.comCallWith(h, 'GetTextFile', ['UNICODE', '']));
console.log('  전체: "' + text.replace(/\r?\n/g, '|') + '"');

// ── T4: \x02 구분자로 PutFieldText 한번에 ──
console.log('\n=== T4: PutFieldText 구분자로 한번에 ===');
try {
  bridge.comCallWith(h, 'PutFieldText', ['p1\x02p2\x02p3', 'XXX\x02YYY\x02ZZZ']);
  console.log('  p1="' + bridge.comCallWith(h, 'GetFieldText', ['p1']) + '"');
  console.log('  p2="' + bridge.comCallWith(h, 'GetFieldText', ['p2']) + '"');
  console.log('  p3="' + bridge.comCallWith(h, 'GetFieldText', ['p3']) + '"');
} catch(e) { console.log('  ❌ ' + e.message); }

// ── T5: MoveToField 후 GetFieldText ──
console.log('\n=== T5: MoveToField → 위치 확인 ===');
try {
  var r1 = bridge.comCallWith(h, 'MoveToField', ['p1', true, true, false]);
  var ps = bridge.comCallWith(h, 'GetPosBySet', []);
  console.log('  p1: ret=' + r1 + ', Para=' + bridge.comCallWith(ps, 'Item', ['Para']));
} catch(e) { console.log('  p1: ❌ ' + e.message); }

try {
  var r2 = bridge.comCallWith(h, 'MoveToField', ['p2', true, true, false]);
  var ps = bridge.comCallWith(h, 'GetPosBySet', []);
  console.log('  p2: ret=' + r2 + ', Para=' + bridge.comCallWith(ps, 'Item', ['Para']));
} catch(e) { console.log('  p2: ❌ ' + e.message); }

try {
  var r3 = bridge.comCallWith(h, 'MoveToField', ['p3', true, true, false]);
  var ps = bridge.comCallWith(h, 'GetPosBySet', []);
  console.log('  p3: ret=' + r3 + ', Para=' + bridge.comCallWith(ps, 'Item', ['Para']));
} catch(e) { console.log('  p3: ❌ ' + e.message); }

console.log('\n확인해주세요!');
