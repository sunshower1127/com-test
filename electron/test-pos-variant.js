/**
 * GetPosBySet().Item("Pos") 반환 타입 조사
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

insert('ABCDEFGHIJ');
bridge.comCallWith(h, 'Run', ['MoveDocBegin']);

// GetPosBySet 호출
var posSet = bridge.comCallWith(h, 'GetPosBySet', []);
console.log('posSet type: ' + typeof posSet);

// listMembers로 Set의 멤버 확인
var members = bridge.comListMembers(posSet);
console.log('\nSet 멤버:');
members.forEach(function(m) {
  if (['QueryInterface','AddRef','Release','GetTypeInfoCount','GetTypeInfo','GetIDsOfNames','Invoke'].indexOf(m.name) === -1) {
    console.log('  ' + JSON.stringify(m));
  }
});

// Item 호출 — 다양한 방법 시도
console.log('\n=== Item 호출 시도 ===');

// 방법 1: comCallWith로 Item
try {
  var val = bridge.comCallWith(posSet, 'Item', ['Pos']);
  console.log('comCallWith Item("Pos"): ' + String(val) + ' (type: ' + typeof val + ')');
} catch(e) { console.log('comCallWith Item("Pos"): ❌ ' + e.message); }

try {
  var val = bridge.comCallWith(posSet, 'Item', ['Para']);
  console.log('comCallWith Item("Para"): ' + String(val) + ' (type: ' + typeof val + ')');
} catch(e) { console.log('comCallWith Item("Para"): ❌ ' + e.message); }

try {
  var val = bridge.comCallWith(posSet, 'Item', ['List']);
  console.log('comCallWith Item("List"): ' + String(val) + ' (type: ' + typeof val + ')');
} catch(e) { console.log('comCallWith Item("List"): ❌ ' + e.message); }

// 방법 2: comGet으로 직접
try {
  var val = bridge.comGet(posSet, 'Pos');
  console.log('comGet "Pos": ' + String(val) + ' (type: ' + typeof val + ')');
} catch(e) { console.log('comGet "Pos": ❌ ' + e.message); }

try {
  var val = bridge.comGet(posSet, 'Para');
  console.log('comGet "Para": ' + String(val) + ' (type: ' + typeof val + ')');
} catch(e) { console.log('comGet "Para": ❌ ' + e.message); }

// 방법 3: ItemExist 확인
try {
  var exists = bridge.comCallWith(posSet, 'ItemExist', ['Pos']);
  console.log('ItemExist("Pos"): ' + String(exists));
  exists = bridge.comCallWith(posSet, 'ItemExist', ['Para']);
  console.log('ItemExist("Para"): ' + String(exists));
  exists = bridge.comCallWith(posSet, 'ItemExist', ['List']);
  console.log('ItemExist("List"): ' + String(exists));
} catch(e) { console.log('ItemExist: ❌ ' + e.message); }

// 방법 4: SetID, Count 확인
try {
  var sid = bridge.comGet(posSet, 'SetID');
  console.log('SetID: ' + String(sid));
  var count = bridge.comGet(posSet, 'Count');
  console.log('Count: ' + String(count));
} catch(e) { console.log('SetID/Count: ❌ ' + e.message); }

// 종료
console.log('\n완료');
bridge.comCallWith(h, 'Clear', [1]);
bridge.comCallWith(h, 'Quit', []);
