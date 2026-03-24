/**
 * 필드 디버그 2: 구분자 + 두번째 필드 접근
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

// 2개 필드 만들기
insert('Name: ');
bridge.comCallWith(h, 'CreateField', ['SET', '', 'f_name']);
insert(' Age: ');
bridge.comCallWith(h, 'CreateField', ['SET', '', 'f_age']);

// GetFieldList 구분자 분석
var list = String(bridge.comCallWith(h, 'GetFieldList', [0, 0]));
console.log('GetFieldList raw length: ' + list.length);
console.log('GetFieldList raw: "' + list + '"');

// 각 문자 코드 출력
var codes = [];
for (var i = 0; i < list.length; i++) {
  codes.push(list.charCodeAt(i));
}
console.log('charCodes: [' + codes.join(', ') + ']');

// 구분자 찾기
var sep = null;
for (var i = 0; i < list.length; i++) {
  var c = list.charCodeAt(i);
  if (c < 32) {
    sep = String.fromCharCode(c);
    console.log('구분자 발견: charCode=' + c + ' (\\x' + c.toString(16).padStart(2, '0') + ') at index ' + i);
    break;
  }
}

if (sep) {
  var fields = list.split(sep);
  console.log('필드 목록: ' + JSON.stringify(fields));
} else {
  console.log('구분자 없음 — 필드명이 연속');
}

// FieldExist 확인
console.log('\nFieldExist:');
console.log('  f_name: ' + bridge.comCallWith(h, 'FieldExist', ['f_name']));
console.log('  f_age: ' + bridge.comCallWith(h, 'FieldExist', ['f_age']));

// PutFieldText — 개별로
console.log('\nPutFieldText 개별:');
bridge.comCallWith(h, 'PutFieldText', ['f_name', 'John']);
bridge.comCallWith(h, 'PutFieldText', ['f_age', '25']);
console.log('  f_name: "' + bridge.comCallWith(h, 'GetFieldText', ['f_name']) + '"');
console.log('  f_age: "' + bridge.comCallWith(h, 'GetFieldText', ['f_age']) + '"');

// MoveToField
console.log('\nMoveToField:');
var ret1 = bridge.comCallWith(h, 'MoveToField', ['f_name', true, true, false]);
console.log('  f_name: ret=' + ret1);
var ret2 = bridge.comCallWith(h, 'MoveToField', ['f_age', true, true, false]);
console.log('  f_age: ret=' + ret2);

// MoveToField 다른 파라미터
var ret3 = bridge.comCallWith(h, 'MoveToField', ['f_age']);
console.log('  f_age (파라미터 1개): ret=' + ret3);

// GetFieldText 에서 구분자로 여러 필드 읽기
console.log('\nGetFieldText 구분자:');
if (sep) {
  var allText = bridge.comCallWith(h, 'GetFieldText', ['f_name' + sep + 'f_age']);
  console.log('  "f_name' + sep + 'f_age": "' + String(allText) + '"');
  // 결과도 구분자로 나뉘는지
  var vals = String(allText).split(sep);
  console.log('  분리: ' + JSON.stringify(vals));
}

// 전체 텍스트
var text = String(bridge.comCallWith(h, 'GetTextFile', ['UNICODE', '']));
console.log('\n전체 텍스트: "' + text + '"');

console.log('\n확인해주세요!');
