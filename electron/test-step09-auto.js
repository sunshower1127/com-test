/**
 * 단계 9: 필드 — 자동 테스트
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

console.log('HWP Version: ' + String(hwp.Version) + '\n');

// 템플릿 문서 만들기
insert('계약서\r\n\r\n이름: ');

// ================================================================
// 9-1: CreateField — 필드 생성
// ================================================================
run('9-1 CreateField("name")', function() {
  var ret = bridge.comCallWith(h, 'CreateField', ['SET', '', 'name']);
  console.log('ret=' + ret);
});

insert('\r\n날짜: ');

run('9-2 CreateField("date")', function() {
  var ret = bridge.comCallWith(h, 'CreateField', ['SET', '', 'date']);
  console.log('ret=' + ret);
});

insert('\r\n금액: ');

run('9-3 CreateField("amount")', function() {
  var ret = bridge.comCallWith(h, 'CreateField', ['SET', '', 'amount']);
  console.log('ret=' + ret);
});

insert('\r\n비고: ');

run('9-4 CreateField("note")', function() {
  var ret = bridge.comCallWith(h, 'CreateField', ['SET', '', 'note']);
  console.log('ret=' + ret);
});

// ================================================================
// 9-5: FieldExist — 필드 존재 확인
// ================================================================
run('9-5 FieldExist', function() {
  var e1 = bridge.comCallWith(h, 'FieldExist', ['name']);
  var e2 = bridge.comCallWith(h, 'FieldExist', ['date']);
  var e3 = bridge.comCallWith(h, 'FieldExist', ['amount']);
  var e4 = bridge.comCallWith(h, 'FieldExist', ['nonexist']);
  console.log('name=' + e1 + ', date=' + e2 + ', amount=' + e3 + ', nonexist=' + e4);
});

// ================================================================
// 9-6: PutFieldText — 필드에 값 넣기
// ================================================================
run('9-6 PutFieldText', function() {
  bridge.comCallWith(h, 'PutFieldText', ['name', '홍길동']);
  bridge.comCallWith(h, 'PutFieldText', ['date', '2026-03-24']);
  bridge.comCallWith(h, 'PutFieldText', ['amount', '1,000,000원']);
  bridge.comCallWith(h, 'PutFieldText', ['note', '특이사항 없음']);
  console.log('✅ 4개 필드에 값 설정');
});

// ================================================================
// 9-7: GetFieldText — 필드 값 읽기
// ================================================================
run('9-7 GetFieldText', function() {
  var v1 = bridge.comCallWith(h, 'GetFieldText', ['name']);
  var v2 = bridge.comCallWith(h, 'GetFieldText', ['date']);
  var v3 = bridge.comCallWith(h, 'GetFieldText', ['amount']);
  var v4 = bridge.comCallWith(h, 'GetFieldText', ['note']);
  console.log('name="' + v1 + '", date="' + v2 + '", amount="' + v3 + '", note="' + v4 + '"');
});

// ================================================================
// 9-8: GetFieldList — 필드 목록
// ================================================================
run('9-8 GetFieldList', function() {
  var list = bridge.comCallWith(h, 'GetFieldList', [0, 0]);
  console.log('"' + String(list) + '"');
});

// ================================================================
// 9-9: MoveToField — 필드로 이동
// ================================================================
run('9-9 MoveToField("amount")', function() {
  var ret = bridge.comCallWith(h, 'MoveToField', ['amount', true, true, false]);
  console.log('ret=' + ret);
  // 현재 위치 확인
  var ps = bridge.comCallWith(h, 'GetPosBySet', []);
  var para = bridge.comCallWith(ps, 'Item', ['Para']);
  var pos = bridge.comCallWith(ps, 'Item', ['Pos']);
  console.log('  Para=' + para + ', Pos=' + pos);
});

// ================================================================
// 9-10: PutFieldText로 값 덮어쓰기
// ================================================================
run('9-10 PutFieldText 덮어쓰기', function() {
  bridge.comCallWith(h, 'PutFieldText', ['amount', '2,500,000원']);
  var v = bridge.comCallWith(h, 'GetFieldText', ['amount']);
  console.log('amount="' + v + '" (기대: "2,500,000원")');
});

// ================================================================
// 9-11: 여러 필드 한번에 설정 (구분자 사용)
// ================================================================
run('9-11 PutFieldText 여러 필드', function() {
  // GetFieldList로 필드 순서 확인 후, \x02 구분자로 한번에
  var list = String(bridge.comCallWith(h, 'GetFieldList', [0, 0]));
  console.log('필드 목록: "' + list + '"');

  // 한번에 여러 필드 설정 시도
  try {
    bridge.comCallWith(h, 'PutFieldText', ['name\x02date', '김철수\x022026-12-25']);
    var v1 = bridge.comCallWith(h, 'GetFieldText', ['name']);
    var v2 = bridge.comCallWith(h, 'GetFieldText', ['date']);
    console.log('  name="' + v1 + '", date="' + v2 + '"');
  } catch(e) { console.log('  ❌ ' + e.message); }
});

// ================================================================
// 9-12: GetTextFile로 전체 확인
// ================================================================
run('9-12 GetTextFile', function() {
  var text = String(bridge.comCallWith(h, 'GetTextFile', ['UNICODE', '']));
  console.log('\n' + text);
});

console.log('\n확인해주세요!');
