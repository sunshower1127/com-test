/**
 * 단계 7 디버그: 표 진입 + 셀 서식 + 표 서식
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

// 표 생성 + 셀 채우기
hwp.HAction.GetDefault('TableCreate', hwp.HParameterSet.HTableCreation.HSet);
hwp.HParameterSet.HTableCreation.Rows = 3;
hwp.HParameterSet.HTableCreation.Cols = 2;
hwp.HParameterSet.HTableCreation.WidthType = 2;
hwp.HAction.Execute('TableCreate', hwp.HParameterSet.HTableCreation.HSet);

var cells = ['A1', 'B1', 'A2', 'B2', 'A3', 'B3'];
for (var i = 0; i < cells.length; i++) {
  insert(cells[i]);
  if (i < cells.length - 1) bridge.comCallWith(h, 'Run', ['TableRightCell']);
}
bridge.comCallWith(h, 'Run', ['MoveDocEnd']);
console.log('3x2 표 생성 완료\n');

// ================================================================
// D1: 표 안으로 다시 들어가기
// ================================================================
run('D1 표 진입 방법 탐색', function() {
  // MoveDocBegin → MoveRight로 진입?
  bridge.comCallWith(h, 'Run', ['MoveDocBegin']);
  bridge.comCallWith(h, 'Run', ['MoveRight']); // 표 컨트롤로 진입?

  var text = String(bridge.comCallWith(h, 'GetTextFile', ['UNICODE', 'saveblock']));
  bridge.comCallWith(h, 'Run', ['Cancel']);

  // GetPosBySet으로 위치 확인
  var ps = bridge.comCallWith(h, 'GetPosBySet', []);
  var list = bridge.comCallWith(ps, 'Item', ['List']);
  var para = bridge.comCallWith(ps, 'Item', ['Para']);
  var pos = bridge.comCallWith(ps, 'Item', ['Pos']);
  console.log('List=' + list + ', Para=' + para + ', Pos=' + pos);

  // 텍스트 확인
  bridge.comCallWith(h, 'Run', ['MoveSelLineEnd']);
  text = String(bridge.comCallWith(h, 'GetTextFile', ['UNICODE', 'saveblock']));
  bridge.comCallWith(h, 'Run', ['Cancel']);
  console.log('  현재 줄: "' + text + '"');
});

// ================================================================
// D2: MoveNextPos로 진입
// ================================================================
run('D2 MoveDocBegin → MoveNextPos', function() {
  bridge.comCallWith(h, 'Run', ['MoveDocBegin']);
  bridge.comCallWith(h, 'Run', ['MoveNextPosEx']); // 확장 이동

  var ps = bridge.comCallWith(h, 'GetPosBySet', []);
  var list = bridge.comCallWith(ps, 'Item', ['List']);
  var para = bridge.comCallWith(ps, 'Item', ['Para']);
  var pos = bridge.comCallWith(ps, 'Item', ['Pos']);
  console.log('List=' + list + ', Para=' + para + ', Pos=' + pos);

  bridge.comCallWith(h, 'Run', ['MoveSelLineEnd']);
  var text = String(bridge.comCallWith(h, 'GetTextFile', ['UNICODE', 'saveblock']));
  bridge.comCallWith(h, 'Run', ['Cancel']);
  console.log('  현재 줄: "' + text + '"');
});

// ================================================================
// D3: MovePos로 표 안으로 — list 파라미터
// ================================================================
run('D3 MovePos 다양한 시도', function() {
  // 표 안 셀은 다른 "list"에 속할 수 있음
  bridge.comCallWith(h, 'Run', ['MoveDocBegin']);

  // ShapeObject 접근
  try {
    bridge.comCallWith(h, 'Run', ['MoveRight']); // 표 진입 시도
    bridge.comCallWith(h, 'Run', ['MoveRight']);

    var ps = bridge.comCallWith(h, 'GetPosBySet', []);
    var list = bridge.comCallWith(ps, 'Item', ['List']);
    var para = bridge.comCallWith(ps, 'Item', ['Para']);
    var pos = bridge.comCallWith(ps, 'Item', ['Pos']);
    console.log('2번 MoveRight 후: List=' + list + ', Para=' + para + ', Pos=' + pos);

    bridge.comCallWith(h, 'Run', ['MoveSelLineEnd']);
    var text = String(bridge.comCallWith(h, 'GetTextFile', ['UNICODE', 'saveblock']));
    bridge.comCallWith(h, 'Run', ['Cancel']);
    console.log('  현재 줄: "' + text + '"');
  } catch(e) { console.log('❌ ' + e.message); }
});

// ================================================================
// D4: 표 안에서 CharShape 적용
// ================================================================
run('D4 표 안에서 CharShape', function() {
  // 이미 표 안에 있다면
  bridge.comCallWith(h, 'Run', ['MoveSelLineEnd']);
  var text = String(bridge.comCallWith(h, 'GetTextFile', ['UNICODE', 'saveblock']));

  if (text && text !== 'null' && text.length > 0) {
    hwp.HAction.GetDefault('CharShape', hwp.HParameterSet.HCharShape.HSet);
    hwp.HParameterSet.HCharShape.Bold = 1;
    hwp.HParameterSet.HCharShape.TextColor = bridge.comCallWith(h, 'RGBColor', [255, 0, 0]);
    hwp.HAction.Execute('CharShape', hwp.HParameterSet.HCharShape.HSet);
    bridge.comCallWith(h, 'Run', ['Cancel']);
    console.log('✅ "' + text + '"에 볼드+빨간색 적용');
  } else {
    bridge.comCallWith(h, 'Run', ['Cancel']);
    console.log('⚠️ 표 안에 없음');
  }
});

// ================================================================
// D5: 셀 이동하면서 서식 적용
// ================================================================
run('D5 여러 셀에 서식', function() {
  // B1으로 이동
  bridge.comCallWith(h, 'Run', ['TableRightCell']);
  bridge.comCallWith(h, 'Run', ['MoveSelLineEnd']);
  hwp.HAction.GetDefault('CharShape', hwp.HParameterSet.HCharShape.HSet);
  hwp.HParameterSet.HCharShape.Height = 2000;
  hwp.HParameterSet.HCharShape.Italic = 1;
  hwp.HAction.Execute('CharShape', hwp.HParameterSet.HCharShape.HSet);
  bridge.comCallWith(h, 'Run', ['Cancel']);
  console.log('✅ B1에 20pt 이탤릭');
});

// ================================================================
// D6: CellShape / 표 관련 프로퍼티 탐색
// ================================================================
run('D6 CellShape 프로퍼티', function() {
  try {
    var cs = bridge.comGet(h, 'CellShape');
    console.log('type=' + typeof cs);
    if (cs && typeof cs === 'object') {
      var members = bridge.comListMembers(cs);
      var names = members.filter(function(m) {
        return ['QueryInterface','AddRef','Release','GetTypeInfoCount','GetTypeInfo','GetIDsOfNames','Invoke'].indexOf(m.name) === -1;
      }).map(function(m) { return m.name + '(' + m.kind + ')'; });
      console.log('  멤버: ' + names.join(', '));
    }
  } catch(e) { console.log('❌ ' + e.message); }
});

// ================================================================
// D7: HTableCreation 멤버 전체 확인
// ================================================================
run('D7 HTableCreation 멤버', function() {
  var hps = bridge.comGet(h, 'HParameterSet');
  var htc = bridge.comGet(hps, 'HTableCreation');
  var members = bridge.comListMembers(htc);
  var props = members.filter(function(m) {
    return m.kind === 'put' || m.kind === 'get/put';
  });
  console.log('put 프로퍼티:');
  props.forEach(function(m) {
    console.log('  ' + m.name);
  });
});

// ================================================================
// D8: 표 셀 배경색 — CellShape에서
// ================================================================
run('D8 CellShape.Item 시도', function() {
  try {
    var cs = bridge.comGet(h, 'CellShape');
    // ParameterSet이면 Item으로 접근
    var items = ['Width', 'Height', 'BackColor', 'BorderFill',
                 'CellMarginLeft', 'CellMarginRight', 'CellMarginTop', 'CellMarginBottom'];
    items.forEach(function(name) {
      try {
        var v = bridge.comCallWith(cs, 'Item', [name]);
        if (v !== null && v !== undefined) {
          console.log('  ' + name + '=' + v);
        }
      } catch(e) {}
    });
    // ItemExist로 확인
    var exist = ['Width', 'Height', 'BackColor', 'BorderFill', 'CellWidth', 'CellHeight',
                 'LeftMargin', 'RightMargin', 'TopMargin', 'BottomMargin', 'Protect',
                 'Editable', 'Dirty', 'Header', 'VertAlign'];
    console.log('  ItemExist:');
    exist.forEach(function(name) {
      try {
        var e = bridge.comCallWith(cs, 'ItemExist', [name]);
        if (e) console.log('    ' + name + ': ✅');
      } catch(err) {}
    });
  } catch(e) { console.log('❌ ' + e.message); }
});

// 닫지 않음
console.log('\n확인해주세요!');
