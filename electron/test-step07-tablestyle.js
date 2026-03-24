/**
 * 표 자체 서식: 셀 배경색, 테두리, 셀 크기 등
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

// 3x3 표 생성
hwp.HAction.GetDefault('TableCreate', hwp.HParameterSet.HTableCreation.HSet);
hwp.HParameterSet.HTableCreation.Rows = 3;
hwp.HParameterSet.HTableCreation.Cols = 3;
hwp.HParameterSet.HTableCreation.WidthType = 2;
hwp.HAction.Execute('TableCreate', hwp.HParameterSet.HTableCreation.HSet);

// 셀 채우기
var cells = ['A1','B1','C1','A2','B2','C2','A3','B3','C3'];
for (var i = 0; i < 9; i++) {
  insert(cells[i]);
  if (i < 8) bridge.comCallWith(h, 'Run', ['TableRightCell']);
}
// 첫 셀로 돌아가기
bridge.comCallWith(h, 'Run', ['TableUpperCell']);
bridge.comCallWith(h, 'Run', ['TableUpperCell']);
bridge.comCallWith(h, 'Run', ['TableLeftCell']);
bridge.comCallWith(h, 'Run', ['TableLeftCell']);
console.log('3x3 표 생성 완료\n');

// ================================================================
// T1: CellBorderFill — 셀 배경색/테두리
// ================================================================
run('T1 CellBorderFill 액션 탐색', function() {
  // CellBorderFill 관련 액션 찾기
  var actions = ['CellBorderFill', 'TableCellBorderFill', 'CellFill',
                 'TableCellFill', 'TableProperty', 'TablePropertyDialog',
                 'CellProperty', 'BorderFill'];
  actions.forEach(function(name) {
    try {
      var act = bridge.comCallWith(h, 'CreateAction', [name]);
      if (act) {
        var set = bridge.comCallWith(act, 'CreateSet', []);
        var sid = set ? bridge.comGet(set, 'SetID') : 'null';
        console.log('\n  ' + name + ': ✅ SetID=' + sid);
      } else {
        console.log('\n  ' + name + ': null');
      }
    } catch(e) {}
  });
});

// ================================================================
// T2: HParameterSet에서 표 관련 찾기
// ================================================================
run('\nT2 HParameterSet 표 관련', function() {
  var hps = bridge.comGet(h, 'HParameterSet');
  var members = bridge.comListMembers(hps);
  var tableRelated = members.filter(function(m) {
    var n = m.name.toLowerCase();
    return n.indexOf('table') >= 0 || n.indexOf('cell') >= 0 || n.indexOf('border') >= 0;
  });
  tableRelated.forEach(function(m) {
    console.log('\n  ' + m.name + ' (' + m.kind + ')');
  });
});

// ================================================================
// T3: CellBorderFill 액션으로 배경색 시도
// ================================================================
run('\nT3 CellBorderFill 배경색', function() {
  // A1 셀에서 (이미 있음)
  var act = bridge.comCallWith(h, 'CreateAction', ['CellBorderFill']);
  var set = bridge.comCallWith(act, 'CreateSet', []);
  bridge.comCallWith(act, 'GetDefault', [set]);

  // Set 멤버 확인
  var members = bridge.comListMembers(set);
  var names = members.filter(function(m) {
    return ['QueryInterface','AddRef','Release','GetTypeInfoCount','GetTypeInfo','GetIDsOfNames','Invoke'].indexOf(m.name) === -1;
  }).map(function(m) { return m.name; });
  console.log('Set 멤버: ' + names.join(', '));

  // Item으로 값 탐색
  var items = ['FillColor', 'BackColor', 'FillType', 'PatternColor',
               'GradationType', 'WindowBrush', 'FillColorType',
               'BorderType', 'BorderWidth', 'BorderColor'];
  items.forEach(function(name) {
    try {
      var exists = bridge.comCallWith(set, 'ItemExist', [name]);
      if (exists) {
        var v = bridge.comCallWith(set, 'Item', [name]);
        console.log('  ' + name + ' = ' + v);
      }
    } catch(e) {}
  });
});

// ================================================================
// T4: HParameterSet.HCellBorderFill 직접 접근
// ================================================================
run('\nT4 HParameterSet.HCellBorderFill', function() {
  var hps = bridge.comGet(h, 'HParameterSet');

  // HCellBorderFill 시도
  try {
    var hcbf = bridge.comGet(hps, 'HCellBorderFill');
    console.log('HCellBorderFill type=' + typeof hcbf);
    var members = bridge.comListMembers(hcbf);
    var props = members.filter(function(m) {
      return m.kind === 'put' || m.kind === 'get/put' || m.kind === 'get';
    }).filter(function(m) {
      return m.name !== 'HSet';
    });
    props.forEach(function(m) {
      console.log('  ' + m.name + ' (' + m.kind + ')');
      try {
        var v = bridge.comGet(hcbf, m.name);
        if (v !== null && v !== undefined && typeof v !== 'object') {
          console.log('    = ' + v);
        }
      } catch(e) {}
    });
  } catch(e) {
    console.log('❌ ' + e.message);
    // 다른 이름 시도
    try {
      var htcp = bridge.comGet(hps, 'HTableCellProperty');
      console.log('HTableCellProperty 존재!');
    } catch(e2) {}
    try {
      var htcp = bridge.comGet(hps, 'HTableProperty');
      console.log('HTableProperty 존재!');
    } catch(e2) {}
  }
});

// ================================================================
// T5: HAction.GetDefault("CellBorderFill") + HParameterSet
// ================================================================
run('\nT5 HAction CellBorderFill', function() {
  hwp.HAction.GetDefault('CellBorderFill', hwp.HParameterSet.HCellBorderFill.HSet);
  var hcbf = hwp.HParameterSet.HCellBorderFill;

  // 멤버 덤프
  var hps = bridge.comGet(h, 'HParameterSet');
  var handle = bridge.comGet(hps, 'HCellBorderFill');
  var members = bridge.comListMembers(handle);
  var props = members.filter(function(m) {
    return (m.kind === 'put' || m.kind === 'get/put') && m.name !== 'HSet';
  });

  console.log('프로퍼티:');
  props.forEach(function(m) {
    try {
      var v = bridge.comGet(handle, m.name);
      if (typeof v !== 'object') {
        console.log('  ' + m.name + ' = ' + v);
      } else {
        console.log('  ' + m.name + ' = [object]');
      }
    } catch(e) {
      console.log('  ' + m.name + ' = (읽기 실패)');
    }
  });
});

// ================================================================
// T6: 배경색 설정 시도
// ================================================================
run('\nT6 배경색 설정 시도', function() {
  // A1 셀에서
  hwp.HAction.GetDefault('CellBorderFill', hwp.HParameterSet.HCellBorderFill.HSet);

  // FillAttr.FaceColor 같은 중첩 구조일 수 있음
  var hps = bridge.comGet(h, 'HParameterSet');
  var handle = bridge.comGet(hps, 'HCellBorderFill');

  // 가능한 프로퍼티 시도
  var tryProps = [
    ['FillColor', bridge.comCallWith(h, 'RGBColor', [255, 200, 200])],
    ['BackColor', bridge.comCallWith(h, 'RGBColor', [255, 200, 200])],
    ['FaceColor', bridge.comCallWith(h, 'RGBColor', [255, 200, 200])],
  ];

  tryProps.forEach(function(p) {
    try {
      bridge.comPut(handle, p[0], p[1]);
      console.log('  ' + p[0] + ' 설정 성공');
    } catch(e) {
      console.log('  ' + p[0] + ' ❌');
    }
  });

  hwp.HAction.Execute('CellBorderFill', hwp.HParameterSet.HCellBorderFill.HSet);
  console.log('  Execute 완료 — 화면 확인');
});

// 닫지 않음
console.log('\n확인해주세요!');
