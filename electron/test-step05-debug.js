/**
 * 단계 5 디버그: Underline/StrikeOut 프로퍼티명 + 서식 읽기
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

// ── HCharShape 멤버에서 Underline/StrikeOut 관련 찾기 ──
console.log('=== HCharShape 멤버 (Underline/Strike 관련) ===');
var hps = bridge.comGet(h, 'HParameterSet');
var hcs = bridge.comGet(hps, 'HCharShape');
var members = bridge.comListMembers(hcs);
var filtered = members.filter(function(m) {
  var n = m.name.toLowerCase();
  return n.indexOf('under') >= 0 || n.indexOf('strike') >= 0 || n.indexOf('line') >= 0;
});
filtered.forEach(function(m) {
  console.log('  ' + m.name + ' (' + m.kind + ')');
});

// 전체 멤버 중 put 가능한 것만 (서식 설정 가능한 프로퍼티)
console.log('\n=== HCharShape put 가능 프로퍼티 (서식 관련) ===');
var putProps = members.filter(function(m) {
  return m.kind === 'put' || m.kind === 'get/put';
});
putProps.forEach(function(m) {
  console.log('  ' + m.name);
});

// ── 서식 읽기 디버그 ──
console.log('\n=== 서식 읽기 디버그 ===');
insert('TEST TEXT');

// 선택 후 볼드 적용
bridge.comCallWith(h, 'Run', ['MoveDocBegin']);
for (var i = 0; i < 4; i++) bridge.comCallWith(h, 'Run', ['MoveSelRight']); // TEST 선택
hwp.HAction.GetDefault('CharShape', hwp.HParameterSet.HCharShape.HSet);
hwp.HParameterSet.HCharShape.Bold = 1;
hwp.HParameterSet.HCharShape.Height = 2000;
hwp.HAction.Execute('CharShape', hwp.HParameterSet.HCharShape.HSet);
bridge.comCallWith(h, 'Run', ['Cancel']);

// 방법 1: GetDefault 후 읽기
bridge.comCallWith(h, 'Run', ['MoveDocBegin']);
bridge.comCallWith(h, 'Run', ['MoveRight']); // E 뒤 (TEST 안)
hwp.HAction.GetDefault('CharShape', hwp.HParameterSet.HCharShape.HSet);
console.log('방법1 (GetDefault): Height=' + hwp.HParameterSet.HCharShape.Height +
  ', Bold=' + hwp.HParameterSet.HCharShape.Bold);

// 방법 2: CreateAction으로 읽기
var act = hwp.CreateAction('CharShape');
var set = act.CreateSet();
act.GetDefault(set);
try {
  var height = bridge.comCallWith(set, 'Item', ['Height']);
  var bold = bridge.comCallWith(set, 'Item', ['Bold']);
  console.log('방법2 (CreateAction): Height=' + height + ', Bold=' + bold);
} catch(e) { console.log('방법2: ❌ ' + e.message); }

// 방법 3: hwp.CharShape 프로퍼티
try {
  var cs = bridge.comGet(h, 'CharShape');
  console.log('방법3 (hwp.CharShape): type=' + typeof cs);
  if (cs && typeof cs === 'object') {
    var csMembers = bridge.comListMembers(cs);
    var csNames = csMembers.filter(function(m) {
      return ['QueryInterface','AddRef','Release','GetTypeInfoCount','GetTypeInfo','GetIDsOfNames','Invoke'].indexOf(m.name) === -1;
    }).map(function(m) { return m.name + '(' + m.kind + ')'; });
    console.log('  멤버: ' + csNames.join(', '));
  }
} catch(e) { console.log('방법3: ❌ ' + e.message); }

// 종료
console.log('\n완료');
bridge.comCallWith(h, 'Clear', [1]);
bridge.comCallWith(h, 'Quit', []);
