/**
 * 자동교정 5차: comListMembers로 HQCorrect 멤버 덤프 + QuickCorrect 액션 탐색
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

function testBreakPara() {
  clear();
  insert('첫번째');
  bridge.comCallWith(h, 'Run', ['BreakPara']);
  insert('끝');
  var r = getText();
  var ok = r.indexOf('첫번째') >= 0;
  console.log('    "' + r + '" ' + (ok ? '✅' : '⚠️'));
  return ok;
}

console.log('HWP Version: ' + String(hwp.Version) + '\n');

// ── T1: HQCorrect 멤버 덤프 ──
console.log('── T1: HParameterSet.HQCorrect 멤버');
try {
  var hps = bridge.comGet(h, 'HParameterSet');
  var qcHandle = bridge.comGet(hps, 'HQCorrect');
  console.log('  HQCorrect handle type: ' + typeof qcHandle);
  var members = bridge.comListMembers(qcHandle);
  console.log('  멤버 수: ' + members.length);
  members.forEach(function(m) {
    console.log('  - ' + JSON.stringify(m));
  });
} catch(e) { console.log('  ❌ ' + e.message); }

// ── T2: QuickCorrect 액션의 Set 멤버 ──
console.log('\n── T2: CreateAction("QuickCorrect") Set 멤버');
try {
  var act = bridge.comCallWith(h, 'CreateAction', ['QuickCorrect']);
  var set = bridge.comCallWith(act, 'CreateSet', []);
  var members = bridge.comListMembers(set);
  console.log('  Set 멤버 수: ' + members.length);
  members.forEach(function(m) {
    console.log('  - ' + JSON.stringify(m));
  });
  // GetDefault 후 값 읽기
  bridge.comCallWith(act, 'GetDefault', [set]);
  console.log('\n  GetDefault 후 값 읽기:');
  members.forEach(function(m) {
    if (m.name && m.name !== 'QueryInterface' && m.name !== 'AddRef' && m.name !== 'Release'
        && m.name !== 'GetTypeInfoCount' && m.name !== 'GetTypeInfo'
        && m.name !== 'GetIDsOfNames' && m.name !== 'Invoke') {
      try {
        var val = bridge.comGet(set, m.name);
        console.log('  .' + m.name + ' = ' + String(val) + ' (' + typeof val + ')');
      } catch(e) {}
    }
  });
} catch(e) { console.log('  ❌ ' + e.message); }

// ── T3: SpellingCheck 액션도 덤프 ──
console.log('\n── T3: CreateAction("SpellingCheck") Set 멤버');
try {
  var act2 = bridge.comCallWith(h, 'CreateAction', ['SpellingCheck']);
  var set2 = bridge.comCallWith(act2, 'CreateSet', []);
  var members2 = bridge.comListMembers(set2);
  console.log('  Set 멤버 수: ' + members2.length);
  members2.forEach(function(m) {
    console.log('  - ' + JSON.stringify(m));
  });
  bridge.comCallWith(act2, 'GetDefault', [set2]);
  console.log('\n  GetDefault 후 값 읽기:');
  members2.forEach(function(m) {
    if (m.name && m.name !== 'QueryInterface' && m.name !== 'AddRef' && m.name !== 'Release'
        && m.name !== 'GetTypeInfoCount' && m.name !== 'GetTypeInfo'
        && m.name !== 'GetIDsOfNames' && m.name !== 'Invoke') {
      try {
        var val = bridge.comGet(set2, m.name);
        console.log('  .' + m.name + ' = ' + String(val) + ' (' + typeof val + ')');
      } catch(e) {}
    }
  });
} catch(e) { console.log('  ❌ ' + e.message); }

// ── T4: HQCorrect 프로퍼티 값 읽고 0으로 세팅 시도 ──
console.log('\n── T4: HQCorrect 프로퍼티 값 읽기 + 비활성화 시도');
try {
  var hps = bridge.comGet(h, 'HParameterSet');
  var qcHandle = bridge.comGet(hps, 'HQCorrect');
  var members = bridge.comListMembers(qcHandle);

  // 값 읽기
  members.forEach(function(m) {
    if (m.name && m.name !== 'QueryInterface' && m.name !== 'AddRef' && m.name !== 'Release'
        && m.name !== 'GetTypeInfoCount' && m.name !== 'GetTypeInfo'
        && m.name !== 'GetIDsOfNames' && m.name !== 'Invoke') {
      try {
        var val = bridge.comGet(qcHandle, m.name);
        console.log('  .' + m.name + ' = ' + String(val) + ' (' + typeof val + ')');
      } catch(e) {}
    }
  });

  // 비활성화 시도: 가능한 프로퍼티들을 0/false로
  console.log('\n  비활성화 시도...');
  members.forEach(function(m) {
    if (m.invoke_kind === 'PROPERTYPUT' || m.invoke_kind === 'PROPERTYGET|PROPERTYPUT') {
      try {
        bridge.comPut(qcHandle, m.name, 0);
        console.log('  .' + m.name + ' = 0 설정 성공');
      } catch(e) {}
      try {
        bridge.comPut(qcHandle, m.name, false);
        console.log('  .' + m.name + ' = false 설정 성공');
      } catch(e) {}
    }
  });

  console.log('\n  비활성화 후 BreakPara 테스트:');
  testBreakPara();
} catch(e) { console.log('  ❌ ' + e.message); }

// 종료
console.log('\n── 완료');
bridge.comCallWith(h, 'Clear', [1]);
bridge.comCallWith(h, 'Quit', []);
