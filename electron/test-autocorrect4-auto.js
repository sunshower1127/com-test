/**
 * 자동교정 4차: 빠른 교정(Quick Correct) 비활성화 방법 탐색
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
  var ok = r.indexOf('첫번째') >= 0; // 교정 안 됐으면 "첫번째" 그대로
  console.log('  결과: "' + r + '" ' + (ok ? '✅ 교정 안 됨!' : '⚠️ 교정됨'));
  return ok;
}

console.log('HWP Version: ' + String(hwp.Version) + '\n');

// ── T1: HQCorrect 파라미터셋 탐색 ──
console.log('── T1: HParameterSet.HQCorrect 접근');
try {
  var qc = hwp.HParameterSet.HQCorrect;
  console.log('  type=' + typeof qc);
  // HSet 접근
  try {
    var hset = qc.HSet;
    console.log('  HSet type=' + typeof hset);
  } catch(e) { console.log('  HSet 실패: ' + e.message); }
} catch(e) { console.log('  ❌ ' + e.message); }

// ── T2: QuickCorrect 관련 액션 이름 탐색 ──
console.log('\n── T2: CreateAction 액션 이름 탐색');
var actions = [
  'QuickCorrect', 'QuickCorrectAll', 'QuickCorrectDic',
  'QCorrect', 'AutoCorrectOption', 'SpellOption',
  'SpellingCheck', 'OptionDlg', 'SetupAutoCorrect',
  'ToolsAutoCorrect', 'AutoCorrect', 'SpellCorrect',
  'SetAutoSpell', 'ToggleQuickCorrect'
];
actions.forEach(function(name) {
  try {
    var act = bridge.comCallWith(h, 'CreateAction', [name]);
    if (act) {
      console.log('  ' + name + ': ✅ 존재');
      // set 만들어보기
      try {
        var set = bridge.comCallWith(act, 'CreateSet', []);
        console.log('    CreateSet: ' + (set ? '✅' : '❌ null'));
      } catch(e2) {}
    } else {
      console.log('  ' + name + ': null');
    }
  } catch(e) {
    console.log('  ' + name + ': ❌ ' + e.message);
  }
});

// ── T3: IsActionEnable 확인 ──
console.log('\n── T3: IsActionEnable 확인');
var moreActions = [
  'AutoSpellCheck', 'ToggleAutoCorrect', 'QuickCorrectAll',
  'QuickCorrectDic', 'SpellCheck', 'AutoCorrectOption'
];
moreActions.forEach(function(name) {
  try {
    var en = bridge.comCallWith(h, 'IsActionEnable', [name]);
    console.log('  ' + name + ': ' + String(en));
  } catch(e) {
    console.log('  ' + name + ': ❌');
  }
});

// ── T4: HAction.GetDefault로 QuickCorrectAll 시도 ──
console.log('\n── T4: HAction.GetDefault("QuickCorrectAll", HQCorrect.HSet)');
try {
  hwp.HAction.GetDefault('QuickCorrectAll', hwp.HParameterSet.HQCorrect.HSet);
  console.log('  GetDefault 성공');
  // HQCorrect 멤버 나열 시도
  var qc = hwp.HParameterSet.HQCorrect;
  var props = [
    'AutoCorrect', 'Enable', 'UseAutoCorrect', 'AutoSpacing',
    'SpacingCorrect', 'UseSpacingCorrect', 'Active', 'Flag',
    'Flags', 'Type', 'UseQuickCorrect', 'Mode', 'Option',
    'SpellCheck', 'AutoFormat', 'CorrectSpacing'
  ];
  props.forEach(function(p) {
    try {
      var val = bridge.comGet(qc._handle || bridge.comGet(h, 'HParameterSet'), p);
      console.log('  .' + p + ' = ' + String(val));
    } catch(e) {}
  });
} catch(e) { console.log('  ❌ ' + e.message); }

// ── T5: HQCorrect를 Proxy로 접근해서 프로퍼티 읽기 ──
console.log('\n── T5: HQCorrect 프로퍼티 직접 읽기');
try {
  var qcProps = [
    'AutoCorrect', 'Enable', 'UseAutoCorrect', 'AutoSpacing',
    'SpacingCorrect', 'UseSpacingCorrect', 'Active', 'Flag',
    'Flags', 'Type', 'UseQuickCorrect', 'Mode', 'Option',
    'CorrectSpacing', 'AutoFormatAsYouType', 'SuggestCorrection',
    'AutoText', 'AutoTextEntry', 'EngCapital', 'EngCapitalSmall',
    'HangulToHanja', 'UseAutoSpace', 'UseAutoSpacing'
  ];
  var qc = hwp.HParameterSet.HQCorrect;
  qcProps.forEach(function(p) {
    try {
      var val = qc[p];
      if (val !== undefined) {
        console.log('  .' + p + ' = ' + String(val) + ' (' + typeof val + ')');
      }
    } catch(e) {}
  });
} catch(e) { console.log('  ❌ ' + e.message); }

// ── T6: list_members 스타일로 ITypeInfo 덤프 (Rust 없이는 안 되지만 시도) ──
// bridge에 listMembers 같은 게 있는지 확인
console.log('\n── T6: bridge에 listMembers 있는지');
console.log('  bridge keys: ' + Object.keys(bridge).join(', '));

// ── T7: ToggleAutoCorrect 2번 토글 후 상태 확인 ──
console.log('\n── T7: ToggleAutoCorrect 상태 확인');
console.log('  현재 상태에서 BreakPara 테스트:');
testBreakPara();
console.log('  ToggleAutoCorrect 1회:');
bridge.comCallWith(h, 'Run', ['ToggleAutoCorrect']);
testBreakPara();
console.log('  ToggleAutoCorrect 2회 (복구):');
bridge.comCallWith(h, 'Run', ['ToggleAutoCorrect']);
testBreakPara();

// 종료
console.log('\n── 완료');
bridge.comCallWith(h, 'Clear', [1]);
bridge.comCallWith(h, 'Quit', []);
