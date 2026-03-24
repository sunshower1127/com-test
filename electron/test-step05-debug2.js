/**
 * 단계 5 디버그 2: UnderlineType/StrikeOutType + 서식 읽기
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

// 텍스트 준비 + 볼드 적용
insert('ABCDEFGHIJ');
bridge.comCallWith(h, 'Run', ['MoveDocBegin']);
for (var i = 0; i < 3; i++) bridge.comCallWith(h, 'Run', ['MoveSelRight']); // ABC
hwp.HAction.GetDefault('CharShape', hwp.HParameterSet.HCharShape.HSet);
hwp.HParameterSet.HCharShape.Bold = 1;
hwp.HParameterSet.HCharShape.Height = 2400;
hwp.HAction.Execute('CharShape', hwp.HParameterSet.HCharShape.HSet);
bridge.comCallWith(h, 'Run', ['Cancel']);
console.log('ABC에 24pt 볼드 적용\n');

// ── UnderlineType ──
console.log('=== UnderlineType ===');
bridge.comCallWith(h, 'Run', ['MoveDocBegin']);
for (var i = 0; i < 3; i++) bridge.comCallWith(h, 'Run', ['MoveRight']);
for (var i = 0; i < 3; i++) bridge.comCallWith(h, 'Run', ['MoveSelRight']); // DEF
hwp.HAction.GetDefault('CharShape', hwp.HParameterSet.HCharShape.HSet);
hwp.HParameterSet.HCharShape.UnderlineType = 1;
hwp.HAction.Execute('CharShape', hwp.HParameterSet.HCharShape.HSet);
bridge.comCallWith(h, 'Run', ['Cancel']);
console.log('  DEF에 밑줄 적용 ✅');

// ── StrikeOutType ──
console.log('\n=== StrikeOutType ===');
bridge.comCallWith(h, 'Run', ['MoveDocBegin']);
for (var i = 0; i < 6; i++) bridge.comCallWith(h, 'Run', ['MoveRight']);
for (var i = 0; i < 3; i++) bridge.comCallWith(h, 'Run', ['MoveSelRight']); // GHI
hwp.HAction.GetDefault('CharShape', hwp.HParameterSet.HCharShape.HSet);
hwp.HParameterSet.HCharShape.StrikeOutType = 1;
hwp.HAction.Execute('CharShape', hwp.HParameterSet.HCharShape.HSet);
bridge.comCallWith(h, 'Run', ['Cancel']);
console.log('  GHI에 취소선 적용 ✅');

// ── 서식 읽기: hwp.CharShape (ParameterSet) ──
console.log('\n=== 서식 읽기 시도 ===');

// ABC (볼드 24pt) 위치로
bridge.comCallWith(h, 'Run', ['MoveDocBegin']);
bridge.comCallWith(h, 'Run', ['MoveRight']); // B 위치

// 방법 1: hwp.CharShape.Item()
try {
  var cs = bridge.comGet(h, 'CharShape');
  var height = bridge.comCallWith(cs, 'Item', ['Height']);
  var bold = bridge.comCallWith(cs, 'Item', ['Bold']);
  var italic = bridge.comCallWith(cs, 'Item', ['Italic']);
  var color = bridge.comCallWith(cs, 'Item', ['TextColor']);
  console.log('hwp.CharShape.Item: Height=' + height + ', Bold=' + bold +
    ', Italic=' + italic + ', Color=' + color);
} catch(e) { console.log('hwp.CharShape.Item: ❌ ' + e.message); }

// DEF (밑줄) 위치
bridge.comCallWith(h, 'Run', ['MoveDocBegin']);
for (var i = 0; i < 4; i++) bridge.comCallWith(h, 'Run', ['MoveRight']);
try {
  var cs = bridge.comGet(h, 'CharShape');
  var underline = bridge.comCallWith(cs, 'Item', ['UnderlineType']);
  var height = bridge.comCallWith(cs, 'Item', ['Height']);
  console.log('DEF위치: UnderlineType=' + underline + ', Height=' + height);
} catch(e) { console.log('DEF: ❌ ' + e.message); }

// 일반 텍스트 위치 (HIJ)
bridge.comCallWith(h, 'Run', ['MoveDocEnd']);
bridge.comCallWith(h, 'Run', ['MoveLeft']);
try {
  var cs = bridge.comGet(h, 'CharShape');
  var height = bridge.comCallWith(cs, 'Item', ['Height']);
  var bold = bridge.comCallWith(cs, 'Item', ['Bold']);
  console.log('J위치(기본): Height=' + height + ', Bold=' + bold);
} catch(e) { console.log('J: ❌ ' + e.message); }

// 종료
console.log('\n완료');
bridge.comCallWith(h, 'Clear', [1]);
bridge.comCallWith(h, 'Quit', []);
