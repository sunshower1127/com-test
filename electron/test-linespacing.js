/**
 * LineSpacing 디버그: 타입별 동작 + 값 매핑
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

function applyParaShape(props) {
  hwp.HAction.GetDefault('ParagraphShape', hwp.HParameterSet.HParaShape.HSet);
  Object.keys(props).forEach(function(k) {
    hwp.HParameterSet.HParaShape[k] = props[k];
  });
  hwp.HAction.Execute('ParagraphShape', hwp.HParameterSet.HParaShape.HSet);
}

function readParaShape() {
  var ps = bridge.comGet(h, 'ParaShape');
  return {
    lineSpacing: bridge.comCallWith(ps, 'Item', ['LineSpacing']),
    lineSpacingType: bridge.comCallWith(ps, 'Item', ['LineSpacingType']),
  };
}

function selectPara(para) {
  bridge.comCallWith(h, 'SetPos', [0, para, 0]);
  bridge.comCallWith(h, 'Run', ['MoveParaBegin']);
  bridge.comCallWith(h, 'Run', ['MoveSelParaEnd']);
}

// ── T1: 기본 문서의 LineSpacing 값 ──
console.log('=== T1: 기본 LineSpacing ===');
insert('Default line spacing');
bridge.comCallWith(h, 'Run', ['MoveDocBegin']);
bridge.comCallWith(h, 'Run', ['MoveRight']);
var r = readParaShape();
console.log('  기본값: LineSpacing=' + r.lineSpacing + ', Type=' + r.lineSpacingType);

// ── T2: 다양한 LineSpacing 값 설정 후 읽기 ──
console.log('\n=== T2: LineSpacing 값별 테스트 (Type 안 건드림) ===');
var values = [100, 150, 160, 200, 250, 300, 400];
values.forEach(function(v) {
  bridge.comCallWith(h, 'Run', ['SelectAll']);
  bridge.comCallWith(h, 'Run', ['Delete']);
  insert('Test ' + v);
  selectPara(0);
  applyParaShape({ LineSpacing: v });
  bridge.comCallWith(h, 'Run', ['Cancel']);

  bridge.comCallWith(h, 'Run', ['MoveDocBegin']);
  bridge.comCallWith(h, 'Run', ['MoveRight']);
  var r = readParaShape();
  console.log('  설정=' + v + ' → 읽기: spacing=' + r.lineSpacing + ', type=' + r.lineSpacingType +
    (v === r.lineSpacing ? ' ✅' : ' ⚠️'));
});

// ── T3: LineSpacingType별 테스트 ──
console.log('\n=== T3: LineSpacingType별 테스트 ===');
for (var type = 0; type <= 4; type++) {
  bridge.comCallWith(h, 'Run', ['SelectAll']);
  bridge.comCallWith(h, 'Run', ['Delete']);
  insert('Type ' + type + ' test');
  selectPara(0);
  try {
    applyParaShape({ LineSpacingType: type, LineSpacing: 200 });
    bridge.comCallWith(h, 'Run', ['Cancel']);

    bridge.comCallWith(h, 'Run', ['MoveDocBegin']);
    bridge.comCallWith(h, 'Run', ['MoveRight']);
    var r = readParaShape();
    console.log('  Type=' + type + ', 설정=200 → 읽기: spacing=' + r.lineSpacing + ', type=' + r.lineSpacingType);
  } catch(e) {
    console.log('  Type=' + type + ': ❌ ' + e.message);
    bridge.comCallWith(h, 'Run', ['Cancel']);
  }
}

// ── T4: LineSpacingType=0으로 고정 후 다양한 값 ──
console.log('\n=== T4: Type=0 (비율%) 고정, 값 변경 ===');
var vals = [100, 130, 150, 160, 200, 250, 300];
vals.forEach(function(v) {
  bridge.comCallWith(h, 'Run', ['SelectAll']);
  bridge.comCallWith(h, 'Run', ['Delete']);
  insert('Pct ' + v);
  selectPara(0);
  applyParaShape({ LineSpacingType: 0, LineSpacing: v });
  bridge.comCallWith(h, 'Run', ['Cancel']);

  bridge.comCallWith(h, 'Run', ['MoveDocBegin']);
  bridge.comCallWith(h, 'Run', ['MoveRight']);
  var r = readParaShape();
  console.log('  설정=' + v + ' → 읽기: spacing=' + r.lineSpacing + ', type=' + r.lineSpacingType +
    (v === r.lineSpacing ? ' ✅' : ' ⚠️'));
});

// ── T5: CreateAction으로 LineSpacing 설정 ──
console.log('\n=== T5: CreateAction 패턴 ===');
bridge.comCallWith(h, 'Run', ['SelectAll']);
bridge.comCallWith(h, 'Run', ['Delete']);
insert('CreateAction test');
selectPara(0);
var act = hwp.CreateAction('ParagraphShape');
var set = act.CreateSet();
act.GetDefault(set);
set.SetItem('LineSpacingType', 0);
set.SetItem('LineSpacing', 200);
act.Execute(set);
bridge.comCallWith(h, 'Run', ['Cancel']);

bridge.comCallWith(h, 'Run', ['MoveDocBegin']);
bridge.comCallWith(h, 'Run', ['MoveRight']);
var r = readParaShape();
console.log('  CreateAction(Type=0, Spacing=200) → 읽기: spacing=' + r.lineSpacing + ', type=' + r.lineSpacingType);

// 종료
console.log('\n완료');
bridge.comCallWith(h, 'Clear', [1]);
bridge.comCallWith(h, 'Quit', []);
