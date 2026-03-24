/**
 * 단계 6: 문단 서식 — 수정된 프로퍼티명으로 재테스트
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

function selectPara(para) {
  bridge.comCallWith(h, 'SetPos', [0, para, 0]);
  bridge.comCallWith(h, 'Run', ['MoveParaBegin']);
  bridge.comCallWith(h, 'Run', ['MoveSelParaEnd']);
}

function applyParaShape(propName, value) {
  hwp.HAction.GetDefault('ParagraphShape', hwp.HParameterSet.HParaShape.HSet);
  hwp.HParameterSet.HParaShape[propName] = value;
  hwp.HAction.Execute('ParagraphShape', hwp.HParameterSet.HParaShape.HSet);
}

// 5개 문단
insert('Left aligned (default)\r\nCenter aligned\r\nRight aligned\r\nIndented + spaced\r\nWide line spacing');

// 6-1: 가운데 정렬 (AlignType)
run('6-1 AlignType=3 (Center)', function() {
  selectPara(1);
  applyParaShape('AlignType', 3);
  bridge.comCallWith(h, 'Run', ['Cancel']);
  console.log('✅');
});

// 6-2: 오른쪽 정렬
run('6-2 AlignType=2 (Right)', function() {
  selectPara(2);
  applyParaShape('AlignType', 2);
  bridge.comCallWith(h, 'Run', ['Cancel']);
  console.log('✅');
});

// 6-3: 왼쪽 들여쓰기
run('6-3 LeftMargin=2000', function() {
  selectPara(3);
  applyParaShape('LeftMargin', 2000);
  bridge.comCallWith(h, 'Run', ['Cancel']);
  console.log('✅');
});

// 6-4: 문단 간격 (PrevSpacing/NextSpacing)
run('6-4 PrevSpacing/NextSpacing=500', function() {
  selectPara(3);
  applyParaShape('PrevSpacing', 500);
  applyParaShape('NextSpacing', 500);
  bridge.comCallWith(h, 'Run', ['Cancel']);
  console.log('✅');
});

// 6-5: 줄간격
run('6-5 LineSpacing=250', function() {
  selectPara(4);
  applyParaShape('LineSpacing', 250);
  bridge.comCallWith(h, 'Run', ['Cancel']);
  console.log('✅');
});

// 6-6: 들여쓰기 (Indentation)
run('6-6 Indentation=1000', function() {
  selectPara(0);
  applyParaShape('Indentation', 1000);
  bridge.comCallWith(h, 'Run', ['Cancel']);
  console.log('✅');
});

// 6-7: 서식 읽기 — hwp.ParaShape.Item()
run('6-7 서식 읽기', function() {
  // 가운데 정렬 문단
  bridge.comCallWith(h, 'SetPos', [0, 1, 0]);
  bridge.comCallWith(h, 'Run', ['MoveRight']);
  var ps = bridge.comGet(h, 'ParaShape');
  var align = bridge.comCallWith(ps, 'Item', ['AlignType']);
  var leftMargin = bridge.comCallWith(ps, 'Item', ['LeftMargin']);
  var lineSpacing = bridge.comCallWith(ps, 'Item', ['LineSpacing']);
  console.log('문단2: AlignType=' + align + ' (기대: 3)');

  // 오른쪽 정렬 문단
  bridge.comCallWith(h, 'SetPos', [0, 2, 0]);
  bridge.comCallWith(h, 'Run', ['MoveRight']);
  ps = bridge.comGet(h, 'ParaShape');
  align = bridge.comCallWith(ps, 'Item', ['AlignType']);
  console.log('         문단3: AlignType=' + align + ' (기대: 2)');

  // 들여쓰기 문단
  bridge.comCallWith(h, 'SetPos', [0, 3, 0]);
  bridge.comCallWith(h, 'Run', ['MoveRight']);
  ps = bridge.comGet(h, 'ParaShape');
  leftMargin = bridge.comCallWith(ps, 'Item', ['LeftMargin']);
  var prev = bridge.comCallWith(ps, 'Item', ['PrevSpacing']);
  var next = bridge.comCallWith(ps, 'Item', ['NextSpacing']);
  console.log('         문단4: LeftMargin=' + leftMargin + ', Prev=' + prev + ', Next=' + next);

  // 줄간격 문단
  bridge.comCallWith(h, 'SetPos', [0, 4, 0]);
  bridge.comCallWith(h, 'Run', ['MoveRight']);
  ps = bridge.comGet(h, 'ParaShape');
  lineSpacing = bridge.comCallWith(ps, 'Item', ['LineSpacing']);
  var lineType = bridge.comCallWith(ps, 'Item', ['LineSpacingType']);
  console.log('         문단5: LineSpacing=' + lineSpacing + ', LineSpacingType=' + lineType);
});

// 6-8: TextAlignment (세로 정렬?)
run('6-8 TextAlignment', function() {
  selectPara(0);
  // TextAlignment 값 확인
  var ps = bridge.comGet(h, 'ParaShape');
  var ta = bridge.comCallWith(ps, 'Item', ['TextAlignment']);
  console.log('TextAlignment=' + ta);
  bridge.comCallWith(h, 'Run', ['Cancel']);
});

console.log('\n확인해주세요! (수동으로 닫으세요)');
