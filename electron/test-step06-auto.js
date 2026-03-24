/**
 * 단계 6: 문단 서식 (ParagraphShape) — 자동 테스트
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

// 5개 문단 준비
insert('Left aligned paragraph\r\nCenter aligned paragraph\r\nRight aligned paragraph\r\nIndented paragraph\r\nSpaced paragraph');
console.log('5개 문단 준비 완료\n');

// ── HParaShape 멤버 확인 ──
console.log('=== HParaShape put 가능 프로퍼티 ===');
var hps = bridge.comGet(h, 'HParameterSet');
var hpara = bridge.comGet(hps, 'HParaShape');
var members = bridge.comListMembers(hpara);
var putProps = members.filter(function(m) {
  return m.kind === 'put' || m.kind === 'get/put';
});
putProps.forEach(function(m) {
  console.log('  ' + m.name);
});

// 문단 선택 헬퍼: 해당 문단으로 이동 후 전체 선택
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

// ================================================================
// 6-1: Alignment (정렬)
// ================================================================
console.log('');

run('6-1 Center (Alignment)', function() {
  selectPara(1);
  applyParaShape('Alignment', 3); // 가운데
  bridge.comCallWith(h, 'Run', ['Cancel']);
  console.log('✅ 문단2 가운데 정렬');
});

run('6-2 Right (Alignment)', function() {
  selectPara(2);
  applyParaShape('Alignment', 2); // 오른쪽
  bridge.comCallWith(h, 'Run', ['Cancel']);
  console.log('✅ 문단3 오른쪽 정렬');
});

// ================================================================
// 6-3: LeftMargin (왼쪽 들여쓰기)
// ================================================================
run('6-3 LeftMargin', function() {
  selectPara(3);
  applyParaShape('LeftMargin', 2000); // 약 7mm
  bridge.comCallWith(h, 'Run', ['Cancel']);
  console.log('✅ 문단4 왼쪽 들여쓰기');
});

// ================================================================
// 6-4: LineSpacing (줄간격)
// ================================================================
run('6-4 LineSpacing', function() {
  selectPara(4);
  // LineSpacing 관련 프로퍼티 시도
  try {
    applyParaShape('LineSpacing', 300); // 300%?
    console.log('✅ 문단5 줄간격 300');
  } catch(e) {
    console.log('LineSpacing 실패: ' + e.message);
    try {
      applyParaShape('LineSpacingValue', 300);
      console.log('✅ LineSpacingValue=300');
    } catch(e2) {
      console.log('LineSpacingValue도 실패: ' + e2.message);
    }
  }
  bridge.comCallWith(h, 'Run', ['Cancel']);
});

// ================================================================
// 6-5: SpaceBeforePara / SpaceAfterPara
// ================================================================
run('6-5 SpaceBeforePara/SpaceAfterPara', function() {
  selectPara(0);
  try {
    applyParaShape('SpaceBeforePara', 500);
    applyParaShape('SpaceAfterPara', 500);
    console.log('✅ 문단1 앞뒤 간격');
  } catch(e) {
    console.log('실패: ' + e.message);
    // 다른 이름 시도
    try {
      applyParaShape('SpaceBefore', 500);
      applyParaShape('SpaceAfter', 500);
      console.log('✅ SpaceBefore/After=500');
    } catch(e2) {
      console.log('SpaceBefore/After도 실패: ' + e2.message);
    }
  }
  bridge.comCallWith(h, 'Run', ['Cancel']);
});

// ================================================================
// 6-6: 서식 읽기 — hwp.ParaShape
// ================================================================
run('6-6 서식 읽기 (hwp.ParaShape)', function() {
  // 가운데 정렬 문단으로 이동
  bridge.comCallWith(h, 'SetPos', [0, 1, 0]);
  bridge.comCallWith(h, 'Run', ['MoveRight']);

  try {
    var ps = bridge.comGet(h, 'ParaShape');
    var align = bridge.comCallWith(ps, 'Item', ['Alignment']);
    console.log('문단2: Alignment=' + align + ' (기대: 3=Center)');
  } catch(e) {
    console.log('❌ ' + e.message);
  }

  // 오른쪽 정렬 문단
  bridge.comCallWith(h, 'SetPos', [0, 2, 0]);
  bridge.comCallWith(h, 'Run', ['MoveRight']);
  try {
    var ps = bridge.comGet(h, 'ParaShape');
    var align = bridge.comCallWith(ps, 'Item', ['Alignment']);
    var leftMargin = bridge.comCallWith(ps, 'Item', ['LeftMargin']);
    console.log('         문단3: Alignment=' + align + ' (기대: 2=Right)');
  } catch(e) {
    console.log('         ❌ ' + e.message);
  }
});

// ================================================================
// 6-7: CreateAction 패턴
// ================================================================
run('6-7 CreateAction("ParagraphShape")', function() {
  selectPara(0);
  var act = hwp.CreateAction('ParagraphShape');
  var set = act.CreateSet();
  act.GetDefault(set);
  set.SetItem('Alignment', 3); // 가운데
  act.Execute(set);
  bridge.comCallWith(h, 'Run', ['Cancel']);
  console.log('✅ 문단1도 가운데 정렬');
});

// 종료하지 않음 — 눈으로 확인
console.log('\n서식 적용 완료! 한글 창에서 확인해주세요.');
console.log('(수동으로 닫으세요)');
