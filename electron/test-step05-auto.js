/**
 * 단계 5: 글자 서식 (CharShape) — 자동 테스트
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

// 텍스트 준비
insert('ABCDEFGHIJ\r\nKLMNOPQRST');
console.log('준비: "ABCDEFGHIJ | KLMNOPQRST"\n');

// 선택 후 서식 적용하는 헬퍼
function applyCharShape(propName, value) {
  hwp.HAction.GetDefault('CharShape', hwp.HParameterSet.HCharShape.HSet);
  hwp.HParameterSet.HCharShape[propName] = value;
  hwp.HAction.Execute('CharShape', hwp.HParameterSet.HCharShape.HSet);
}

// ================================================================
// 5-1: Height (글자 크기)
// ================================================================
run('5-1 Height=2400 (24pt)', function() {
  bridge.comCallWith(h, 'Run', ['MoveDocBegin']);
  for (var i = 0; i < 3; i++) bridge.comCallWith(h, 'Run', ['MoveSelRight']); // ABC 선택
  applyCharShape('Height', 2400);
  bridge.comCallWith(h, 'Run', ['Cancel']);
  console.log('✅ ABC에 24pt 적용 → 화면에서 크기 확인');
});

// ================================================================
// 5-2: Bold
// ================================================================
run('5-2 Bold=1', function() {
  bridge.comCallWith(h, 'Run', ['MoveDocBegin']);
  for (var i = 0; i < 3; i++) bridge.comCallWith(h, 'Run', ['MoveRight']);
  for (var i = 0; i < 3; i++) bridge.comCallWith(h, 'Run', ['MoveSelRight']); // DEF 선택
  applyCharShape('Bold', 1);
  bridge.comCallWith(h, 'Run', ['Cancel']);
  console.log('✅ DEF에 볼드 적용');
});

// ================================================================
// 5-3: Italic
// ================================================================
run('5-3 Italic=1', function() {
  bridge.comCallWith(h, 'Run', ['MoveDocBegin']);
  for (var i = 0; i < 6; i++) bridge.comCallWith(h, 'Run', ['MoveRight']);
  for (var i = 0; i < 3; i++) bridge.comCallWith(h, 'Run', ['MoveSelRight']); // GHI 선택
  applyCharShape('Italic', 1);
  bridge.comCallWith(h, 'Run', ['Cancel']);
  console.log('✅ GHI에 이탤릭 적용');
});

// ================================================================
// 5-4: TextColor (빨간색)
// ================================================================
run('5-4 TextColor=RGBColor(255,0,0)', function() {
  // 두번째 문단 KLM 선택
  bridge.comCallWith(h, 'SetPos', [0, 1, 0]);
  for (var i = 0; i < 3; i++) bridge.comCallWith(h, 'Run', ['MoveSelRight']);
  var color = bridge.comCallWith(h, 'RGBColor', [255, 0, 0]);
  console.log('RGBColor(255,0,0)=' + color);
  applyCharShape('TextColor', color);
  bridge.comCallWith(h, 'Run', ['Cancel']);
  console.log('  KLM에 빨간색 적용');
});

// ================================================================
// 5-5: Underline
// ================================================================
run('5-5 Underline=1', function() {
  bridge.comCallWith(h, 'SetPos', [0, 1, 3]);
  for (var i = 0; i < 3; i++) bridge.comCallWith(h, 'Run', ['MoveSelRight']); // NOP 선택
  applyCharShape('Underline', 1);
  bridge.comCallWith(h, 'Run', ['Cancel']);
  console.log('✅ NOP에 밑줄 적용');
});

// ================================================================
// 5-6: StrikeOut
// ================================================================
run('5-6 StrikeOut=1', function() {
  bridge.comCallWith(h, 'SetPos', [0, 1, 6]);
  for (var i = 0; i < 3; i++) bridge.comCallWith(h, 'Run', ['MoveSelRight']); // QRS 선택
  applyCharShape('StrikeOut', 1);
  bridge.comCallWith(h, 'Run', ['Cancel']);
  console.log('✅ QRS에 취소선 적용');
});

// ================================================================
// 5-7: 서식 설정 후 텍스트 입력 (빈 커서에서)
// ================================================================
run('5-7 서식 먼저 → InsertText', function() {
  bridge.comCallWith(h, 'Run', ['MoveDocEnd']);
  insert('\r\n');
  // 서식 설정
  hwp.HAction.GetDefault('CharShape', hwp.HParameterSet.HCharShape.HSet);
  hwp.HParameterSet.HCharShape.Height = 3200; // 32pt
  hwp.HParameterSet.HCharShape.Bold = 1;
  hwp.HParameterSet.HCharShape.TextColor = bridge.comCallWith(h, 'RGBColor', [0, 0, 255]); // 파란색
  hwp.HAction.Execute('CharShape', hwp.HParameterSet.HCharShape.HSet);
  // 텍스트 입력
  insert('STYLED TEXT');
  console.log('✅ 32pt 볼드 파란색으로 입력됨?');
});

// ================================================================
// 5-8: 서식 읽기 — HParameterSet.HCharShape 값 확인
// ================================================================
run('5-8 서식 읽기', function() {
  // ABC (24pt) 위치로 이동
  bridge.comCallWith(h, 'Run', ['MoveDocBegin']);
  bridge.comCallWith(h, 'Run', ['MoveRight']); // A 뒤

  // 현재 서식 읽기
  hwp.HAction.GetDefault('CharShape', hwp.HParameterSet.HCharShape.HSet);
  var height = hwp.HParameterSet.HCharShape.Height;
  var bold = hwp.HParameterSet.HCharShape.Bold;
  var italic = hwp.HParameterSet.HCharShape.Italic;
  var color = hwp.HParameterSet.HCharShape.TextColor;
  console.log('A 위치: Height=' + height + ', Bold=' + bold + ', Italic=' + italic + ', Color=' + color);

  // DEF (볼드) 위치
  bridge.comCallWith(h, 'Run', ['MoveDocBegin']);
  for (var i = 0; i < 4; i++) bridge.comCallWith(h, 'Run', ['MoveRight']); // D 뒤
  hwp.HAction.GetDefault('CharShape', hwp.HParameterSet.HCharShape.HSet);
  height = hwp.HParameterSet.HCharShape.Height;
  bold = hwp.HParameterSet.HCharShape.Bold;
  console.log('         D 위치: Height=' + height + ', Bold=' + bold);
});

// ================================================================
// 5-9: CreateAction 패턴으로 CharShape
// ================================================================
run('5-9 CreateAction("CharShape")', function() {
  bridge.comCallWith(h, 'Run', ['MoveDocBegin']);
  bridge.comCallWith(h, 'Run', ['MoveSelLineEnd']); // 첫줄 전체 선택

  var act = hwp.CreateAction('CharShape');
  var set = act.CreateSet();
  act.GetDefault(set);
  set.SetItem('Height', 1600); // 16pt
  act.Execute(set);
  bridge.comCallWith(h, 'Run', ['Cancel']);
  console.log('✅ 첫줄 전체 16pt로 변경');
});

// ================================================================
// 5-10: SelectText + CharShape (pos 인코딩)
// ================================================================
run('5-10 SelectText + CharShape', function() {
  // 첫 문단의 오프셋 감지
  bridge.comCallWith(h, 'SetPos', [0, 0, 0]);
  bridge.comCallWith(h, 'Run', ['MoveParaBegin']);
  var ps = bridge.comCallWith(h, 'GetPosBySet', []);
  var offset = bridge.comCallWith(ps, 'Item', ['Pos']);
  console.log('offset=' + offset);

  // JK를 SelectText로 선택 (J=첫문단9번째, 근데 두번째 문단이니... 첫문단 마지막)
  bridge.comCallWith(h, 'SelectText', [0, offset + 7, 0, offset + 10]); // HIJ
  applyCharShape('TextColor', bridge.comCallWith(h, 'RGBColor', [0, 128, 0])); // 초록색
  bridge.comCallWith(h, 'Run', ['Cancel']);
  console.log('         HIJ에 초록색 적용');
});

// 종료
console.log('\n완료');
bridge.comCallWith(h, 'Clear', [1]);
bridge.comCallWith(h, 'Quit', []);
