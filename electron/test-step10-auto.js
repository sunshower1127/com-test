/**
 * 단계 10: 페이지 설정 — 자동 테스트
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

insert('Page Setup Test\r\nThis is a test document for page settings.');

// ── HPageDef 멤버 확인 ──
console.log('=== HParameterSet.HSecDef 멤버 ===');
try {
  var hps = bridge.comGet(h, 'HParameterSet');
  var hsd = bridge.comGet(hps, 'HSecDef');
  var members = bridge.comListMembers(hsd);
  var props = members.filter(function(m) {
    return (m.kind === 'put' || m.kind === 'get/put') && m.name !== 'HSet';
  });
  props.forEach(function(m) {
    try {
      var v = bridge.comGet(hsd, m.name);
      if (typeof v !== 'object') console.log('  ' + m.name + ' = ' + v);
      else console.log('  ' + m.name + ' = [object]');
    } catch(e) { console.log('  ' + m.name + ' (읽기 실패)'); }
  });
} catch(e) { console.log('HSecDef: ❌ ' + e.message); }

// ── 현재 페이지 설정 읽기 ──
console.log('\n=== 현재 페이지 설정 읽기 ===');
run('10-1 PageSetup 읽기', function() {
  hwp.HAction.GetDefault('PageSetup', hwp.HParameterSet.HSecDef.HSet);
  var hsd = hwp.HParameterSet.HSecDef;
  var pageDef = hsd.PageDef;

  var members = bridge.comListMembers(bridge.comGet(bridge.comGet(h, 'HParameterSet'), 'HSecDef'));
  var pageDefHandle = bridge.comGet(bridge.comGet(h, 'HParameterSet'), 'HSecDef');
  var pdHandle = bridge.comGet(pageDefHandle, 'PageDef');

  if (pdHandle && typeof pdHandle === 'object') {
    var pdMembers = bridge.comListMembers(pdHandle);
    var pdProps = pdMembers.filter(function(m) {
      return (m.kind === 'get' || m.kind === 'get/put') &&
        ['QueryInterface','AddRef','Release','GetTypeInfoCount','GetTypeInfo','GetIDsOfNames','Invoke'].indexOf(m.name) === -1;
    });
    console.log('PageDef 프로퍼티:');
    pdProps.forEach(function(m) {
      try {
        var v = bridge.comGet(pdHandle, m.name);
        if (typeof v !== 'object') console.log('  ' + m.name + ' = ' + v);
      } catch(e) {}
    });
  } else {
    console.log('PageDef 접근 실패');
  }
});

// ── 여백 변경 ──
run('\n10-2 여백 변경', function() {
  hwp.HAction.GetDefault('PageSetup', hwp.HParameterSet.HSecDef.HSet);
  var hsd = hwp.HParameterSet.HSecDef;
  var pd = hsd.PageDef;

  // 현재 여백 읽기
  console.log('변경 전:');
  console.log('  LeftMargin=' + pd.LeftMargin);
  console.log('  RightMargin=' + pd.RightMargin);
  console.log('  TopMargin=' + pd.TopMargin);
  console.log('  BottomMargin=' + pd.BottomMargin);

  // 여백 넓게
  pd.LeftMargin = 3000;
  pd.RightMargin = 3000;
  pd.TopMargin = 3000;
  pd.BottomMargin = 3000;

  var ret = hwp.HAction.Execute('PageSetup', hwp.HParameterSet.HSecDef.HSet);
  console.log('  Execute ret=' + ret);
});

// ── 용지 방향 ──
run('\n10-3 용지 방향 (가로)', function() {
  hwp.HAction.GetDefault('PageSetup', hwp.HParameterSet.HSecDef.HSet);
  var pd = hwp.HParameterSet.HSecDef.PageDef;

  console.log('현재 Landscape=' + pd.Landscape);
  pd.Landscape = 1; // 가로
  var ret = hwp.HAction.Execute('PageSetup', hwp.HParameterSet.HSecDef.HSet);
  console.log('  가로 설정 ret=' + ret);
});

// ── 다시 세로 ──
run('\n10-4 용지 방향 (세로 복구)', function() {
  hwp.HAction.GetDefault('PageSetup', hwp.HParameterSet.HSecDef.HSet);
  hwp.HParameterSet.HSecDef.PageDef.Landscape = 0;
  hwp.HAction.Execute('PageSetup', hwp.HParameterSet.HSecDef.HSet);
  console.log('✅ 세로 복구');
});

// ── 용지 크기 변경 ──
run('\n10-5 용지 크기 (B5)', function() {
  hwp.HAction.GetDefault('PageSetup', hwp.HParameterSet.HSecDef.HSet);
  var pd = hwp.HParameterSet.HSecDef.PageDef;
  console.log('현재 Width=' + pd.PaperWidth + ', Height=' + pd.PaperHeight);

  // B5: 182mm x 257mm → HWP 단위 (1mm ≈ 283.46 HWP unit)
  pd.PaperWidth = 51590;  // 182mm
  pd.PaperHeight = 72852; // 257mm
  var ret = hwp.HAction.Execute('PageSetup', hwp.HParameterSet.HSecDef.HSet);
  console.log('  B5 설정 ret=' + ret);
});

// ── A4로 복구 ──
run('\n10-6 용지 크기 (A4 복구)', function() {
  hwp.HAction.GetDefault('PageSetup', hwp.HParameterSet.HSecDef.HSet);
  var pd = hwp.HParameterSet.HSecDef.PageDef;
  pd.PaperWidth = 59528;   // 210mm
  pd.PaperHeight = 84186;  // 297mm
  pd.LeftMargin = 8504;
  pd.RightMargin = 8504;
  pd.TopMargin = 5668;
  pd.BottomMargin = 4252;
  hwp.HAction.Execute('PageSetup', hwp.HParameterSet.HSecDef.HSet);
  console.log('✅ A4 복구');
});

// ── CreateAction 패턴 ──
run('\n10-7 CreateAction("PageSetup")', function() {
  var act = hwp.CreateAction('PageSetup');
  var set = act.CreateSet();
  act.GetDefault(set);
  console.log('SetID=' + bridge.comGet(set, 'SetID'));
  // ItemExist 확인
  var names = ['PageWidth', 'PageHeight', 'Landscape', 'LeftMargin', 'TopMargin',
               'PaperWidth', 'PaperHeight'];
  names.forEach(function(n) {
    var exists = bridge.comCallWith(set, 'ItemExist', [n]);
    if (exists) {
      var v = bridge.comCallWith(set, 'Item', [n]);
      console.log('  ' + n + '=' + v);
    }
  });
});

console.log('\n확인해주세요!');
