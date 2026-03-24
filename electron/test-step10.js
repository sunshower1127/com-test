/**
 * 단계 10: 페이지 설정 — 대화형
 */

var bridge = require('./native/com_bridge_node.node');
var { initBridge, createComProxy } = require('./dist/worker/proxy');
var readline = require('readline');

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

var rl = readline.createInterface({ input: process.stdin, output: process.stdout });
function ask(msg) {
  return new Promise(function(resolve) {
    rl.question('\n' + msg + ' → Enter ', function() { resolve(); });
  });
}

(async function() {

  insert('Page Setup Test - 이 텍스트로 변화를 확인하세요');

  // ── 10-1: 여백 넓게 ──
  await ask('── 10-1: 여백을 아주 넓게 (30mm)');
  hwp.HAction.GetDefault('PageSetup', hwp.HParameterSet.HSecDef.HSet);
  var pd = hwp.HParameterSet.HSecDef.PageDef;
  pd.LeftMargin = 8504;   // 기존값 유지 (기본 A4)
  pd.RightMargin = 8504;
  pd.TopMargin = 8504;    // 30mm로 변경
  pd.BottomMargin = 8504;
  pd.PaperWidth = 59528;
  pd.PaperHeight = 84186;
  pd.Landscape = 0;
  hwp.HAction.Execute('PageSetup', hwp.HParameterSet.HSecDef.HSet);
  console.log('  여백 30mm 적용 — 상하 여백이 넓어졌나요?');

  // ── 10-2: 가로 방향 ──
  await ask('── 10-2: 용지 방향 → 가로');
  hwp.HAction.GetDefault('PageSetup', hwp.HParameterSet.HSecDef.HSet);
  pd = hwp.HParameterSet.HSecDef.PageDef;
  pd.LeftMargin = 8504;
  pd.RightMargin = 8504;
  pd.TopMargin = 8504;
  pd.BottomMargin = 8504;
  pd.PaperWidth = 59528;
  pd.PaperHeight = 84186;
  pd.Landscape = 1;
  hwp.HAction.Execute('PageSetup', hwp.HParameterSet.HSecDef.HSet);
  console.log('  가로 방향 적용 — 용지가 가로로 바뀌었나요?');

  // ── 10-3: 세로 복구 ──
  await ask('── 10-3: 용지 방향 → 세로 복구');
  hwp.HAction.GetDefault('PageSetup', hwp.HParameterSet.HSecDef.HSet);
  pd = hwp.HParameterSet.HSecDef.PageDef;
  pd.LeftMargin = 8504;
  pd.RightMargin = 8504;
  pd.TopMargin = 8504;
  pd.BottomMargin = 8504;
  pd.PaperWidth = 59528;
  pd.PaperHeight = 84186;
  pd.Landscape = 0;
  hwp.HAction.Execute('PageSetup', hwp.HParameterSet.HSecDef.HSet);
  console.log('  세로 복구됨');

  // ── 10-4: 용지 크기 B5 ──
  await ask('── 10-4: 용지 크기 → B5 (182x257mm)');
  hwp.HAction.GetDefault('PageSetup', hwp.HParameterSet.HSecDef.HSet);
  pd = hwp.HParameterSet.HSecDef.PageDef;
  pd.PaperWidth = 51590;
  pd.PaperHeight = 72852;
  pd.LeftMargin = 8504;
  pd.RightMargin = 8504;
  pd.TopMargin = 5668;
  pd.BottomMargin = 4252;
  pd.Landscape = 0;
  hwp.HAction.Execute('PageSetup', hwp.HParameterSet.HSecDef.HSet);
  console.log('  B5 적용 — 용지가 작아졌나요?');

  // ── 10-5: A4 복구 ──
  await ask('── 10-5: A4 복구');
  hwp.HAction.GetDefault('PageSetup', hwp.HParameterSet.HSecDef.HSet);
  pd = hwp.HParameterSet.HSecDef.PageDef;
  pd.PaperWidth = 59528;
  pd.PaperHeight = 84186;
  pd.LeftMargin = 8504;
  pd.RightMargin = 8504;
  pd.TopMargin = 5668;
  pd.BottomMargin = 4252;
  pd.Landscape = 0;
  hwp.HAction.Execute('PageSetup', hwp.HParameterSet.HSecDef.HSet);
  console.log('  A4 복구됨');

  // ── 10-6: 여백 극단적으로 좁게 ──
  await ask('── 10-6: 여백 극소 (5mm)');
  hwp.HAction.GetDefault('PageSetup', hwp.HParameterSet.HSecDef.HSet);
  pd = hwp.HParameterSet.HSecDef.PageDef;
  pd.PaperWidth = 59528;
  pd.PaperHeight = 84186;
  pd.LeftMargin = 1417;
  pd.RightMargin = 1417;
  pd.TopMargin = 1417;
  pd.BottomMargin = 1417;
  pd.Landscape = 0;
  hwp.HAction.Execute('PageSetup', hwp.HParameterSet.HSecDef.HSet);
  console.log('  여백 5mm — 텍스트가 페이지 가장자리에 가까워졌나요?');

  // ── 종료 ──
  await ask('── 완료! 종료');
  bridge.comCallWith(h, 'Clear', [1]);
  bridge.comCallWith(h, 'Quit', []);
  rl.close();
  process.exit(0);

})();
