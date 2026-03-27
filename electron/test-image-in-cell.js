/**
 * 표 셀 안에 이미지 넣기 — 단계별 확인
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

function readPos() {
  var ps = bridge.comCallWith(h, 'GetPosBySet', []);
  return {
    list: bridge.comCallWith(ps, 'Item', ['List']),
    para: bridge.comCallWith(ps, 'Item', ['Para']),
    pos: bridge.comCallWith(ps, 'Item', ['Pos']),
  };
}

var imgPath = 'C:/Users/seng1/Documents/Projects/com-test/data/예시이미지.jpg';

(async function() {

  // ── 단계 1: 간단한 표에서 먼저 테스트 ──
  await ask('── 1. 간단한 2x2 표 생성');
  hwp.HAction.GetDefault('TableCreate', hwp.HParameterSet.HTableCreation.HSet);
  hwp.HParameterSet.HTableCreation.Rows = 2;
  hwp.HParameterSet.HTableCreation.Cols = 2;
  hwp.HParameterSet.HTableCreation.WidthType = 2;
  hwp.HAction.Execute('TableCreate', hwp.HParameterSet.HTableCreation.HSet);

  insert('A1');
  bridge.comCallWith(h, 'Run', ['TableRightCell']);
  insert('B1');
  bridge.comCallWith(h, 'Run', ['TableRightCell']);
  insert('A2');
  bridge.comCallWith(h, 'Run', ['TableRightCell']);
  insert('B2');

  // A1으로
  for (var i = 0; i < 10; i++) bridge.comCallWith(h, 'Run', ['TableUpperCell']);
  for (var i = 0; i < 10; i++) bridge.comCallWith(h, 'Run', ['TableLeftCell']);
  var p = readPos();
  console.log('  A1 위치: L=' + p.list + ', P=' + p.para);
  console.log('  → A1 셀 안에 커서가 있는지 확인');

  // ── 단계 2: A1 셀에 이미지 삽입 ──
  await ask('── 2. A1에 이미지 삽입 (원본 크기)');
  bridge.comCallWith(h, 'Run', ['MoveLineBegin']);
  bridge.comCallWith(h, 'Run', ['MoveSelLineEnd']);
  bridge.comCallWith(h, 'Run', ['Delete']); // "A1" 삭제

  var ctrl = bridge.comCallWith(h, 'InsertPicture', [imgPath, 1, 0, 0, 0, 0, 0, 0]);
  var props = bridge.comGet(ctrl, 'Properties');
  var w = bridge.comCallWith(props, 'Item', ['Width']);
  var ht = bridge.comCallWith(props, 'Item', ['Height']);
  console.log('  삽입됨: W=' + w + ', H=' + ht);
  console.log('  → 이미지가 A1 셀 안에 있는지, 넘치는지 확인');

  // ── 단계 3: 크기 조절 ──
  await ask('── 3. 이미지 크기를 셀에 맞게 줄이기');
  bridge.comCallWith(props, 'SetItem', ['Width', 8000]);
  bridge.comCallWith(props, 'SetItem', ['Height', 10000]);
  bridge.comPut(ctrl, 'Properties', props);

  var p2 = bridge.comGet(ctrl, 'Properties');
  console.log('  조절 후: W=' + bridge.comCallWith(p2, 'Item', ['Width']) +
    ', H=' + bridge.comCallWith(p2, 'Item', ['Height']));
  console.log('  → 이미지가 작아졌는지 확인');

  // ── 단계 4: 이력서에서 테스트 ──
  await ask('── 4. 이력서 열기');
  bridge.comCallWith(h, 'Clear', [1]);
  var filePath = 'C:/Users/seng1/Documents/Projects/com-test/data/이력서_자기소개서_기본.hwp';
  bridge.comCallWith(h, 'Open', [filePath, 'HWP', '']);
  console.log('  이력서 열림');

  // ── 단계 5: MoveDown으로 표 진입 ──
  await ask('── 5. MoveDown → 위치 확인');
  bridge.comCallWith(h, 'Run', ['MoveDocBegin']);
  bridge.comCallWith(h, 'Run', ['MoveDown']);
  var p3 = readPos();
  console.log('  MoveDown 후: L=' + p3.list + ', P=' + p3.para + ', pos=' + p3.pos);

  // 좌상단으로
  for (var i = 0; i < 30; i++) bridge.comCallWith(h, 'Run', ['TableUpperCell']);
  for (var i = 0; i < 30; i++) bridge.comCallWith(h, 'Run', ['TableLeftCell']);
  var p4 = readPos();
  console.log('  좌상단: L=' + p4.list + ', P=' + p4.para + ', pos=' + p4.pos);

  // 셀 텍스트 확인
  bridge.comCallWith(h, 'Run', ['MoveSelLineEnd']);
  var cellText = String(bridge.comCallWith(h, 'GetTextFile', ['UNICODE', 'saveblock']));
  bridge.comCallWith(h, 'Run', ['Cancel']);
  console.log('  셀 텍스트: "' + cellText + '"');
  console.log('  → 사진 셀(좌상단)에 있는지 확인');

  // ── 단계 6: 사진 셀에 이미지 삽입 ──
  await ask('── 6. 사진 셀에 이미지 삽입');
  var ctrl2 = bridge.comCallWith(h, 'InsertPicture', [imgPath, 1, 0, 0, 0, 0, 0, 0]);
  if (ctrl2 && typeof ctrl2 === 'object') {
    var props2 = bridge.comGet(ctrl2, 'Properties');
    var w2 = bridge.comCallWith(props2, 'Item', ['Width']);
    var h2 = bridge.comCallWith(props2, 'Item', ['Height']);
    console.log('  삽입됨: W=' + w2 + ', H=' + h2);

    // 크기 조절 (사진 셀: 약 8041 x 10440)
    bridge.comCallWith(props2, 'SetItem', ['Width', 8041]);
    bridge.comCallWith(props2, 'SetItem', ['Height', 10440]);
    bridge.comPut(ctrl2, 'Properties', props2);
    console.log('  크기 조절 완료: 8041 x 10440');
  } else {
    console.log('  ❌ 삽입 실패');
  }
  console.log('  → 이미지가 사진 셀 안에 제대로 들어갔는지 확인');

  await ask('── 완료! 종료');
  bridge.comCallWith(h, 'Clear', [1]);
  bridge.comCallWith(h, 'Quit', []);
  rl.close();
  process.exit(0);

})();
