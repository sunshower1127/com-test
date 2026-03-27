/**
 * 플레이스홀더 → AllReplace 후 커서 위치 확인
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

(async function() {

  // HWPML2X 라운드트립으로 이력서 + 플레이스홀더
  var filePath = 'C:/Users/seng1/Documents/Projects/com-test/data/이력서_자기소개서_기본.hwp';
  bridge.comCallWith(h, 'Open', [filePath, 'HWP', '']);

  var xml = String(bridge.comCallWith(h, 'GetTextFile', ['HWPML2X', '']));

  // 사진 셀에 플레이스홀더 넣기 (RowAddr=1, ColAddr=0)
  var tableStart = xml.indexOf('<TABLE');
  var tableEnd = xml.indexOf('</TABLE>', tableStart) + 8;
  var tableXml = xml.substring(tableStart, tableEnd);

  var searchStart = 0;
  while (true) {
    var cellStart = tableXml.indexOf('<CELL', searchStart);
    if (cellStart < 0) break;
    var cellTagEnd = tableXml.indexOf('>', cellStart);
    var cellTag = tableXml.substring(cellStart, cellTagEnd + 1);

    if (cellTag.indexOf('RowAddr="1"') >= 0 && cellTag.indexOf('ColAddr="0"') >= 0) {
      var cellEnd = tableXml.indexOf('</CELL>', cellStart) + 7;
      var cellContent = tableXml.substring(cellStart, cellEnd);
      var modified = cellContent.replace(
        /<TEXT CharShape="(\d+)"\/>/,
        '<TEXT CharShape="$1"><CHAR>{{PHOTO}}</CHAR></TEXT>'
      );
      xml = xml.substring(0, tableStart + cellStart) + modified + xml.substring(tableStart + cellEnd);
      console.log('플레이스홀더 삽입 완료');
      break;
    }
    searchStart = cellStart + 1;
  }

  // 라운드트립
  bridge.comCallWith(h, 'Run', ['SelectAll']);
  bridge.comCallWith(h, 'Run', ['Delete']);
  bridge.comCallWith(h, 'SetTextFile', [xml, 'HWPML2X', '']);

  await ask('── 1. 이력서에 {{PHOTO}} 보이는지 확인');

  // Find로 플레이스홀더 위치 찾기
  console.log('\n=== Find로 {{PHOTO}} 찾기 ===');
  bridge.comCallWith(h, 'Run', ['MoveDocBegin']);

  // HAction Find (Replace 아닌 찾기만)
  hwp.HAction.GetDefault('RepeatFind', hwp.HParameterSet.HFindReplace.HSet);
  hwp.HParameterSet.HFindReplace.FindString = '{{PHOTO}}';
  hwp.HParameterSet.HFindReplace.IgnoreMessage = 1;
  var found = hwp.HAction.Execute('RepeatFind', hwp.HParameterSet.HFindReplace.HSet);
  console.log('Find ret=' + found);

  var p = readPos();
  console.log('커서 위치: L=' + p.list + ', P=' + p.para + ', pos=' + p.pos);
  console.log('L=' + p.list + ' → 0이면 본문, 다른 값이면 표 셀 안');

  await ask('── 2. 커서가 표 셀 안에 있는지 확인 (커서 깜빡임 위치)');

  // 찾은 텍스트 삭제
  bridge.comCallWith(h, 'Run', ['Delete']);
  console.log('{{PHOTO}} 삭제 완료');

  var p2 = readPos();
  console.log('삭제 후 위치: L=' + p2.list + ', P=' + p2.para);

  await ask('── 3. 삭제 후 커서 위치 확인');

  // 이미지 삽입
  var imgPath = 'C:/Users/seng1/Documents/Projects/com-test/data/예시이미지.jpg';
  var ctrl = bridge.comCallWith(h, 'InsertPicture', [imgPath, 1, 0, 0, 0, 0, 0, 0]);
  console.log('InsertPicture: ' + (ctrl ? '✅' : '❌'));

  if (ctrl && typeof ctrl === 'object') {
    var props = bridge.comGet(ctrl, 'Properties');
    var w = bridge.comCallWith(props, 'Item', ['Width']);
    var ht = bridge.comCallWith(props, 'Item', ['Height']);
    console.log('원본 크기: W=' + w + ', H=' + ht);

    // 사진 셀 크기에 맞게 조절 (8041 x 10440)
    bridge.comCallWith(props, 'SetItem', ['Width', 8041]);
    bridge.comCallWith(props, 'SetItem', ['Height', 10440]);
    bridge.comPut(ctrl, 'Properties', props);
    console.log('크기 조절: 8041 x 10440');
  }

  await ask('── 4. 사진이 셀 안에 제대로 들어갔는지 확인');

  await ask('── 종료');
  bridge.comCallWith(h, 'Clear', [1]);
  bridge.comCallWith(h, 'Quit', []);
  rl.close();
  process.exit(0);

})();
