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
var ask = (q) => new Promise(r => rl.question(q + ' → Enter\n', r));

(async () => {
  var docPath = 'C:/Users/seng1/Documents/Projects/com-test/data/이력서_자기소개서_기본.hwp';
  hwp.Open(docPath, 'HWP', '');
  console.log('문서 열림');

  await ask('── MoveDown 루프로 표 진입 테스트');

  hwp.Run('MoveDocBegin');

  var maxTry = 20;
  var entered = false;
  for (var i = 0; i < maxTry; i++) {
    hwp.Run('MoveDown');

    var ps = hwp.GetPosBySet();
    var list = ps.Item('List');
    var para = ps.Item('Para');
    var pos = ps.Item('Pos');

    console.log('  MoveDown[' + i + ']: L=' + list + ' P=' + para + ' pos=' + pos);

    if (list > 0) {
      console.log('  ✅ 표 진입 성공! (MoveDown ' + (i+1) + '회)');
      entered = true;

      // 현재 셀 텍스트 확인
      hwp.Run('MoveLineBegin');
      hwp.Run('MoveSelLineEnd');
      var text = hwp.GetTextFile('UNICODE', 'saveblock');
      hwp.Run('Cancel');
      console.log('  현재 셀: "' + text + '"');
      break;
    }
  }

  if (!entered) {
    console.log('  ❌ ' + maxTry + '회 시도했지만 표 진입 실패');
  }

  await ask('── 종료');
  hwp.Clear(1);
  hwp.Quit();
  rl.close();
})();
