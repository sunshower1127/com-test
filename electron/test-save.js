var bridge = require('./native/com_bridge_node.node');
var { initBridge, createComProxy } = require('./dist/worker/proxy');
bridge.comInit();
initBridge(bridge);

var h = bridge.comCreate('HWPFrame.HwpObject');
var hwp = createComProxy(h);
var win0 = bridge.comCallWith(bridge.comGet(h, 'XHwpWindows'), 'Item', [0]);
bridge.comPut(win0, 'Visible', false);
bridge.comPut(win0, 'Visible', true);

hwp.HAction.GetDefault('InsertText', hwp.HParameterSet.HInsertText.HSet);
hwp.HParameterSet.HInsertText.Text = '저장 테스트';
hwp.HAction.Execute('InsertText', hwp.HParameterSet.HInsertText.HSet);

var rl = require('readline').createInterface({ input: process.stdin, output: process.stdout });

rl.question('1차 SaveAs 시도합니다. Enter: ', function() {
  console.log('1차 시도...');
  try {
    hwp.SaveAs('C:/Users/seng1/Desktop/hwp-test.hwp', 'HWP', '');
    console.log('1차: ✅ 성공!');
  } catch(e) {
    console.log('1차: ❌ ' + e.message);
  }

  rl.question('팝업 허용 후 Enter: ', function() {
    console.log('2차 시도...');
    try {
      hwp.SaveAs('C:/Users/seng1/Desktop/hwp-test.hwp', 'HWP', '');
      console.log('2차: ✅ 성공!');
    } catch(e) {
      console.log('2차: ❌ ' + e.message);
    }

    rl.question('Enter로 종료: ', function() {
      bridge.comCallWith(h, 'Clear', [1]);
      bridge.comCallWith(h, 'Quit', []);
      process.exit(0);
    });
  });
});
