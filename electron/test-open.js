var bridge = require('./native/com_bridge_node.node');
var { initBridge, createComProxy } = require('./dist/worker/proxy');
bridge.comInit();
initBridge(bridge);

var h = bridge.comCreate('HWPFrame.HwpObject');
var hwp = createComProxy(h);
var win0 = bridge.comCallWith(bridge.comGet(h, 'XHwpWindows'), 'Item', [0]);
bridge.comPut(win0, 'Visible', false);
bridge.comPut(win0, 'Visible', true);

var paths = [
  ['슬래시 /', 'C:/Users/seng1/OneDrive/문서/Welcome to Hwp.hwp'],
  ['백슬래시 \\', 'C:\\Users\\seng1\\OneDrive\\문서\\Welcome to Hwp.hwp'],
];

console.log('팝업이 뜨면 매번 허용 눌러주세요!\n');

paths.forEach(function(p) {
  var label = p[0];
  var filePath = p[1];
  console.log('── ' + label + ' ──');
  console.log('  경로: ' + filePath);
  try {
    hwp.Clear(1);
    var ret = hwp.Open(filePath, 'HWP', '');
    var text = bridge.comCallWith(h, 'GetTextFile', ['UNICODE', '']);
    console.log('  ✅ 반환=' + String(ret) + ', 텍스트="' + String(text).substring(0, 20) + '..."');
  } catch(e) {
    console.log('  ❌ ' + e.message);
  }
});

bridge.comCallWith(h, 'Clear', [1]);
bridge.comCallWith(h, 'Quit', []);
console.log('\n완료');
