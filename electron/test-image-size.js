/**
 * InsertPicture sizeoption + Width/Height 테스트
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

var imgPath = 'C:/Users/seng1/Documents/Projects/com-test/data/예시이미지.jpg';

// 여러 sizeoption으로 테스트
insert('sizeoption=0 (원본):\r\n');
bridge.comCallWith(h, 'InsertPicture', [imgPath, 1, 0, 0, 0, 0, 0, 0]);

insert('\r\n\r\nsizeoption=1, W=8000 H=10000:\r\n');
bridge.comCallWith(h, 'InsertPicture', [imgPath, 1, 1, 0, 0, 0, 8000, 10000]);

insert('\r\n\r\nsizeoption=2, W=8000 H=10000:\r\n');
bridge.comCallWith(h, 'InsertPicture', [imgPath, 1, 2, 0, 0, 0, 8000, 10000]);

insert('\r\n\r\nsizeoption=3, W=8000 H=10000:\r\n');
bridge.comCallWith(h, 'InsertPicture', [imgPath, 1, 3, 0, 0, 0, 8000, 10000]);

console.log('확인해주세요! 어떤 sizeoption이 크기를 제한하는지');
