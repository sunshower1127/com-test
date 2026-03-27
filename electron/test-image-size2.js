/**
 * InsertPicture 크기 극단적으로 다르게
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

// 아주 작게 (2cm x 2.5cm 정도 = ~5670 x ~7088 HWP단위)
// 1mm ≈ 283.46 HWP unit
// 20mm = 5669, 25mm = 7087

insert('opt=0, W=2000 H=2500:\r\n');
bridge.comCallWith(h, 'InsertPicture', [imgPath, 1, 0, 0, 0, 0, 2000, 2500]);

insert('\r\n\r\nopt=1, W=2000 H=2500:\r\n');
bridge.comCallWith(h, 'InsertPicture', [imgPath, 1, 1, 0, 0, 0, 2000, 2500]);

insert('\r\n\r\nopt=2, W=2000 H=2500:\r\n');
bridge.comCallWith(h, 'InsertPicture', [imgPath, 1, 2, 0, 0, 0, 2000, 2500]);

insert('\r\n\r\n원본 (비교용):\r\n');
bridge.comCallWith(h, 'InsertPicture', [imgPath, 1, 0, 0, 0, 0, 0, 0]);

console.log('확인! 작은 이미지가 보이면 그 sizeoption이 맞는거');
