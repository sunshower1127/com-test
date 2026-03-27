var bridge = require('./native/com_bridge_node.node');
var { initBridge, createComProxy } = require('./dist/worker/proxy');

bridge.comInit();
initBridge(bridge);

var h = bridge.comCreate('HWPFrame.HwpObject');
var hwp = createComProxy(h);
var win0 = bridge.comCallWith(bridge.comGet(h, 'XHwpWindows'), 'Item', [0]);
bridge.comPut(win0, 'Visible', true);

// RGBColor 반환값 확인
var red = hwp.RGBColor(255, 0, 0);
var green = hwp.RGBColor(0, 255, 0);
var blue = hwp.RGBColor(0, 0, 255);

console.log('RGBColor(255,0,0) = ' + red + ' (hex: 0x' + red.toString(16).padStart(8,'0') + ')');
console.log('RGBColor(0,255,0) = ' + green + ' (hex: 0x' + green.toString(16).padStart(8,'0') + ')');
console.log('RGBColor(0,0,255) = ' + blue + ' (hex: 0x' + blue.toString(16).padStart(8,'0') + ')');

// 실제로 빨간 글자를 써보기
hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet);
hwp.HParameterSet.HInsertText.Text = "RED ";
hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet);

// 방금 쓴 텍스트 선택
hwp.Run("MoveLineBegin");
hwp.Run("MoveSelLineEnd");

// CharShape로 빨간색 적용
hwp.HAction.GetDefault("CharShape", hwp.HParameterSet.HCharShape.HSet);
hwp.HParameterSet.HCharShape.TextColor = hwp.RGBColor(255, 0, 0);
hwp.HAction.Execute("CharShape", hwp.HParameterSet.HCharShape.HSet);

hwp.Run("MoveDocEnd");

console.log('\n"RED" 텍스트가 빨간색이면 RGBColor가 올바르게 변환하는 것');
console.log('확인 후 Ctrl+C로 종료');

// 프로세스 유지
setInterval(() => {}, 60000);
