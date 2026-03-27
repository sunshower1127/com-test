var bridge = require('./native/com_bridge_node.node');
var { initBridge, createComProxy } = require('./dist/worker/proxy');
var fs = require('fs');

bridge.comInit();
initBridge(bridge);

var h = bridge.comCreate('HWPFrame.HwpObject');
var hwp = createComProxy(h);
var win0 = bridge.comCallWith(bridge.comGet(h, 'XHwpWindows'), 'Item', [0]);
bridge.comPut(win0, 'Visible', true);

var filePath = 'C:/Users/seng1/Documents/Projects/com-test/data/이력서_자기소개서_기본.hwp';
bridge.comCallWith(h, 'Open', [filePath, 'HWP', '']);

// HTML 추출
var html = String(bridge.comCallWith(h, 'GetTextFile', ['HTML', '']));

// meta charset 확인
var charsetMatch = html.match(/charset[=:]\s*([^"'\s;]+)/i);
console.log('charset: ' + (charsetMatch ? charsetMatch[1] : '없음'));
console.log('처음 300자:\n' + html.substring(0, 300));

// 한글이 깨지는지 확인
var hasKorean = /[가-힣]/.test(html);
console.log('\n한글 포함: ' + hasKorean);

// 파일로 저장 (그대로)
var dataDir = 'C:/Users/seng1/Documents/Projects/com-test/data/';
fs.writeFileSync(dataDir + '이력서.html', html, 'utf8');
console.log('저장: 이력서.html (' + fs.statSync(dataDir + '이력서.html').size + ' bytes)');

bridge.comCallWith(h, 'Clear', [1]);
bridge.comCallWith(h, 'Quit', []);
