/**
 * 단계 8 디버그: InsertPicture 파라미터 + CreatePageImage 포맷별
 */

var bridge = require('./native/com_bridge_node.node');
var { initBridge, createComProxy } = require('./dist/worker/proxy');
var fs = require('fs');
var path = require('path');

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
function run(label, fn) {
  process.stdout.write('── ' + label + ': ');
  try { fn(); } catch(e) { console.log('❌ ' + e.message); }
}

var testImgDir = path.join(require('os').tmpdir(), 'hwp-test');
var testImg = 'C:/Windows/Web/Wallpaper/Windows/img0.jpg';

insert('그림 삽입 테스트\r\n\r\n');

// ================================================================
// InsertPicture — ITypeInfo 시그니처 확인
// ================================================================
console.log('=== InsertPicture 시그니처 ===');
var members = bridge.comListMembers(h);
var ip = members.filter(function(m) { return m.name === 'InsertPicture'; });
ip.forEach(function(m) { console.log(JSON.stringify(m)); });

// ================================================================
// D1: comCallWith로 다양한 파라미터 조합
// ================================================================
run('\nD1 InsertPicture (문자열만)', function() {
  var ret = bridge.comCallWith(h, 'InsertPicture', [testImg]);
  console.log('ret=' + ret);
});

run('D2 InsertPicture (path, embed)', function() {
  var ret = bridge.comCallWith(h, 'InsertPicture', [testImg, 1]);
  console.log('ret=' + ret);
});

run('D3 InsertPicture (path, embed, sizeoption)', function() {
  var ret = bridge.comCallWith(h, 'InsertPicture', [testImg, 1, 0]);
  console.log('ret=' + ret);
});

run('D4 InsertPicture 전체 파라미터 (숫자)', function() {
  var ret = bridge.comCallWith(h, 'InsertPicture', [testImg, 1, 0, 0, 0, 0, 0, 0]);
  console.log('ret=' + ret);
});

run('D5 InsertPicture (bool 대신 숫자)', function() {
  bridge.comCallWith(h, 'Run', ['BreakPara']);
  var ret = bridge.comCallWith(h, 'InsertPicture', [testImg, true, 0, false, false, 0, 0, 0]);
  console.log('ret=' + ret);
});

// ================================================================
// CreatePageImage 포맷별
// ================================================================
console.log('\n=== CreatePageImage 포맷별 ===');
var formats = ['BMP', 'PNG', 'JPG', 'JPEG', 'GIF', 'EMF', 'WMF', 'TIFF'];
formats.forEach(function(fmt) {
  var outPath = path.join(testImgDir, 'page0.' + fmt.toLowerCase()).replace(/\\/g, '/');
  try {
    if (fs.existsSync(outPath)) fs.unlinkSync(outPath);
    var ret = bridge.comCallWith(h, 'CreatePageImage', [outPath, 0, 150, 24, fmt]);
    var exists = fs.existsSync(outPath);
    var size = exists ? fs.statSync(outPath).size : 0;
    console.log('  ' + fmt + ': ret=' + ret + ', ' + (exists ? '✅ ' + size + 'B' : '❌ 없음'));
  } catch(e) {
    console.log('  ' + fmt + ': ❌ ' + e.message);
  }
});

console.log('\n확인해주세요!');
