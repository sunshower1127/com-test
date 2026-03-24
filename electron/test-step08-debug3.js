/**
 * 그림 삽입 추가 테스트: URL, 클립보드, Buffer
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
if (!fs.existsSync(testImgDir)) fs.mkdirSync(testImgDir, { recursive: true });

// ================================================================
// T1: 로컬 파일 (기준 — 이건 됨)
// ================================================================
run('T1 로컬 파일', function() {
  insert('T1: local file\r\n');
  var ctrl = bridge.comCallWith(h, 'InsertPicture', [
    'C:/Windows/Web/Wallpaper/Windows/img0.jpg', 1, 0, 0, 0, 0, 0, 0
  ]);
  console.log(typeof ctrl === 'object' ? '✅ 삽입됨' : '❌ ' + String(ctrl));
  insert('\r\n');
});

// ================================================================
// T2: HTTP URL
// ================================================================
run('T2 HTTP URL', function() {
  insert('T2: http url\r\n');
  var ctrl = bridge.comCallWith(h, 'InsertPicture', [
    'https://www.google.com/images/branding/googlelogo/2x/googlelogo_color_272x92dp.png',
    1, 0, 0, 0, 0, 0, 0
  ]);
  console.log(typeof ctrl === 'object' ? '✅ 삽입됨' : '⚠️ ' + String(ctrl));
  insert('\r\n');
});

// ================================================================
// T3: HTTPS URL (다른 이미지)
// ================================================================
run('T3 HTTPS URL', function() {
  insert('T3: https url\r\n');
  var ctrl = bridge.comCallWith(h, 'InsertPicture', [
    'https://upload.wikimedia.org/wikipedia/commons/thumb/4/47/PNG_transparency_demonstration_1.png/280px-PNG_transparency_demonstration_1.png',
    1, 0, 0, 0, 0, 0, 0
  ]);
  console.log(typeof ctrl === 'object' ? '✅ 삽입됨' : '⚠️ ' + String(ctrl));
  insert('\r\n');
});

// ================================================================
// T4: 존재하지 않는 로컬 파일
// ================================================================
run('T4 존재하지 않는 파일', function() {
  insert('T4: nonexistent\r\n');
  var ctrl = bridge.comCallWith(h, 'InsertPicture', [
    'C:/no_such_file.png', 1, 0, 0, 0, 0, 0, 0
  ]);
  console.log(typeof ctrl === 'object' ? '⚠️ 객체 반환?' : '❌ ' + String(ctrl));
  insert('\r\n');
});

// ================================================================
// T5: Paste — 클립보드에 이미지 넣고 붙여넣기
// ================================================================
run('T5 Paste (클립보드)', function() {
  insert('T5: paste\r\n');
  // 먼저 클립보드에 이미지 복사 — 로컬 파일을 읽어서 클립보드에 못 넣으니
  // Run("Paste") 자체가 되는지만 확인
  try {
    bridge.comCallWith(h, 'Run', ['Paste']);
    console.log('Paste 실행됨 (클립보드에 이미지 있어야 보임)');
  } catch(e) { console.log('❌ ' + e.message); }
  insert('\r\n');
});

// ================================================================
// T6: SetTextFile로 HTML img 태그
// ================================================================
run('T6 SetTextFile HTML img', function() {
  insert('T6: html img\r\n');
  var html = '<img src="C:/Windows/Web/Wallpaper/Windows/img0.jpg" width="200">';
  var ret = bridge.comCallWith(h, 'SetTextFile', [html, 'HTML', 'insertfile']);
  console.log('ret=' + ret);
  insert('\r\n');
});

// ================================================================
// T7: Base64 이미지를 SetTextFile HTML로
// ================================================================
run('T7 SetTextFile HTML base64 img', function() {
  insert('T7: base64 img\r\n');
  // 1x1 빨간 PNG base64
  var b64 = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8/5+hHgAHggJ/PchI7wAAAABJRU5ErkJggg==';
  var html = '<img src="data:image/png;base64,' + b64 + '" width="100" height="100">';
  var ret = bridge.comCallWith(h, 'SetTextFile', [html, 'HTML', 'insertfile']);
  console.log('ret=' + ret);
  insert('\r\n');
});

// ================================================================
// T8: 임시 파일에 buffer 쓰고 InsertPicture
// ================================================================
run('T8 Buffer → 임시파일 → InsertPicture', function() {
  insert('T8: buffer->tmpfile\r\n');
  // BMP 생성 (20x20 파란색)
  var w = 20, ht = 20;
  var rowSize = Math.ceil(w * 3 / 4) * 4;
  var dataSize = rowSize * ht;
  var fileSize = 54 + dataSize;
  var bmp = Buffer.alloc(fileSize);
  bmp.write('BM', 0);
  bmp.writeUInt32LE(fileSize, 2);
  bmp.writeUInt32LE(54, 10);
  bmp.writeUInt32LE(40, 14);
  bmp.writeInt32LE(w, 18);
  bmp.writeInt32LE(ht, 22);
  bmp.writeUInt16LE(1, 26);
  bmp.writeUInt16LE(24, 28);
  bmp.writeUInt32LE(dataSize, 34);
  for (var y = 0; y < ht; y++) {
    for (var x = 0; x < w; x++) {
      var offset = 54 + y * rowSize + x * 3;
      bmp[offset] = 255;   // B
      bmp[offset+1] = 0;   // G
      bmp[offset+2] = 0;   // R
    }
  }
  var tmpPath = path.join(testImgDir, 'blue20x20.bmp').replace(/\\/g, '/');
  fs.writeFileSync(tmpPath, bmp);

  var ctrl = bridge.comCallWith(h, 'InsertPicture', [tmpPath, 1, 0, 0, 0, 0, 0, 0]);
  console.log(typeof ctrl === 'object' ? '✅ 삽입됨' : '❌');
});

console.log('\n확인해주세요!');
