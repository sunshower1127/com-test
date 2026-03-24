/**
 * 단계 8: 그림 삽입 — 자동 테스트
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

// 테스트 이미지 생성 (간단한 PNG)
var testImgDir = path.join(require('os').tmpdir(), 'hwp-test');
if (!fs.existsSync(testImgDir)) fs.mkdirSync(testImgDir, { recursive: true });
var testImgPath = path.join(testImgDir, 'test.png');

// 최소 PNG 생성 (1x1 빨간 픽셀)
if (!fs.existsSync(testImgPath)) {
  var pngBuf = Buffer.from([
    0x89,0x50,0x4E,0x47,0x0D,0x0A,0x1A,0x0A, // PNG signature
    0x00,0x00,0x00,0x0D,0x49,0x48,0x44,0x52, // IHDR chunk
    0x00,0x00,0x00,0x64,0x00,0x00,0x00,0x64, // 100x100
    0x08,0x02,0x00,0x00,0x00,0xFF,0x80,0x02, // 8bit RGB
    0x03,0x00,0x00,0x00,0x01,0x73,0x52,0x47, // sRGB
    0x42,0x00,0xAE,0xCE,0x1C,0xE9
  ]);
  // Actually let's just use a BMP which is simpler
  // Or find an existing image
  console.log('테스트 이미지 경로: ' + testImgPath);
}

// 시스템에 있는 이미지 찾기
var imgCandidates = [
  'C:/Windows/Web/Wallpaper/Windows/img0.jpg',
  'C:/Windows/Web/Screen/img100.jpg',
  'C:/Windows/Web/Wallpaper/Theme1/img1.jpg',
  'C:/Users/seng1/Documents/Projects/com-test/data/testdoc_hwpx/Preview/PrvImage.png',
];

var testImg = null;
for (var i = 0; i < imgCandidates.length; i++) {
  if (fs.existsSync(imgCandidates[i])) {
    testImg = imgCandidates[i].replace(/\\/g, '/');
    break;
  }
}

if (!testImg) {
  // BMP 직접 생성 (10x10 빨간색)
  var w = 10, ht = 10;
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
      bmp[offset] = 0;     // B
      bmp[offset+1] = 0;   // G
      bmp[offset+2] = 255; // R
    }
  }
  testImg = path.join(testImgDir, 'test.bmp').replace(/\\/g, '/');
  fs.writeFileSync(testImg, bmp);
  console.log('BMP 생성: ' + testImg);
}

console.log('테스트 이미지: ' + testImg + '\n');

insert('그림 삽입 테스트\r\n');

// ================================================================
// 8-1: InsertPicture
// ================================================================
run('8-1 InsertPicture', function() {
  var ret = bridge.comCallWith(h, 'InsertPicture', [testImg, true, 0, false, false, 0, 0, 0]);
  console.log('ret=' + ret + ' (true면 성공)');
});

// ================================================================
// 8-2: HParameterSet 패턴으로 InsertPicture (IsActionEnable 먼저 확인)
// ================================================================
run('8-2 IsActionEnable("InsertPicture")', function() {
  var en = bridge.comCallWith(h, 'IsActionEnable', ['InsertPicture']);
  console.log(String(en));
});

// ================================================================
// 8-3: CreatePageImage — 페이지를 이미지로 저장
// ================================================================
run('8-3 CreatePageImage', function() {
  var outPath = path.join(testImgDir, 'page0.png').replace(/\\/g, '/');
  var ret = bridge.comCallWith(h, 'CreatePageImage', [outPath, 0, 300, 24, 'PNG']);
  console.log('ret=' + ret);
  if (fs.existsSync(outPath)) {
    console.log('  ✅ 파일 생성됨: ' + fs.statSync(outPath).size + ' bytes');
  } else {
    console.log('  ❌ 파일 없음 (보안모듈 필요?)');
  }
});

// ================================================================
// 8-4: CreatePageImage BMP
// ================================================================
run('8-4 CreatePageImage BMP', function() {
  var outPath = path.join(testImgDir, 'page0.bmp').replace(/\\/g, '/');
  var ret = bridge.comCallWith(h, 'CreatePageImage', [outPath, 0, 150, 24, 'BMP']);
  console.log('ret=' + ret);
  if (fs.existsSync(outPath)) {
    console.log('  ✅ 파일 생성됨: ' + fs.statSync(outPath).size + ' bytes');
  } else {
    console.log('  ❌ 파일 없음');
  }
});

// ================================================================
// 8-5: 그림 삽입 후 GetTextFile — 그림이 텍스트에 어떻게 나오는지
// ================================================================
run('8-5 그림 삽입 후 GetTextFile', function() {
  var text = String(bridge.comCallWith(h, 'GetTextFile', ['UNICODE', '']));
  console.log('"' + text.substring(0, 50) + '..."');
});

console.log('\n확인해주세요!');
