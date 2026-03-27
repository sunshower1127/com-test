var bridge = require('./native/com_bridge_node.node');
const readline = require('readline');
const rl = readline.createInterface({ input: process.stdin, output: process.stdout });
const ask = (q) => new Promise(r => rl.question(q + ' → Enter\n', r));

(async () => {
  const h = bridge.comCreate('HWPFrame.HwpObject');
  // RegisterModule 불필요
  const ver = bridge.comGet(h, 'Version');
  console.log('HWP Version:', ver);

  // 이력서 열기
  const docPath = 'C:/Users/seng1/Documents/Projects/com-test/data/이력서_자기소개서_기본.hwp';
  bridge.comCallWith(h, 'Open', [docPath, 'HWP', '']);
  console.log('문서 열림');

  // PageCount 확인
  const pageCount = bridge.comGet(h, 'PageCount');
  console.log('PageCount:', pageCount);

  await ask('── 페이지별 텍스트 추출');

  // 포맷별 테스트 (페이지 0만)
  var formats = ['UNICODE', 'TEXT', 'HTML', 'HWPML2X', 'HWP'];
  for (var fmt of formats) {
    console.log('\n=== GetPageText(0, "' + fmt + '") ===');
    try {
      var text = bridge.comCallWith(h, 'GetPageText', [0, fmt]);
      console.log('type:', typeof text, 'length:', text ? String(text).length : 0);
      if (text) console.log('preview:', String(text).substring(0, 150));
    } catch (e) {
      console.log('ERROR:', e.message);
    }
  }

  console.log('\n정리 중...');
  bridge.comCallWith(h, 'Clear', [1]);
  bridge.comCallWith(h, 'Quit', []);
  rl.close();
})();
