var bridge = require('./native/com_bridge_node.node');
var { initBridge, createComProxy } = require('./dist/worker/proxy');
var readline = require('readline');
var fs = require('fs');
var rl = readline.createInterface({ input: process.stdin, output: process.stdout });
var ask = function(q) { return new Promise(function(r) { rl.question(q + ' → Enter\n', r); }); };

bridge.comInit();
initBridge(bridge);

var h = bridge.comCreate('HWPFrame.HwpObject');
var hwp = createComProxy(h);
hwp.RegisterModule('FilePathCheckDLL', '');
var wins = bridge.comGet(h, 'XHwpWindows');
var win0 = bridge.comCallWith(wins, 'Item', [0]);
bridge.comPut(win0, 'Visible', true);

console.log('HWP 열림. 자유롭게 타이핑하세요.');
console.log('엔터, Shift+엔터 등 다양하게 시도 후 여기서 Enter 누르면 XML 출력합니다.');
console.log('');

(async function() {
  await ask('── 타이핑 완료 후 XML 확인');

  var xml = String(bridge.comCallWith(h, 'GetTextFile', ['HWPML2X', '']));

  // BODY > SECTION 추출
  var sectionMatch = xml.match(/<SECTION[^>]*>([\s\S]*?)<\/SECTION>/);
  if (sectionMatch) {
    var section = sectionMatch[1];
    // 보기 좋게 정리
    var formatted = section
      .replace(/></g, '>\n<')
      .replace(/\n\n+/g, '\n');

    fs.writeFileSync('../data/paragraph_test.xml', formatted, 'utf-8');
    console.log('저장: data/paragraph_test.xml');
    console.log('');

    // P 태그 분석
    var pRe = /<P[\s\S]*?<\/P>/g;
    var m, pi = 0;
    while ((m = pRe.exec(section))) {
      var chars = [];
      var cRe = /<CHAR\b[^>]*?>([\s\S]*?)<\/CHAR>/g;
      var cm;
      while ((cm = cRe.exec(m[0]))) {
        chars.push(cm[1]);
      }
      var textCount = (m[0].match(/<TEXT/g) || []).length;
      var hasLineBreak = m[0].indexOf('LINEBREAK') >= 0 || m[0].indexOf('BreakLine') >= 0;
      console.log('P[' + pi + '] TEXT×' + textCount + (hasLineBreak ? ' [LINEBREAK]' : '') + ': ' + JSON.stringify(chars.join(' | ')));
      pi++;
    }
  }

  await ask('── 종료');
  hwp.Clear(1);
  hwp.Quit();
  rl.close();
})();
