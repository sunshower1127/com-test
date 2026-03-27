var fs = require('fs');
var buf = fs.readFileSync('C:/Users/seng1/Documents/Projects/com-test/data/이력서_hwpml2x.xml');
var xml = buf.toString('utf16le');

var headEnd = xml.indexOf('</HEAD>') + 7;
var bodyStart = xml.indexOf('<BODY>');
var bodyEnd = xml.indexOf('</BODY>') + 7;

var head = xml.substring(0, headEnd);
var body = xml.substring(bodyStart, bodyEnd);

// 보기 좋게 정돈
function prettify(s) {
  return s
    .replace(/></g, '>\n<')
    .replace(/\n\n+/g, '\n');
}

var dataDir = 'C:/Users/seng1/Documents/Projects/com-test/data/';
fs.writeFileSync(dataDir + '이력서_head.xml', prettify(head), 'utf8');
fs.writeFileSync(dataDir + '이력서_body.xml', prettify(body), 'utf8');
fs.writeFileSync(dataDir + '이력서_full.xml', prettify(xml), 'utf8');

console.log('저장 완료 (UTF-8):');
console.log('  이력서_head.xml: ' + fs.statSync(dataDir + '이력서_head.xml').size + ' bytes');
console.log('  이력서_body.xml: ' + fs.statSync(dataDir + '이력서_body.xml').size + ' bytes');
console.log('  이력서_full.xml: ' + fs.statSync(dataDir + '이력서_full.xml').size + ' bytes');
