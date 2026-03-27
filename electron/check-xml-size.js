var fs = require('fs');
var buf = fs.readFileSync('C:/Users/seng1/Documents/Projects/com-test/data/이력서_hwpml2x.xml');
var xml = buf.toString('utf16le');

var headEnd = xml.indexOf('</HEAD>') + 7;
var bodyStart = xml.indexOf('<BODY>');
var bodyEnd = xml.indexOf('</BODY>') + 7;

console.log('전체: ' + xml.length + '자');
console.log('HEAD: ' + headEnd + '자 (' + Math.round(headEnd/xml.length*100) + '%)');
console.log('BODY: ' + (bodyEnd - bodyStart) + '자 (' + Math.round((bodyEnd-bodyStart)/xml.length*100) + '%)');

// BODY에서 표 부분만
var tbl1Start = xml.indexOf('<TABLE', bodyStart);
var tbl1End = xml.indexOf('</TABLE>', tbl1Start) + 8;
var tbl2Start = xml.indexOf('<TABLE', tbl1End);
var tbl2End = xml.indexOf('</TABLE>', tbl2Start) + 8;

console.log('\n표1: ' + (tbl1End - tbl1Start) + '자');
console.log('표2: ' + (tbl2End - tbl2Start) + '자');
console.log('표 합계: ' + ((tbl1End-tbl1Start)+(tbl2End-tbl2Start)) + '자');
