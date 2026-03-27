// HWPML2X에서 표1의 행1(성명 행) 부분만 추출
var fs = require('fs');
var buf = fs.readFileSync('C:/Users/seng1/Documents/Projects/com-test/data/이력서_hwpml2x.xml');
var xml = buf.toString('utf16le');

// 첫 번째 ROW 몇 개만
var rows = xml.match(/<ROW[^>]*>[\s\S]*?<\/ROW>/g);
if (rows) {
  // 행 1 (성명 행) 출력
  console.log('=== 행1 (성명 행) ===\n');
  var row1 = rows[1]; // 0번은 지원분야
  // 정돈해서 출력
  var formatted = row1
    .replace(/></g, '>\n<')
    .replace(/(<\/?(ROW|CELL|TABLE|CHAR|TEXT|P |PARASHAPE|CHARSHAPE)[^>]*>)/gi, '\n$1');
  // 너무 길면 자르기
  var lines = formatted.split('\n').filter(function(l) { return l.trim(); });
  lines.forEach(function(l, i) {
    if (i < 60) console.log(l.trim());
  });

  console.log('\n\n=== 표2 행0 (성장과정) ===\n');
  // 표2의 첫 ROW 찾기 — 표2 시작점
  var tbl2Start = xml.indexOf('<TABLE', xml.indexOf('</TABLE>') + 1);
  var tbl2 = xml.substring(tbl2Start, xml.indexOf('</TABLE>', tbl2Start) + 8);
  var rows2 = tbl2.match(/<ROW[^>]*>[\s\S]*?<\/ROW>/g);
  if (rows2) {
    var formatted2 = rows2[0]
      .replace(/></g, '>\n<')
      .replace(/(<\/?(ROW|CELL|TABLE|CHAR|TEXT|P )[^>]*>)/gi, '\n$1');
    var lines2 = formatted2.split('\n').filter(function(l) { return l.trim(); });
    lines2.forEach(function(l, i) {
      if (i < 40) console.log(l.trim());
    });
  }
}
