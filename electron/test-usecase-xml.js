/**
 * 이력서 HWPML2X 추출 → 표 구조 분석
 */

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

// HWPML2X 추출
var xml = String(bridge.comCallWith(h, 'GetTextFile', ['HWPML2X', '']));
console.log('XML length: ' + xml.length);

// 파일로 저장
fs.writeFileSync('C:/Users/seng1/Documents/Projects/com-test/data/이력서_hwpml2x.xml', xml, 'utf16le');
console.log('저장 완료: data/이력서_hwpml2x.xml');

// 간단한 표 구조 파싱 — <TABLE>, <ROW>, <CELL>, <CHAR> 태그 추출
var tables = xml.match(/<TABLE[^>]*>[\s\S]*?<\/TABLE>/g);
console.log('\n표 개수: ' + (tables ? tables.length : 0));

if (tables) {
  tables.forEach(function(tbl, ti) {
    console.log('\n=== 표 ' + (ti + 1) + ' ===');
    var rows = tbl.match(/<ROW[^>]*>[\s\S]*?<\/ROW>/g);
    if (!rows) {
      // 다른 태그명?
      rows = tbl.match(/<TR[^>]*>[\s\S]*?<\/TR>/g);
    }
    if (!rows) {
      rows = tbl.match(/<Row[^>]*>[\s\S]*?<\/Row>/g);
    }
    console.log('행 수: ' + (rows ? rows.length : '?'));

    if (rows) {
      rows.forEach(function(row, ri) {
        var cells = row.match(/<CELL[^>]*>[\s\S]*?<\/CELL>/g);
        if (!cells) cells = row.match(/<TD[^>]*>[\s\S]*?<\/TD>/g);
        if (!cells) cells = row.match(/<Cell[^>]*>[\s\S]*?<\/Cell>/g);

        var cellTexts = [];
        if (cells) {
          cells.forEach(function(cell) {
            // CHAR 태그에서 텍스트 추출
            var chars = cell.match(/<CHAR[^>]*>([^<]*)<\/CHAR>/g);
            if (!chars) chars = cell.match(/<Text[^>]*>([^<]*)<\/Text>/g);
            if (!chars) chars = cell.match(/<T>([^<]*)<\/T>/g);
            var text = '';
            if (chars) {
              chars.forEach(function(c) {
                var m = c.match(/>([^<]*)</);
                if (m) text += m[1];
              });
            }
            cellTexts.push(text.trim() || '(빈칸)');
          });
        }

        if (cellTexts.length > 0) {
          console.log('  [행' + ri + '] ' + cellTexts.join(' | '));
        }
      });
    }
  });
}

bridge.comCallWith(h, 'Clear', [1]);
bridge.comCallWith(h, 'Quit', []);
console.log('\n완료');
