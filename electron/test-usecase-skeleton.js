/**
 * HWPML2X에서 스타일 제거, 구조+텍스트만 추출
 * → LLM에게 넘길 수 있는 가벼운 구조 맵 생성
 */

var fs = require('fs');
var buf = fs.readFileSync('C:/Users/seng1/Documents/Projects/com-test/data/이력서_hwpml2x.xml');
var xml = buf.toString('utf16le');

// ── 1. BODY만 추출 ──
var body = xml.substring(xml.indexOf('<BODY>'), xml.indexOf('</BODY>') + 7);

// ── 2. 표 파싱 ──
var tables = body.match(/<TABLE[\s\S]*?<\/TABLE>/g) || [];

function parseTables(tables) {
  var result = [];

  tables.forEach(function(tbl, ti) {
    var rows = tbl.match(/<ROW>[\s\S]*?<\/ROW>/g) || [];
    var tableData = { id: ti, rows: [] };

    rows.forEach(function(row, ri) {
      var cells = row.match(/<CELL[^>]*>[\s\S]*?<\/CELL>/g) || [];
      var rowData = { row: ri, cells: [] };

      cells.forEach(function(cell, ci) {
        // 셀 속성 추출
        var colAddr = (cell.match(/ColAddr="(\d+)"/) || [])[1];
        var rowAddr = (cell.match(/RowAddr="(\d+)"/) || [])[1];
        var colSpan = (cell.match(/ColSpan="(\d+)"/) || [])[1] || '1';
        var rowSpan = (cell.match(/RowSpan="(\d+)"/) || [])[1] || '1';

        // 텍스트 추출 — <CHAR>...</CHAR> 안의 내용
        var chars = cell.match(/<CHAR>([\s\S]*?)<\/CHAR>/g) || [];
        var text = '';
        chars.forEach(function(c) {
          var m = c.match(/<CHAR>([\s\S]*?)<\/CHAR>/);
          if (m) text += m[1].replace(/\n/g, '');
        });
        text = text.trim();

        // 빈 셀인지 판단
        var isEmpty = !text || text === '';

        rowData.cells.push({
          col: parseInt(colAddr),
          row: parseInt(rowAddr),
          colSpan: parseInt(colSpan),
          rowSpan: parseInt(rowSpan),
          text: text,
          empty: isEmpty,
        });
      });

      tableData.rows.push(rowData);
    });

    result.push(tableData);
  });

  return result;
}

var parsed = parseTables(tables);

// ── 3. 구조 맵 생성 (LLM용) ──
function generateStructureMap(parsed) {
  var output = '';

  parsed.forEach(function(table, ti) {
    // 표 제목 추론 — 첫 행의 텍스트
    var firstText = '';
    if (table.rows[0] && table.rows[0].cells[0]) {
      firstText = table.rows[0].cells[0].text;
    }
    output += '## 표' + (ti + 1) + (firstText ? ' (' + firstText + ')' : '') + '\n\n';

    table.rows.forEach(function(row) {
      var parts = row.cells.map(function(cell) {
        var span = '';
        if (cell.colSpan > 1 || cell.rowSpan > 1) {
          span = '[' + cell.colSpan + 'x' + cell.rowSpan + ']';
        }

        if (cell.empty) {
          return '{빈칸:r' + cell.row + 'c' + cell.col + '}' + span;
        } else {
          return '"' + cell.text.substring(0, 20) + '"' + span;
        }
      });

      output += '| ' + parts.join(' | ') + ' |\n';
    });

    output += '\n';
  });

  return output;
}

var structureMap = generateStructureMap(parsed);

console.log('=== 구조 맵 (LLM용) ===\n');
console.log(structureMap);
console.log('크기: ' + structureMap.length + '자 (' + Math.round(structureMap.length / 1024 * 100) / 100 + 'KB)');

// ── 4. JSON 형태도 생성 ──
function generateJSON(parsed) {
  return parsed.map(function(table, ti) {
    return {
      table: ti,
      rows: table.rows.map(function(row) {
        return row.cells.map(function(cell) {
          var obj = {};
          if (cell.text) obj.text = cell.text;
          if (cell.empty) obj.empty = true;
          obj.pos = 'r' + cell.row + 'c' + cell.col;
          if (cell.colSpan > 1) obj.colSpan = cell.colSpan;
          if (cell.rowSpan > 1) obj.rowSpan = cell.rowSpan;
          return obj;
        });
      })
    };
  });
}

var json = generateJSON(parsed);
var jsonStr = JSON.stringify(json, null, 2);

console.log('\n=== JSON 형태 ===\n');
console.log(jsonStr.substring(0, 2000) + '...');
console.log('\nJSON 크기: ' + jsonStr.length + '자 (' + Math.round(jsonStr.length / 1024 * 100) / 100 + 'KB)');

// 파일 저장
var dataDir = 'C:/Users/seng1/Documents/Projects/com-test/data/';
fs.writeFileSync(dataDir + '이력서_structure.txt', structureMap, 'utf8');
fs.writeFileSync(dataDir + '이력서_structure.json', jsonStr, 'utf8');
console.log('\n저장 완료: 이력서_structure.txt, 이력서_structure.json');
