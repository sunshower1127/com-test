/**
 * 이력서 채우기 — 최종 프로토타입
 *
 * 파이프라인:
 * 1. Open → GetTextFile("HWPML2X") → 원본 XML 보관
 * 2. XML 파싱 → 격자 구조맵 (null로 병합 채움)
 * 3. fill(table, row, col, value) 호출들
 * 4. 원본 XML에서 해당 셀 수정 → SetTextFile
 */

var bridge = require('./native/com_bridge_node.node');
var { initBridge, createComProxy } = require('./dist/worker/proxy');

bridge.comInit();
initBridge(bridge);

var h = bridge.comCreate('HWPFrame.HwpObject');
var hwp = createComProxy(h);
var win0 = bridge.comCallWith(bridge.comGet(h, 'XHwpWindows'), 'Item', [0]);
bridge.comPut(win0, 'Visible', true);

// ════════════════════════════════════════
// 런타임 레이어
// ════════════════════════════════════════

var originalXml = null;  // 원본 HWPML2X 보관
var gridMap = null;       // 격자 구조맵
var cellIndex = null;     // (table,row,col) → XML CELL 위치 매핑

/**
 * 1단계: 문서 열고 구조 분석
 */
function analyzeDocument(filePath) {
  bridge.comCallWith(h, 'Open', [filePath, 'HWP', '']);

  // HWPML2X 추출 + 보관
  originalXml = String(bridge.comCallWith(h, 'GetTextFile', ['HWPML2X', '']));

  // 파싱
  var body = originalXml.substring(
    originalXml.indexOf('<BODY>'),
    originalXml.indexOf('</BODY>') + 7
  );

  var tables = body.match(/<TABLE[\s\S]*?<\/TABLE>/g) || [];
  gridMap = [];
  cellIndex = [];

  tables.forEach(function(tbl, ti) {
    var rows = tbl.match(/<ROW>[\s\S]*?<\/ROW>/g) || [];
    var tableGrid = [];
    var tableIndex = [];

    // 먼저 최대 열 수 파악
    var maxCol = 0;
    rows.forEach(function(row) {
      var cells = row.match(/<CELL[^>]*>/g) || [];
      cells.forEach(function(cellTag) {
        var col = parseInt((cellTag.match(/ColAddr="(\d+)"/) || [,'0'])[1]);
        var colSpan = parseInt((cellTag.match(/ColSpan="(\d+)"/) || [,'1'])[1]);
        maxCol = Math.max(maxCol, col + colSpan);
      });
    });

    // 격자 초기화
    var maxRow = rows.length + 10; // 여유
    var occupied = {}; // "r,c" → true (병합으로 차지된 자리)

    rows.forEach(function(row, ri) {
      var cells = row.match(/<CELL[^>]*>[\s\S]*?<\/CELL>/g) || [];
      var gridRow = new Array(maxCol).fill(null);
      var indexRow = new Array(maxCol).fill(null);

      cells.forEach(function(cell) {
        var cellTag = cell.match(/<CELL[^>]*>/)[0];
        var colAddr = parseInt((cellTag.match(/ColAddr="(\d+)"/) || [,'0'])[1]);
        var rowAddr = parseInt((cellTag.match(/RowAddr="(\d+)"/) || [,'0'])[1]);
        var colSpan = parseInt((cellTag.match(/ColSpan="(\d+)"/) || [,'1'])[1]);
        var rowSpan = parseInt((cellTag.match(/RowSpan="(\d+)"/) || [,'1'])[1]);

        // 텍스트 추출
        var chars = cell.match(/<CHAR>([\s\S]*?)<\/CHAR>/g) || [];
        var text = '';
        chars.forEach(function(c) {
          var m = c.match(/<CHAR>([\s\S]*?)<\/CHAR>/);
          if (m) text += m[1].replace(/\n/g, '');
        });
        text = text.trim();

        // 격자에 배치
        gridRow[colAddr] = {
          text: text || '',
          colSpan: colSpan,
          rowSpan: rowSpan,
          rowAddr: rowAddr,
          colAddr: colAddr
        };

        // XML에서 이 셀의 위치 기록
        indexRow[colAddr] = {
          rowAddr: rowAddr,
          colAddr: colAddr,
          // 빈 TEXT 태그 패턴 (수정 시 사용)
        };

        // 병합 영역 표시
        for (var dr = 0; dr < rowSpan; dr++) {
          for (var dc = 0; dc < colSpan; dc++) {
            if (dr === 0 && dc === 0) continue;
            occupied[(rowAddr + dr) + ',' + (colAddr + dc)] = true;
          }
        }
      });

      tableGrid.push(gridRow);
      tableIndex.push(indexRow);
    });

    gridMap.push(tableGrid);
    cellIndex.push(tableIndex);
  });

  return gridMap;
}

/**
 * 2단계: 구조맵을 사람이 읽을 수 있는 형태로
 */
function printStructure(gridMap) {
  var output = '';
  gridMap.forEach(function(table, ti) {
    output += '\n## 표' + ti + ' (' + table.length + '행)\n';
    table.forEach(function(row, ri) {
      var parts = [];
      row.forEach(function(cell, ci) {
        if (cell === null) {
          parts.push('--');
        } else if (cell.text) {
          var span = '';
          if (cell.colSpan > 1 || cell.rowSpan > 1) span = '[' + cell.colSpan + 'x' + cell.rowSpan + ']';
          parts.push('"' + cell.text.substring(0, 12) + '"' + span);
        } else {
          parts.push('{빈}');
        }
      });
      output += '  [행' + ri + '] ' + parts.join(' | ') + '\n';
    });
  });
  return output;
}

/**
 * 3단계: fill — XML에서 해당 셀 수정
 */
function fill(tableIdx, rowIdx, colIdx, value) {
  var cell = gridMap[tableIdx][rowIdx][colIdx];
  if (!cell) {
    console.log('  ⚠️ fill(' + tableIdx + ',' + rowIdx + ',' + colIdx + ') — null 셀 (병합 영역)');
    return false;
  }

  var rowAddr = cell.rowAddr;
  var colAddr = cell.colAddr;

  // 원본 XML에서 해당 CELL 찾기
  // RowAddr="N" ColAddr="M" 으로 정확히 매칭
  var cellPattern = 'RowAddr="' + rowAddr + '"';
  var colPattern = 'ColAddr="' + colAddr + '"';

  // 해당 테이블 영역에서 찾기
  var tableStart = originalXml.indexOf('<TABLE');
  for (var t = 0; t < tableIdx; t++) {
    tableStart = originalXml.indexOf('<TABLE', originalXml.indexOf('</TABLE>', tableStart) + 1);
  }
  var tableEnd = originalXml.indexOf('</TABLE>', tableStart) + 8;
  var tableXml = originalXml.substring(tableStart, tableEnd);

  // CELL 찾기 — RowAddr와 ColAddr 둘 다 매칭
  var searchStart = 0;
  while (true) {
    var cellStart = tableXml.indexOf('<CELL', searchStart);
    if (cellStart < 0) break;

    var cellTagEnd = tableXml.indexOf('>', cellStart);
    var cellTag = tableXml.substring(cellStart, cellTagEnd + 1);

    var hasRow = cellTag.indexOf('RowAddr="' + rowAddr + '"') >= 0;
    var hasCol = cellTag.indexOf('ColAddr="' + colAddr + '"') >= 0;

    if (hasRow && hasCol) {
      // 이 셀의 끝 찾기
      var cellEnd = tableXml.indexOf('</CELL>', cellStart) + 7;
      var cellContent = tableXml.substring(cellStart, cellEnd);

      // 빈 TEXT 태그를 값이 있는 걸로 교체
      var modified = cellContent;

      // 패턴 1: <TEXT CharShape="N"/> (자기 닫힘)
      modified = modified.replace(
        /<TEXT CharShape="(\d+)"\/>/,
        '<TEXT CharShape="$1"><CHAR>' + escapeXml(value) + '</CHAR></TEXT>'
      );

      // 패턴 2: <TEXT CharShape="N"></TEXT> (빈 태그)
      modified = modified.replace(
        /<TEXT CharShape="(\d+)"><\/TEXT>/,
        '<TEXT CharShape="$1"><CHAR>' + escapeXml(value) + '</CHAR></TEXT>'
      );

      // 패턴 3: 이미 텍스트가 있으면 교체
      modified = modified.replace(
        /<TEXT CharShape="(\d+)"><CHAR>[\s\S]*?<\/CHAR><\/TEXT>/,
        '<TEXT CharShape="$1"><CHAR>' + escapeXml(value) + '</CHAR></TEXT>'
      );

      // 원본 XML 업데이트
      var globalStart = tableStart + cellStart;
      var globalEnd = tableStart + cellEnd;
      originalXml = originalXml.substring(0, globalStart) + modified + originalXml.substring(globalEnd);

      console.log('  ✅ fill(' + tableIdx + ',' + rowIdx + ',' + colIdx + ') = "' + value + '"');
      return true;
    }

    searchStart = cellStart + 1;
  }

  console.log('  ❌ fill(' + tableIdx + ',' + rowIdx + ',' + colIdx + ') — 셀 못 찾음');
  return false;
}

function escapeXml(s) {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

/**
 * 4단계: 수정된 XML을 문서에 반영
 */
function commit() {
  bridge.comCallWith(h, 'Run', ['SelectAll']);
  bridge.comCallWith(h, 'Run', ['Delete']);
  var ret = bridge.comCallWith(h, 'SetTextFile', [originalXml, 'HWPML2X', '']);
  console.log('  commit ret=' + ret);
  return ret;
}

// ════════════════════════════════════════
// 실행
// ════════════════════════════════════════

console.log('=== 1. 문서 분석 ===\n');
var filePath = 'C:/Users/seng1/Documents/Projects/com-test/data/이력서_자기소개서_기본.hwp';
analyzeDocument(filePath);
console.log('표 수: ' + gridMap.length);
console.log(printStructure(gridMap));

console.log('\n=== 2. 데이터 채우기 ===\n');

// ── 표0 (이력서) ──
// 구조맵에서 확인한 좌표:
// 행1: 사진[2x4] | 성명 | {빈} | 생년월일 | {빈}
// 행2: (병합) | 연락처 | {빈} | 비상연락망 | {빈}
// 행3: (병합) | E-Mail | {빈}
// 행4: (병합) | 주소 | {빈}
// 행5: 학력사항[1x3] | 재학기간 | 학교명
// 행6: | {빈} | {빈}
// ...

// 기본 정보
fill(0, 1, 3, '홍길동');
fill(0, 1, 6, '1990. 01. 15');
fill(0, 2, 3, '010-1234-5678');
fill(0, 2, 6, '010-9876-5432');
fill(0, 3, 3, 'hong@example.com');
fill(0, 4, 3, '서울시 강남구 테헤란로 123');

// 학력
fill(0, 6, 1, '2009.03 ~ 2013.02');
fill(0, 6, 3, '서울대학교 (컴퓨터공학과)');
fill(0, 7, 1, '2013.03 ~ 2015.02');
fill(0, 7, 3, '서울대학교 대학원 (AI 전공)');

// 경력
fill(0, 9, 1, '2015.03 ~ 2020.12');
fill(0, 9, 3, '네이버');
fill(0, 9, 4, '검색개발팀 / 백엔드');
fill(0, 10, 1, '2021.01 ~ 현재');
fill(0, 10, 3, '카카오');
fill(0, 10, 4, 'AI팀 / ML 엔지니어');

// 자격
fill(0, 15, 1, '2014. 05');
fill(0, 15, 3, '정보처리기사');
fill(0, 15, 6, '한국산업인력공단');

// 업무능력
fill(0, 18, 1, 'Python, Java, JavaScript, Docker, K8s, SQL');

// ── 표1 (자기소개서) ──
fill(1, 0, 1, '어릴 때부터 컴퓨터에 관심이 많아 자연스럽게 프로그래밍을 시작했습니다. 대학에서 컴퓨터공학을 전공하며 알고리즘과 시스템 설계의 기초를 다졌습니다.');
fill(1, 1, 1, '꼼꼼하고 논리적인 성격으로 코드 리뷰에 강점이 있습니다. 다만 완벽주의 성향이 있어 때로는 일정 관리에 어려움을 겪기도 합니다.');
fill(1, 2, 1, '오픈소스 커뮤니티에서 활발히 활동하며, 대학 시절 IT 봉사동아리에서 3년간 활동했습니다.');
fill(1, 3, 1, 'AI 기술로 사람들의 삶을 더 편리하게 만들고 싶습니다. 귀사의 AI 연구 비전에 깊이 공감하여 지원합니다.');

console.log('\n=== 3. 커밋 (SetTextFile) ===\n');
commit();

// 검증
var result = String(bridge.comCallWith(h, 'GetTextFile', ['UNICODE', '']));
console.log('\n=== 4. 검증 (텍스트 추출) ===\n');
result.split(/\r?\n/).forEach(function(line, i) {
  if (line.trim()) console.log('[' + i + '] ' + line.trim());
});

console.log('\n확인해주세요!');
