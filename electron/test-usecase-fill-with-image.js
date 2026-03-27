/**
 * 이력서 채우기 + 사진 삽입
 *
 * 1. HWPML2X 라운드트립으로 텍스트 채우기 (사진 셀에 플레이스홀더)
 * 2. AllReplace로 플레이스홀더 찾기 → MoveToField 대신 Find로 위치
 * 3. InsertPicture로 이미지 삽입
 */

var bridge = require('./native/com_bridge_node.node');
var { initBridge, createComProxy } = require('./dist/worker/proxy');

bridge.comInit();
initBridge(bridge);

var h = bridge.comCreate('HWPFrame.HwpObject');
var hwp = createComProxy(h);
var win0 = bridge.comCallWith(bridge.comGet(h, 'XHwpWindows'), 'Item', [0]);
bridge.comPut(win0, 'Visible', true);

var originalXml = null;

function analyzeAndFill(filePath, fills) {
  bridge.comCallWith(h, 'Open', [filePath, 'HWP', '']);
  originalXml = String(bridge.comCallWith(h, 'GetTextFile', ['HWPML2X', '']));

  var body = originalXml.substring(
    originalXml.indexOf('<BODY>'),
    originalXml.indexOf('</BODY>') + 7
  );
  var tables = body.match(/<TABLE[\s\S]*?<\/TABLE>/g) || [];

  // 격자맵 구축
  var gridMap = [];
  tables.forEach(function(tbl, ti) {
    var rows = tbl.match(/<ROW>[\s\S]*?<\/ROW>/g) || [];
    var maxCol = 0;
    rows.forEach(function(row) {
      var cells = row.match(/<CELL[^>]*>/g) || [];
      cells.forEach(function(tag) {
        var col = parseInt((tag.match(/ColAddr="(\d+)"/) || [,'0'])[1]);
        var span = parseInt((tag.match(/ColSpan="(\d+)"/) || [,'1'])[1]);
        maxCol = Math.max(maxCol, col + span);
      });
    });

    var tableGrid = [];
    rows.forEach(function(row) {
      var cells = row.match(/<CELL[^>]*>[\s\S]*?<\/CELL>/g) || [];
      var gridRow = new Array(maxCol).fill(null);
      cells.forEach(function(cell) {
        var tag = cell.match(/<CELL[^>]*>/)[0];
        var colAddr = parseInt((tag.match(/ColAddr="(\d+)"/) || [,'0'])[1]);
        var rowAddr = parseInt((tag.match(/RowAddr="(\d+)"/) || [,'0'])[1]);
        var colSpan = parseInt((tag.match(/ColSpan="(\d+)"/) || [,'1'])[1]);
        var rowSpan = parseInt((tag.match(/RowSpan="(\d+)"/) || [,'1'])[1]);
        var chars = cell.match(/<CHAR>([\s\S]*?)<\/CHAR>/g) || [];
        var text = '';
        chars.forEach(function(c) {
          var m = c.match(/<CHAR>([\s\S]*?)<\/CHAR>/);
          if (m) text += m[1].replace(/\n/g, '');
        });
        gridRow[colAddr] = { text: text.trim(), rowAddr: rowAddr, colAddr: colAddr, colSpan: colSpan, rowSpan: rowSpan };
      });
      tableGrid.push(gridRow);
    });
    gridMap.push(tableGrid);
  });

  // fill 실행
  fills.forEach(function(f) {
    var cell = gridMap[f.table][f.row][f.col];
    if (!cell) { console.log('  ⚠️ (' + f.table + ',' + f.row + ',' + f.col + ') null'); return; }

    var rowAddr = cell.rowAddr;
    var colAddr = cell.colAddr;

    var tableStart = originalXml.indexOf('<TABLE');
    for (var t = 0; t < f.table; t++) {
      tableStart = originalXml.indexOf('<TABLE', originalXml.indexOf('</TABLE>', tableStart) + 1);
    }
    var tableEnd = originalXml.indexOf('</TABLE>', tableStart) + 8;
    var tableXml = originalXml.substring(tableStart, tableEnd);

    var searchStart = 0;
    while (true) {
      var cellStart = tableXml.indexOf('<CELL', searchStart);
      if (cellStart < 0) break;
      var cellTagEnd = tableXml.indexOf('>', cellStart);
      var cellTag = tableXml.substring(cellStart, cellTagEnd + 1);

      if (cellTag.indexOf('RowAddr="' + rowAddr + '"') >= 0 &&
          cellTag.indexOf('ColAddr="' + colAddr + '"') >= 0) {
        var cellEnd = tableXml.indexOf('</CELL>', cellStart) + 7;
        var cellContent = tableXml.substring(cellStart, cellEnd);
        var modified = cellContent;
        var val = f.value.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

        modified = modified.replace(/<TEXT CharShape="(\d+)"\/>/,
          '<TEXT CharShape="$1"><CHAR>' + val + '</CHAR></TEXT>');
        modified = modified.replace(/<TEXT CharShape="(\d+)"><\/TEXT>/,
          '<TEXT CharShape="$1"><CHAR>' + val + '</CHAR></TEXT>');

        var globalStart = tableStart + cellStart;
        var globalEnd = tableStart + cellEnd;
        originalXml = originalXml.substring(0, globalStart) + modified + originalXml.substring(globalEnd);
        console.log('  ✅ (' + f.table + ',' + f.row + ',' + f.col + ') = "' + f.value.substring(0, 30) + '"');
        break;
      }
      searchStart = cellStart + 1;
    }
  });

  // 커밋
  bridge.comCallWith(h, 'Run', ['SelectAll']);
  bridge.comCallWith(h, 'Run', ['Delete']);
  bridge.comCallWith(h, 'SetTextFile', [originalXml, 'HWPML2X', '']);
}

// ════════════════════════════════════════
// 실행
// ════════════════════════════════════════

console.log('=== 1. 텍스트 + 플레이스홀더 채우기 ===\n');

var filePath = 'C:/Users/seng1/Documents/Projects/com-test/data/이력서_자기소개서_기본.hwp';

analyzeAndFill(filePath, [
  // 사진 셀에 플레이스홀더
  { table: 0, row: 1, col: 0, value: '{{PHOTO}}' },

  // 기본 정보
  { table: 0, row: 1, col: 3, value: '홍길동' },
  { table: 0, row: 1, col: 6, value: '1990. 01. 15' },
  { table: 0, row: 2, col: 3, value: '010-1234-5678' },
  { table: 0, row: 2, col: 6, value: '010-9876-5432' },
  { table: 0, row: 3, col: 3, value: 'hong@example.com' },
  { table: 0, row: 4, col: 3, value: '서울시 강남구 테헤란로 123' },

  // 학력
  { table: 0, row: 6, col: 1, value: '2009.03 ~ 2013.02' },
  { table: 0, row: 6, col: 3, value: '서울대학교 (컴퓨터공학과)' },
  { table: 0, row: 7, col: 1, value: '2013.03 ~ 2015.02' },
  { table: 0, row: 7, col: 3, value: '서울대학교 대학원 (AI 전공)' },

  // 경력
  { table: 0, row: 9, col: 1, value: '2015.03 ~ 2020.12' },
  { table: 0, row: 9, col: 3, value: '네이버' },
  { table: 0, row: 9, col: 4, value: '검색개발팀 / 백엔드' },
  { table: 0, row: 10, col: 1, value: '2021.01 ~ 현재' },
  { table: 0, row: 10, col: 3, value: '카카오' },
  { table: 0, row: 10, col: 4, value: 'AI팀 / ML 엔지니어' },

  // 자격
  { table: 0, row: 15, col: 1, value: '2014. 05' },
  { table: 0, row: 15, col: 3, value: '정보처리기사' },
  { table: 0, row: 15, col: 6, value: '한국산업인력공단' },

  // 업무능력
  { table: 0, row: 18, col: 1, value: 'Python, Java, JavaScript, Docker, K8s' },

  // 자기소개서
  { table: 1, row: 0, col: 1, value: '어릴 때부터 컴퓨터에 관심이 많아 프로그래밍을 시작했습니다.' },
  { table: 1, row: 1, col: 1, value: '꼼꼼하고 논리적인 성격입니다.' },
  { table: 1, row: 2, col: 1, value: 'IT 봉사동아리에서 3년간 활동했습니다.' },
  { table: 1, row: 3, col: 1, value: 'AI 기술로 사람들의 삶을 편리하게 만들고 싶습니다.' },
]);

console.log('\n=== 2. 플레이스홀더 찾아서 이미지 삽입 ===\n');

// Find로 {{PHOTO}} 찾기
bridge.comCallWith(h, 'Run', ['MoveDocBegin']);

hwp.HAction.GetDefault('AllReplace', hwp.HParameterSet.HFindReplace.HSet);
hwp.HParameterSet.HFindReplace.FindString = '{{PHOTO}}';
hwp.HParameterSet.HFindReplace.ReplaceString = ' ';
hwp.HParameterSet.HFindReplace.IgnoreMessage = 1;
hwp.HParameterSet.HFindReplace.ReplaceMode = 1;
hwp.HAction.Execute('AllReplace', hwp.HParameterSet.HFindReplace.HSet);
console.log('플레이스홀더 제거 완료');

// 플레이스홀더가 있던 위치로 이동 — Find로 공백 찾기 대신
// 사진 셀은 표의 첫 번째 셀이니까 MoveDown으로 진입 가능
bridge.comCallWith(h, 'Run', ['MoveDocBegin']);
bridge.comCallWith(h, 'Run', ['MoveDown']);
// MoveDown은 표 안으로 들어감
// 사진 셀 찾기 — 현재 위치가 사진 셀일 수 있음

// 혹시 다른 셀이면 좌상단으로
for (var i = 0; i < 10; i++) bridge.comCallWith(h, 'Run', ['TableUpperCell']);
for (var i = 0; i < 10; i++) bridge.comCallWith(h, 'Run', ['TableLeftCell']);

// 기존 내용 삭제
bridge.comCallWith(h, 'Run', ['MoveLineBegin']);
bridge.comCallWith(h, 'Run', ['MoveSelLineEnd']);
bridge.comCallWith(h, 'Run', ['Delete']);

// 이미지 삽입
var imgPath = 'C:/Users/seng1/Documents/Projects/com-test/data/예시이미지.jpg';
try {
  var ctrl = bridge.comCallWith(h, 'InsertPicture', [imgPath, 1, 0, 0, 0, 0, 0, 0]);
  if (ctrl && typeof ctrl === 'object') {
    console.log('✅ 이미지 삽입 성공! CtrlID=' + bridge.comGet(ctrl, 'CtrlID'));
  } else {
    console.log('⚠️ 반환: ' + String(ctrl));
  }
} catch(e) {
  console.log('❌ ' + e.message);
}

console.log('\n확인해주세요!');
