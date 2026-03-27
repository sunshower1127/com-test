/**
 * scanTable() — 표의 모든 셀 List ID + 텍스트를 한번에 반환
 */

var bridge = require('./native/com_bridge_node.node');
var { initBridge, createComProxy } = require('./dist/worker/proxy');

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

/**
 * scanTable() — 현재 커서가 있는 표의 전체 구조를 스캔
 *
 * 반환: { rows: number, cols: number, cells: [{list, row, col, text}] }
 *
 * 원리:
 * 1. TableLeftCell/UpperCell 반복 → 첫 셀(A1)로 이동
 * 2. TableRightCell로 한 행의 셀 수 파악
 * 3. 모든 셀을 순회하면서 List ID + 텍스트 수집
 */
function scanTable() {
  // 1. 첫 셀로 이동
  for (var i = 0; i < 50; i++) {
    var before = getList();
    bridge.comCallWith(h, 'Run', ['TableUpperCell']);
    if (getList() === before) break; // 더 이상 위로 못 감
  }
  for (var i = 0; i < 50; i++) {
    var before = getList();
    bridge.comCallWith(h, 'Run', ['TableLeftCell']);
    if (getList() === before) break;
  }

  // 2. 전체 셀 순회
  var cells = [];
  var visited = {};
  var row = 0;
  var col = 0;
  var maxCol = 0;
  var firstListOfRow = getList();

  // 첫 행의 열 수 파악 + 데이터 수집
  while (true) {
    var list = getList();
    var key = 'L' + list;

    if (!visited[key]) {
      visited[key] = true;
      cells.push({
        list: list,
        row: row,
        col: col,
        text: readCellText()
      });
    }

    // 다음 셀
    var prevList = list;
    bridge.comCallWith(h, 'Run', ['TableRightCell']);
    var newList = getList();

    if (newList === firstListOfRow) {
      // 한 바퀴 돌았음 → 다음 행
      maxCol = Math.max(maxCol, col);
      row++;
      col = 0;
      firstListOfRow = newList;

      // 아래로 이동
      bridge.comCallWith(h, 'Run', ['TableLowerCell']);
      var afterDown = getList();

      if (afterDown === newList) {
        // 더 이상 아래로 못 감 → 끝
        break;
      }

      // 첫 열로 이동
      for (var i = 0; i < 50; i++) {
        var before = getList();
        bridge.comCallWith(h, 'Run', ['TableLeftCell']);
        if (getList() === before) break;
      }
      firstListOfRow = getList();

      // 이미 방문한 셀이면 끝 (병합 셀 때문에 순환 가능)
      if (visited['L' + firstListOfRow]) {
        // 행 시작 셀이 이미 방문됨 → 병합 행일 수 있음
        // 새로운 셀이 나올 때까지 오른쪽으로 탐색
        var foundNew = false;
        for (var j = 0; j < 50; j++) {
          var cl = getList();
          if (!visited['L' + cl]) {
            foundNew = true;
            firstListOfRow = cl;
            break;
          }
          bridge.comCallWith(h, 'Run', ['TableRightCell']);
        }
        if (!foundNew) break;
      }

      continue;
    }

    col++;

    // 안전장치
    if (cells.length > 200) break;
  }

  return cells;
}

function getList() {
  var ps = bridge.comCallWith(h, 'GetPosBySet', []);
  return bridge.comCallWith(ps, 'Item', ['List']);
}

function readCellText() {
  bridge.comCallWith(h, 'Run', ['MoveLineBegin']);
  bridge.comCallWith(h, 'Run', ['MoveSelLineEnd']);
  var t = String(bridge.comCallWith(h, 'GetTextFile', ['UNICODE', 'saveblock']));
  bridge.comCallWith(h, 'Run', ['Cancel']);
  if (t === 'null' || t === 'undefined') return '';
  return t.trim();
}

// ══════════════════════════════════════
// 테스트 1: 간단한 표
// ══════════════════════════════════════
console.log('=== 테스트 1: 간단한 2x3 표 ===\n');
hwp.HAction.GetDefault('TableCreate', hwp.HParameterSet.HTableCreation.HSet);
hwp.HParameterSet.HTableCreation.Rows = 2;
hwp.HParameterSet.HTableCreation.Cols = 3;
hwp.HParameterSet.HTableCreation.WidthType = 2;
hwp.HAction.Execute('TableCreate', hwp.HParameterSet.HTableCreation.HSet);

var testData = ['이름', '나이', '직업', '홍길동', '25', '개발자'];
for (var i = 0; i < 6; i++) {
  insert(testData[i]);
  if (i < 5) bridge.comCallWith(h, 'Run', ['TableRightCell']);
}

// 스캔
var result = scanTable();
console.log('셀 수: ' + result.length);
result.forEach(function(c) {
  console.log('  [r' + c.row + 'c' + c.col + '] List=' + c.list + ' "' + c.text + '"');
});

// ── scanTable 결과로 직접 셀 접근 ──
console.log('\n=== SetPos로 직접 접근 ===');
// "나이" 셀 찾기
var ageCell = result.filter(function(c) { return c.text === '25'; })[0];
if (ageCell) {
  bridge.comCallWith(h, 'SetPos', [ageCell.list, 0, 0]);
  bridge.comCallWith(h, 'Run', ['MoveSelLineEnd']);
  bridge.comCallWith(h, 'Run', ['Delete']);
  insert('30');
  console.log('나이 25 → 30 수정 완료 (List=' + ageCell.list + ')');
}

// ══════════════════════════════════════
// 테스트 2: 이력서
// ══════════════════════════════════════
console.log('\n\n=== 테스트 2: 이력서 ===\n');
bridge.comCallWith(h, 'Run', ['MoveDocEnd']);
bridge.comCallWith(h, 'Clear', [1]);
var filePath = 'C:/Users/seng1/Documents/Projects/com-test/data/이력서_자기소개서_기본.hwp';
bridge.comCallWith(h, 'Open', [filePath, 'HWP', '']);

// 표 진입
bridge.comCallWith(h, 'Run', ['MoveDocBegin']);
bridge.comCallWith(h, 'Run', ['MoveDown']);

var resume = scanTable();
console.log('이력서 셀 수: ' + resume.length);
resume.forEach(function(c) {
  var label = c.text ? '"' + c.text.substring(0, 15) + '"' : '(빈칸)';
  console.log('  [r' + c.row + 'c' + c.col + '] L=' + c.list + ' ' + label);
});

// ── 이력서에 데이터 넣기 ──
console.log('\n=== 이력서 데이터 입력 ===');

// "성명" 옆 빈칸 찾기: text가 "성"으로 시작하는 셀의 다음 셀
function findEmptyCellAfter(cells, labelPattern) {
  for (var i = 0; i < cells.length; i++) {
    if (cells[i].text && cells[i].text.indexOf(labelPattern) >= 0) {
      // 같은 행의 다음 빈 셀
      var targetRow = cells[i].row;
      for (var j = i + 1; j < cells.length; j++) {
        if (cells[j].row === targetRow && !cells[j].text) {
          return cells[j];
        }
      }
    }
  }
  return null;
}

function fillCell(cell, value) {
  if (!cell) return false;
  bridge.comCallWith(h, 'SetPos', [cell.list, 0, 0]);
  insert(value);
  console.log('  ✅ ' + value + ' (List=' + cell.list + ')');
  return true;
}

fillCell(findEmptyCellAfter(resume, '성'), '홍길동');
fillCell(findEmptyCellAfter(resume, '생 년'), '1990. 01. 15');
fillCell(findEmptyCellAfter(resume, '연락처'), '010-1234-5678');
fillCell(findEmptyCellAfter(resume, 'E -'), 'hong@example.com');
fillCell(findEmptyCellAfter(resume, '주소'), '서울시 강남구');

console.log('\n확인해주세요!');
