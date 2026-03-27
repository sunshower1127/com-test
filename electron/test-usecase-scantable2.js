/**
 * scanTable v2 — 단순화: TableRightCell만으로 순회, List ID로 순환 감지
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

/**
 * scanTable v2 — TableRightCell만으로 모든 셀 순회
 *
 * HWP의 TableRightCell은 마지막 셀에서 첫 셀로 돌아감.
 * 그래서 첫 셀의 List ID를 기억하고, 다시 만나면 끝.
 */
function scanTable() {
  var cells = [];
  var firstList = getList();
  var firstText = readCellText();

  cells.push({ list: firstList, text: firstText });

  for (var i = 0; i < 300; i++) {
    bridge.comCallWith(h, 'Run', ['TableRightCell']);
    var list = getList();

    // 첫 셀로 돌아왔으면 끝
    if (list === firstList) break;

    cells.push({ list: list, text: readCellText() });
  }

  return cells;
}

// ══════════════════════════════════════
// 테스트 1: 간단한 2x3 표
// ══════════════════════════════════════
console.log('=== 테스트 1: 2x3 표 ===\n');
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

// 첫 셀로 돌아가기
for (var i = 0; i < 10; i++) bridge.comCallWith(h, 'Run', ['TableLeftCell']);
for (var i = 0; i < 10; i++) bridge.comCallWith(h, 'Run', ['TableUpperCell']);

var result = scanTable();
console.log('셀 수: ' + result.length);
result.forEach(function(c, i) {
  console.log('  [' + i + '] L=' + c.list + ' "' + c.text + '"');
});

// SetPos로 직접 접근
console.log('\n=== SetPos 직접 접근 ===');
var target = result.filter(function(c) { return c.text === '25'; })[0];
if (target) {
  bridge.comCallWith(h, 'SetPos', [target.list, 0, 0]);
  bridge.comCallWith(h, 'Run', ['MoveSelLineEnd']);
  bridge.comCallWith(h, 'Run', ['Delete']);
  insert('30');
  console.log('25 → 30 수정 (L=' + target.list + ')');
}

// ══════════════════════════════════════
// 테스트 2: 이력서
// ══════════════════════════════════════
console.log('\n\n=== 테스트 2: 이력서 ===\n');
bridge.comCallWith(h, 'Run', ['MoveDocEnd']);
bridge.comCallWith(h, 'Clear', [1]);

var filePath = 'C:/Users/seng1/Documents/Projects/com-test/data/이력서_자기소개서_기본.hwp';
bridge.comCallWith(h, 'Open', [filePath, 'HWP', '']);

bridge.comCallWith(h, 'Run', ['MoveDocBegin']);
bridge.comCallWith(h, 'Run', ['MoveDown']);

var resume = scanTable();
console.log('셀 수: ' + resume.length);
resume.forEach(function(c, i) {
  var label = c.text ? '"' + c.text.substring(0, 20) + '"' : '(빈칸)';
  console.log('  [' + i + '] L=' + c.list + ' ' + label);
});

// 데이터 넣기
console.log('\n=== 데이터 입력 ===');

function findAndFill(cells, labelPart, value) {
  for (var i = 0; i < cells.length; i++) {
    if (cells[i].text && cells[i].text.indexOf(labelPart) >= 0) {
      // 다음 빈 셀
      if (i + 1 < cells.length && !cells[i + 1].text) {
        bridge.comCallWith(h, 'SetPos', [cells[i + 1].list, 0, 0]);
        insert(value);
        console.log('  ✅ ' + labelPart + ' → "' + value + '" (L=' + cells[i + 1].list + ')');
        return true;
      }
    }
  }
  console.log('  ❌ ' + labelPart + ' 못 찾음');
  return false;
}

findAndFill(resume, '성', '홍길동');
findAndFill(resume, '생 년', '1990. 01. 15');
findAndFill(resume, '연락처', '010-1234-5678');
findAndFill(resume, 'E -', 'hong@example.com');
findAndFill(resume, '주소', '서울시 강남구');

console.log('\n확인해주세요!');
