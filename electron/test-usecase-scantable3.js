/**
 * scanTable v3 — 마지막 셀 감지 수정
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
 * scanTable v3
 * - TableRightCell로 순회
 * - List ID가 안 바뀌면 → 마지막 셀 (멈춤)
 * - 이미 본 List ID를 다시 만나면 → 순환 (멈춤)
 */
function scanTable() {
  var cells = [];
  var seen = {};

  var list = getList();
  var text = readCellText();
  cells.push({ list: list, text: text });
  seen[list] = true;

  for (var i = 0; i < 300; i++) {
    var prevList = getList();
    bridge.comCallWith(h, 'Run', ['TableRightCell']);
    var newList = getList();

    // 같은 셀에 머물면 → 마지막 셀
    if (newList === prevList) break;

    // 이미 본 셀이면 → 순환 완료
    if (seen[newList]) break;

    seen[newList] = true;
    cells.push({ list: newList, text: readCellText() });
  }

  return cells;
}

// ══════════════════════════════
// 테스트 1: 2x3 표
// ══════════════════════════════
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

// 첫 셀로
for (var i = 0; i < 10; i++) bridge.comCallWith(h, 'Run', ['TableLeftCell']);
for (var i = 0; i < 10; i++) bridge.comCallWith(h, 'Run', ['TableUpperCell']);

var result = scanTable();
console.log('셀 수: ' + result.length);
result.forEach(function(c, i) {
  console.log('  [' + i + '] L=' + c.list + ' "' + c.text + '"');
});

// SetPos 테스트
console.log('\n직접 접근:');
var target = result.filter(function(c) { return c.text === '25'; })[0];
if (target) {
  bridge.comCallWith(h, 'SetPos', [target.list, 0, 0]);
  bridge.comCallWith(h, 'Run', ['MoveSelLineEnd']);
  bridge.comCallWith(h, 'Run', ['Delete']);
  insert('30');
  console.log('  25 → 30 수정 ✅');
}

// ══════════════════════════════
// 테스트 2: 이력서
// ══════════════════════════════
console.log('\n=== 테스트 2: 이력서 ===\n');
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
      if (i + 1 < cells.length && !cells[i + 1].text) {
        bridge.comCallWith(h, 'SetPos', [cells[i + 1].list, 0, 0]);
        insert(value);
        console.log('  ✅ ' + labelPart + ' → "' + value + '"');
        return true;
      }
    }
  }
  console.log('  ❌ ' + labelPart);
  return false;
}

findAndFill(resume, '성', '홍길동');
findAndFill(resume, '생 년', '1990. 01. 15');
findAndFill(resume, '연락처', '010-1234-5678');
findAndFill(resume, 'E -', 'hong@example.com');
findAndFill(resume, '주소', '서울시 강남구');

console.log('\n확인해주세요!');
