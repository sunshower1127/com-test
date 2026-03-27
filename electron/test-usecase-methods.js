/**
 * 표 구조 파악 방법 비교
 */

var bridge = require('./native/com_bridge_node.node');
var { initBridge, createComProxy } = require('./dist/worker/proxy');

bridge.comInit();
initBridge(bridge);

var h = bridge.comCreate('HWPFrame.HwpObject');
var hwp = createComProxy(h);
var win0 = bridge.comCallWith(bridge.comGet(h, 'XHwpWindows'), 'Item', [0]);
bridge.comPut(win0, 'Visible', true);

var filePath = 'C:/Users/seng1/Documents/Projects/com-test/data/이력서_자기소개서_기본.hwp';
bridge.comCallWith(h, 'Open', [filePath, 'HWP', '']);

// ================================================================
// 방법 1: GetTextFile("UNICODE") — 순수 텍스트
// ================================================================
console.log('=== 방법 1: UNICODE (순수 텍스트) ===');
var unicode = String(bridge.comCallWith(h, 'GetTextFile', ['UNICODE', '']));
console.log('길이: ' + unicode.length);
console.log('처음 200자: "' + unicode.substring(0, 200).replace(/\r?\n/g, '|') + '"');
console.log('→ 표 구조 파악: ❌ (줄 구분만, 행/열 구분 없음)\n');

// ================================================================
// 방법 2: GetTextFile("HTML") — HTML
// ================================================================
console.log('=== 방법 2: HTML ===');
var html = String(bridge.comCallWith(h, 'GetTextFile', ['HTML', '']));
console.log('길이: ' + html.length);

// 표 태그 확인
var tableCount = (html.match(/<table/gi) || []).length;
var trCount = (html.match(/<tr/gi) || []).length;
var tdCount = (html.match(/<td/gi) || []).length;
console.log('table: ' + tableCount + ', tr: ' + trCount + ', td: ' + tdCount);

// 첫 번째 <table> 일부 출력
var tblStart = html.toLowerCase().indexOf('<table');
if (tblStart >= 0) {
  var snippet = html.substring(tblStart, tblStart + 500);
  console.log('첫 table: "' + snippet.substring(0, 300) + '..."');
}
console.log('→ 표 구조 파악: ✅ (HTML table/tr/td로 행/열 구분 가능)\n');

// ================================================================
// 방법 3: GetTextFile("HWPML2X") — HWPML XML
// ================================================================
console.log('=== 방법 3: HWPML2X ===');
var hwpml = String(bridge.comCallWith(h, 'GetTextFile', ['HWPML2X', '']));
console.log('길이: ' + hwpml.length);
var tblCount = (hwpml.match(/<TABLE/g) || []).length;
var rowCount = (hwpml.match(/<ROW/g) || []).length;
var cellCount = (hwpml.match(/<CELL/g) || []).length;
console.log('TABLE: ' + tblCount + ', ROW: ' + rowCount + ', CELL: ' + cellCount);
console.log('→ 표 구조 파악: ✅ (TABLE/ROW/CELL + 셀 속성 + 서식 전부)\n');

// ================================================================
// 방법 4: 컨트롤 순회 → 표 Properties
// ================================================================
console.log('=== 방법 4: 컨트롤 순회 ===');
var ctrl = bridge.comGet(h, 'HeadCtrl');
while (ctrl && typeof ctrl === 'object') {
  var id = bridge.comGet(ctrl, 'CtrlID');
  if (id === 'tbl') {
    var props = bridge.comGet(ctrl, 'Properties');
    var sid = bridge.comGet(props, 'SetID');
    var cnt = bridge.comGet(props, 'Count');
    console.log('표 Properties: SetID=' + sid + ', Count=' + cnt);

    // 어떤 Item이 있는지
    var tryNames = ['Width', 'Height', 'RowCount', 'ColCount', 'Rows', 'Cols',
                    'CellSpacing', 'BorderFill', 'RepeatHeader',
                    'PageBreak', 'ColSz', 'RowSz'];
    tryNames.forEach(function(n) {
      try {
        if (bridge.comCallWith(props, 'ItemExist', [n])) {
          var v = bridge.comCallWith(props, 'Item', [n]);
          console.log('  ' + n + '=' + (typeof v === 'object' ? '[obj]' : v));
        }
      } catch(e) {}
    });

    // CreateItemSet으로 하위 구조?
    try {
      var rowSet = bridge.comCallWith(props, 'CreateItemSet', ['Row', 'Row']);
      console.log('  CreateItemSet("Row"): ' + (rowSet ? 'exists' : 'null'));
    } catch(e) {}

    // 전체 Item 나열
    console.log('  모든 Item:');
    var allNames = ['Width', 'Height', 'CellSpacing', 'BorderFill',
                    'RepeatHeader', 'PageBreak', 'TextWrap', 'ZOrder',
                    'NumberingType', 'RowCnt', 'ColCnt', 'CellCnt',
                    'NoAdjust', 'BodyWidth', 'BodyHeight',
                    'TreatAsChar', 'Lock', 'DropcapStyle'];
    allNames.forEach(function(n) {
      try {
        if (bridge.comCallWith(props, 'ItemExist', [n])) {
          var v = bridge.comCallWith(props, 'Item', [n]);
          console.log('    ' + n + '=' + (typeof v === 'object' ? '[obj]' : v));
        }
      } catch(e) {}
    });

    break;
  }
  ctrl = bridge.comGet(ctrl, 'Next');
}
console.log('→ 표 구조 파악: ⚠️ (표 크기/속성은 있지만 개별 셀 내용 접근 불가)\n');

// ================================================================
// 방법 5: InitScan/GetText — 저수준 텍스트 스캔
// ================================================================
console.log('=== 방법 5: InitScan/GetText ===');
try {
  bridge.comCallWith(h, 'InitScan', []);
  var scanResults = [];
  for (var i = 0; i < 30; i++) {
    var result = bridge.comCallWith(h, 'GetText', []);
    // result는 보통 [상태코드, 텍스트] 같은 형태
    scanResults.push(String(result));
    if (String(result) === '1' || result === null) break;
  }
  bridge.comCallWith(h, 'ReleaseScan', []);
  console.log('GetText 결과 ' + scanResults.length + '개:');
  scanResults.forEach(function(r, i) {
    if (i < 15) console.log('  [' + i + '] ' + r.substring(0, 60));
  });
} catch(e) {
  console.log('InitScan: ❌ ' + e.message);
}

// ================================================================
// 방법 6: MoveDown + 셀 순회로 직접 읽기
// ================================================================
console.log('\n=== 방법 6: MoveDown + 셀 순회 ===');
bridge.comCallWith(h, 'Run', ['MoveDocBegin']);
bridge.comCallWith(h, 'Run', ['MoveDown']);

var cells = [];
var seen = {};
for (var i = 0; i < 50; i++) {
  var ps = bridge.comCallWith(h, 'GetPosBySet', []);
  var list = bridge.comCallWith(ps, 'Item', ['List']);
  var key = 'L' + list;

  if (!seen[key]) {
    bridge.comCallWith(h, 'Run', ['MoveLineBegin']);
    bridge.comCallWith(h, 'Run', ['MoveSelLineEnd']);
    var t = String(bridge.comCallWith(h, 'GetTextFile', ['UNICODE', 'saveblock']));
    bridge.comCallWith(h, 'Run', ['Cancel']);
    seen[key] = true;
    cells.push({ list: list, text: t.substring(0, 30) });
  }

  bridge.comCallWith(h, 'Run', ['TableRightCell']);
  var ps2 = bridge.comCallWith(h, 'GetPosBySet', []);
  var newList = bridge.comCallWith(ps2, 'Item', ['List']);
  if (newList === 0) break;
  // 순환 감지
  if (cells.length > 3 && newList === cells[0].list) break;
}
console.log('셀 ' + cells.length + '개:');
cells.forEach(function(c, i) {
  console.log('  [' + i + '] L=' + c.list + ' "' + c.text + '"');
});
console.log('→ 표 구조 파악: ⚠️ (셀 텍스트는 읽지만 행/열 위치 모름)\n');

// ================================================================
// 비교 요약
// ================================================================
console.log('=== 비교 요약 ===');
console.log('UNICODE:  텍스트만. 표 구조 ❌');
console.log('HTML:     table/tr/td ✅ 행/열 구분. 서식 일부 보존');
console.log('HWPML2X:  TABLE/ROW/CELL ✅ 완전한 구조 + 서식. SetTextFile 라운드트립 가능');
console.log('컨트롤:   표 존재/크기. 셀 내용 ❌');
console.log('셀 순회:  텍스트 읽기 가능. 행/열 위치 ❌');

bridge.comCallWith(h, 'Clear', [1]);
bridge.comCallWith(h, 'Quit', []);
