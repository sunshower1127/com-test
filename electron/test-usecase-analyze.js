/**
 * 이력서 템플릿 구조 분석
 */

var bridge = require('./native/com_bridge_node.node');
var { initBridge, createComProxy } = require('./dist/worker/proxy');

bridge.comInit();
initBridge(bridge);

var h = bridge.comCreate('HWPFrame.HwpObject');
var hwp = createComProxy(h);
var win0 = bridge.comCallWith(bridge.comGet(h, 'XHwpWindows'), 'Item', [0]);
bridge.comPut(win0, 'Visible', true);

// 파일 열기
var filePath = 'C:/Users/seng1/Documents/Projects/com-test/data/이력서_자기소개서_기본.hwp';
var ret = bridge.comCallWith(h, 'Open', [filePath, 'HWP', '']);
console.log('Open ret=' + ret + '\n');

// ── 1. 전체 텍스트 추출 ──
console.log('=== 1. 전체 텍스트 ===');
var text = String(bridge.comCallWith(h, 'GetTextFile', ['UNICODE', '']));
var lines = text.split(/\r?\n/);
lines.forEach(function(line, i) {
  if (line.trim()) console.log('[' + i + '] "' + line + '"');
});

// ── 2. 컨트롤 순회 ──
console.log('\n=== 2. 컨트롤 목록 ===');
var ctrl = bridge.comGet(h, 'HeadCtrl');
var ctrlCount = 0;
while (ctrl && typeof ctrl === 'object' && ctrlCount < 50) {
  var id = bridge.comGet(ctrl, 'CtrlID');
  var desc = bridge.comGet(ctrl, 'UserDesc');

  var extra = '';
  if (id === 'tbl') {
    var props = bridge.comGet(ctrl, 'Properties');
    try {
      var w = bridge.comCallWith(props, 'Item', ['Width']);
      var ht = bridge.comCallWith(props, 'Item', ['Height']);
      var rows = bridge.comCallWith(props, 'Item', ['RowCount']);
      var cols = bridge.comCallWith(props, 'Item', ['ColCount']);
      extra = ' W=' + w + ' H=' + ht + ' Rows=' + rows + ' Cols=' + cols;
    } catch(e) {
      // ItemExist로 찾기
      var names = ['Width', 'Height', 'RowCount', 'ColCount', 'Rows', 'Cols',
                   'CellSpacing', 'TableWidth'];
      names.forEach(function(n) {
        try {
          if (bridge.comCallWith(props, 'ItemExist', [n])) {
            extra += ' ' + n + '=' + bridge.comCallWith(props, 'Item', [n]);
          }
        } catch(e2) {}
      });
    }
  }

  console.log('[' + ctrlCount + '] ' + id + ' "' + desc + '"' + extra);
  ctrl = bridge.comGet(ctrl, 'Next');
  ctrlCount++;
}

// ── 3. 표 안으로 들어가서 셀 내용 읽기 ──
console.log('\n=== 3. 표 내용 (MoveDown 진입) ===');
bridge.comCallWith(h, 'Run', ['MoveDocBegin']);
bridge.comCallWith(h, 'Run', ['MoveDown']); // 표 진입

// 셀 순회 — TableRightCell로 이동하면서 읽기
var cellCount = 0;
var prevText = '';
for (var i = 0; i < 100; i++) { // 최대 100셀
  // 현재 셀 텍스트
  bridge.comCallWith(h, 'Run', ['MoveLineBegin']);
  bridge.comCallWith(h, 'Run', ['MoveSelLineEnd']);
  var cellText = String(bridge.comCallWith(h, 'GetTextFile', ['UNICODE', 'saveblock']));
  bridge.comCallWith(h, 'Run', ['Cancel']);

  // 위치 정보
  var ps = bridge.comCallWith(h, 'GetPosBySet', []);
  var list = bridge.comCallWith(ps, 'Item', ['List']);
  var para = bridge.comCallWith(ps, 'Item', ['Para']);

  if (cellText && cellText !== 'null' && cellText !== 'undefined') {
    console.log('  cell[' + cellCount + '] L=' + list + ' P=' + para + ' "' + cellText.substring(0, 40) + '"');
    cellCount++;
  }

  // 다음 셀
  bridge.comCallWith(h, 'Run', ['TableRightCell']);

  // 같은 셀에 머물면 표 끝
  var ps2 = bridge.comCallWith(h, 'GetPosBySet', []);
  var newList = bridge.comCallWith(ps2, 'Item', ['List']);
  if (i > 0 && newList === list && cellText === prevText) {
    // 마지막 셀에서 TableRightCell하면 첫 셀로 돌아가므로
    // 여기서 멈추지 말고 계속
  }
  prevText = cellText;

  // 표를 벗어났는지 확인 (List가 0이면 본문)
  if (newList === 0) {
    console.log('  (표 밖으로 나감)');
    break;
  }
}

// ── 4. 페이지 수 ──
console.log('\n=== 4. 문서 정보 ===');
try {
  var pageCount = bridge.comGet(h, 'PageCount');
  console.log('PageCount: ' + pageCount);
} catch(e) {}

// ── 5. 필드 확인 ──
try {
  var fields = String(bridge.comCallWith(h, 'GetFieldList', [0, 0]));
  if (fields && fields !== 'null') {
    console.log('Fields: "' + fields.split('\x02').join('", "') + '"');
  } else {
    console.log('Fields: 없음');
  }
} catch(e) { console.log('Fields: ❌'); }

console.log('\n(닫지 않음)');
