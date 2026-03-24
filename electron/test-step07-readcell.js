/**
 * 표색깔.hwpx를 열고 B1 셀의 CellBorderFill 프로퍼티를 전부 읽기
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
var filePath = 'C:/Users/seng1/Documents/Projects/com-test/data/표색깔.hwpx';
var ret = bridge.comCallWith(h, 'Open', [filePath, 'HWPX', '']);
console.log('Open ret=' + ret + '\n');

// 표 안으로 진입 — MoveDocBegin 후 MoveRight
bridge.comCallWith(h, 'Run', ['MoveDocBegin']);

// 여러번 MoveRight 해서 표 안으로 들어가기 시도
for (var i = 0; i < 5; i++) {
  bridge.comCallWith(h, 'Run', ['MoveRight']);
  var ps = bridge.comCallWith(h, 'GetPosBySet', []);
  var list = bridge.comCallWith(ps, 'Item', ['List']);
  var para = bridge.comCallWith(ps, 'Item', ['Para']);
  var pos = bridge.comCallWith(ps, 'Item', ['Pos']);

  // 현재 텍스트 확인
  bridge.comCallWith(h, 'Run', ['MoveSelLineEnd']);
  var text = String(bridge.comCallWith(h, 'GetTextFile', ['UNICODE', 'saveblock']));
  bridge.comCallWith(h, 'Run', ['Cancel']);

  console.log('MoveRight[' + i + ']: List=' + list + ', Para=' + para + ', Pos=' + pos + ', text="' + text + '"');

  if (text === 'A1' || text === 'B1') break;
}

// A1 셀의 CellBorderFill 읽기
console.log('\n=== A1 셀 CellBorderFill ===');
function readCellBorderFill(label) {
  hwp.HAction.GetDefault('CellBorderFill', hwp.HParameterSet.HCellBorderFill.HSet);
  var hps = bridge.comGet(h, 'HParameterSet');
  var hcbf = bridge.comGet(hps, 'HCellBorderFill');
  var fillAttr = bridge.comGet(hcbf, 'FillAttr');

  console.log('── ' + label + ' ──');

  // 주요 CellBorderFill 프로퍼티
  var mainProps = ['BorderTypeTop', 'BorderTypeBottom', 'BorderTypeLeft', 'BorderTypeRight',
                   'BorderWidthTop', 'BorderWidthBottom', 'BorderWidthLeft', 'BorderWidthRight',
                   'BorderColorTop', 'BorderColorBottom', 'BorderCorlorLeft', 'BorderColorRight'];
  mainProps.forEach(function(p) {
    try {
      var v = bridge.comGet(hcbf, p);
      if (v !== 0) console.log('  ' + p + ' = ' + v);
    } catch(e) {}
  });

  // FillAttr 전체 읽기
  var fillProps = ['type', 'WindowsBrush', 'WinBrushFaceColor', 'WinBrushFaceStyle',
                   'WinBrushHatchColor', 'WinBrushAlpha',
                   'GradationType', 'GradationAngle', 'GradationColorNum',
                   'ImageBrush', 'Embedded', 'filename', 'FileNameStr',
                   'ToolbarColor'];
  console.log('  FillAttr:');
  fillProps.forEach(function(p) {
    try {
      var v = bridge.comGet(fillAttr, p);
      if (v !== 0 && v !== '' && v !== null) console.log('    ' + p + ' = ' + v);
    } catch(e) {}
  });

  // 전부 0이면 모든 값 출력
  var allFillMembers = bridge.comListMembers(fillAttr);
  var nonZero = [];
  allFillMembers.forEach(function(m) {
    if (m.kind === 'get' && m.name !== 'HSet') {
      try {
        var v = bridge.comGet(fillAttr, m.name);
        if (v !== 0 && v !== '' && v !== null && typeof v !== 'object') {
          nonZero.push(m.name + '=' + v);
        }
      } catch(e) {}
    }
  });
  if (nonZero.length > 0) {
    console.log('  FillAttr 0이 아닌 값: ' + nonZero.join(', '));
  } else {
    console.log('  FillAttr 전부 0');
  }
}

readCellBorderFill('A1');

// B1으로 이동
bridge.comCallWith(h, 'Run', ['TableRightCell']);
console.log('');
readCellBorderFill('B1 (빨간 배경)');

// A2로 이동
bridge.comCallWith(h, 'Run', ['TableLowerCell']);
bridge.comCallWith(h, 'Run', ['TableLeftCell']);
console.log('');
readCellBorderFill('A2 (기본)');

// B2로 이동
bridge.comCallWith(h, 'Run', ['TableRightCell']);
console.log('');
readCellBorderFill('B2 (기본)');

// hwp.CellShape도 비교
console.log('\n=== hwp.CellShape 비교 ===');
// B1으로 돌아가기
bridge.comCallWith(h, 'Run', ['TableUpperCell']);
try {
  var cs = bridge.comGet(h, 'CellShape');
  var members = bridge.comListMembers(cs);
  var names = members.filter(function(m) {
    return ['QueryInterface','AddRef','Release','GetTypeInfoCount','GetTypeInfo','GetIDsOfNames','Invoke'].indexOf(m.name) === -1;
  });
  console.log('CellShape 멤버:');
  names.forEach(function(m) {
    try {
      if (m.kind === 'method' && m.name === 'Item') return;
      if (m.kind === 'method' && m.name === 'ItemExist') {
        // 다양한 이름 시도
        var tryNames = ['BorderFillIDRef', 'borderFillIDRef', 'FillColor', 'BackColor',
                        'Width', 'Height', 'CellAddr', 'ColAddr', 'RowAddr',
                        'ColSpan', 'RowSpan', 'Protect', 'Header', 'VertAlign'];
        tryNames.forEach(function(n) {
          var exists = bridge.comCallWith(cs, 'ItemExist', [n]);
          if (exists) {
            var val = bridge.comCallWith(cs, 'Item', [n]);
            console.log('  ' + n + ' = ' + val);
          }
        });
        return;
      }
      if (m.kind === 'get') {
        var v = bridge.comGet(cs, m.name);
        if (typeof v !== 'object') console.log('  ' + m.name + ' = ' + v);
        else console.log('  ' + m.name + ' = [object]');
      }
    } catch(e) {}
  });
} catch(e) { console.log('CellShape: ❌ ' + e.message); }

console.log('\n완료 (닫지 않음)');
