/**
 * 단계 3 디버그 2: SetPos pos 동작 + GetPosBySet Item 읽기
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
function peekAfter(n) {
  for (var i = 0; i < n; i++) bridge.comCallWith(h, 'Run', ['MoveSelRight']);
  var t = bridge.comCallWith(h, 'GetTextFile', ['UNICODE', 'saveblock']);
  bridge.comCallWith(h, 'Run', ['Cancel']);
  return String(t);
}

insert('ABCDEFGHIJ');
console.log('텍스트: "ABCDEFGHIJ"\n');

// ── SetPos(0, 0, pos) — pos별 커서 위치 ──
console.log('── SetPos(0, 0, pos)');
for (var pos = 0; pos <= 12; pos++) {
  var ret = bridge.comCallWith(h, 'SetPos', [0, 0, pos]);
  var p = peekAfter(1);
  console.log('  pos=' + pos + ' → ret=' + ret + ', 뒤: "' + p + '"');
}

// ── 한글로도 확인 ──
console.log('\n── 한글 SetPos(0, 0, pos)');
bridge.comCallWith(h, 'Run', ['SelectAll']);
bridge.comCallWith(h, 'Run', ['Delete']);
insert('가나다라마');
console.log('텍스트: "가나다라마"');
for (var pos = 0; pos <= 12; pos++) {
  var ret = bridge.comCallWith(h, 'SetPos', [0, 0, pos]);
  var p = peekAfter(1);
  console.log('  pos=' + pos + ' → ret=' + ret + ', 뒤: "' + p + '"');
}

// ── GetPosBySet의 Item 값 읽기 ──
console.log('\n── GetPosBySet Item 값');
bridge.comCallWith(h, 'Run', ['SelectAll']);
bridge.comCallWith(h, 'Run', ['Delete']);
insert('ABCDEFGHIJ');

// 다양한 위치에서 GetPosBySet
var positions = [0, 3, 5, 9];
positions.forEach(function(target) {
  bridge.comCallWith(h, 'SetPos', [0, 0, target]);
  var posSet = bridge.comCallWith(h, 'GetPosBySet', []);
  try {
    var list = bridge.comCallWith(posSet, 'Item', ['List']);
    var para = bridge.comCallWith(posSet, 'Item', ['Para']);
    var pos = bridge.comCallWith(posSet, 'Item', ['Pos']);
    console.log('  SetPos(0,0,' + target + ') → GetPosBySet: List=' + list + ', Para=' + para + ', Pos=' + pos);
  } catch(e) {
    console.log('  Item("List/Para/Pos") 실패: ' + e.message);
    // ItemExist 확인
    try {
      var exists = bridge.comCallWith(posSet, 'ItemExist', ['List']);
      console.log('  ItemExist("List")=' + exists);
      exists = bridge.comCallWith(posSet, 'ItemExist', ['Para']);
      console.log('  ItemExist("Para")=' + exists);
      exists = bridge.comCallWith(posSet, 'ItemExist', ['Pos']);
      console.log('  ItemExist("Pos")=' + exists);
    } catch(e2) {}
    // Count와 아이템 이름 탐색
    try {
      var count = bridge.comGet(posSet, 'Count');
      console.log('  Count=' + count);
      // 다양한 이름 시도
      var names = ['list', 'para', 'pos', 'List', 'Para', 'Pos',
                   'nList', 'nPara', 'nPos', 'iList', 'iPara', 'iPos',
                   'ListParaPos', 'lp', 'pp', 'cp'];
      names.forEach(function(n) {
        try {
          var v = bridge.comCallWith(posSet, 'Item', [n]);
          if (v !== null && v !== undefined) {
            console.log('  Item("' + n + '")=' + v);
          }
        } catch(e3) {}
      });
    } catch(e3) {}
  }
});

// ── Run("MoveRight") 후 GetPosBySet — pos 값이 어떻게 증가하는지 ──
console.log('\n── MoveRight 후 GetPosBySet Pos 값 변화');
bridge.comCallWith(h, 'Run', ['MoveDocBegin']);
for (var i = 0; i <= 10; i++) {
  var posSet = bridge.comCallWith(h, 'GetPosBySet', []);
  try {
    var pos = bridge.comCallWith(posSet, 'Item', ['Pos']);
    var para = bridge.comCallWith(posSet, 'Item', ['Para']);
    console.log('  [' + i + '] Para=' + para + ', Pos=' + pos);
  } catch(e) {
    console.log('  [' + i + '] Item 실패');
    break;
  }
  bridge.comCallWith(h, 'Run', ['MoveRight']);
}

// 종료
console.log('\n── 완료');
bridge.comCallWith(h, 'Clear', [1]);
bridge.comCallWith(h, 'Quit', []);
