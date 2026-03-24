/**
 * 단계 11: 컨트롤 순회 (HeadCtrl/LastCtrl) — 자동 테스트
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
function run(label, fn) {
  process.stdout.write('── ' + label + ': ');
  try { fn(); } catch(e) { console.log('❌ ' + e.message); }
}

// 문서에 다양한 컨트롤 넣기
insert('Text before table\r\n');

// 표 삽입
hwp.HAction.GetDefault('TableCreate', hwp.HParameterSet.HTableCreation.HSet);
hwp.HParameterSet.HTableCreation.Rows = 2;
hwp.HParameterSet.HTableCreation.Cols = 2;
hwp.HParameterSet.HTableCreation.WidthType = 2;
hwp.HAction.Execute('TableCreate', hwp.HParameterSet.HTableCreation.HSet);
insert('T1');
bridge.comCallWith(h, 'Run', ['TableRightCell']);
insert('T2');
bridge.comCallWith(h, 'Run', ['MoveDocEnd']);

insert('\r\nText between\r\n');

// 그림 삽입
var imgPath = 'C:/Windows/Web/Wallpaper/Windows/img0.jpg';
bridge.comCallWith(h, 'InsertPicture', [imgPath, 1, 0, 0, 0, 0, 0, 0]);

insert('\r\nText after image');

console.log('문서 준비 완료: 텍스트 + 표 + 텍스트 + 그림 + 텍스트\n');

// ================================================================
// 11-1: HeadCtrl
// ================================================================
run('11-1 HeadCtrl', function() {
  var ctrl = bridge.comGet(h, 'HeadCtrl');
  if (ctrl && typeof ctrl === 'object') {
    var id = bridge.comGet(ctrl, 'CtrlID');
    var desc = bridge.comGet(ctrl, 'UserDesc');
    console.log('✅ CtrlID="' + id + '", UserDesc="' + desc + '"');
  } else {
    console.log('❌ null');
  }
});

// ================================================================
// 11-2: HeadCtrl.Next 순회
// ================================================================
run('11-2 컨트롤 순회 (HeadCtrl → Next)', function() {
  var ctrl = bridge.comGet(h, 'HeadCtrl');
  var count = 0;
  console.log('');
  while (ctrl && typeof ctrl === 'object' && count < 20) {
    var id = bridge.comGet(ctrl, 'CtrlID');
    var desc = bridge.comGet(ctrl, 'UserDesc');
    var hasList = bridge.comGet(ctrl, 'HasList');
    console.log('  [' + count + '] CtrlID="' + id + '", UserDesc="' + desc + '", HasList=' + hasList);

    // Properties 접근
    try {
      var props = bridge.comGet(ctrl, 'Properties');
      if (props && typeof props === 'object') {
        var sid = bridge.comGet(props, 'SetID');
        var cnt = bridge.comGet(props, 'Count');
        console.log('       Properties: SetID="' + sid + '", Count=' + cnt);
      }
    } catch(e) {}

    ctrl = bridge.comGet(ctrl, 'Next');
    count++;
  }
  console.log('  총 ' + count + '개 컨트롤');
});

// ================================================================
// 11-3: LastCtrl
// ================================================================
run('\n11-3 LastCtrl', function() {
  var ctrl = bridge.comGet(h, 'LastCtrl');
  if (ctrl && typeof ctrl === 'object') {
    var id = bridge.comGet(ctrl, 'CtrlID');
    var desc = bridge.comGet(ctrl, 'UserDesc');
    console.log('✅ CtrlID="' + id + '", UserDesc="' + desc + '"');
  } else {
    console.log('❌ null');
  }
});

// ================================================================
// 11-4: LastCtrl.Prev 역순회
// ================================================================
run('11-4 역순회 (LastCtrl → Prev)', function() {
  var ctrl = bridge.comGet(h, 'LastCtrl');
  var count = 0;
  console.log('');
  while (ctrl && typeof ctrl === 'object' && count < 20) {
    var id = bridge.comGet(ctrl, 'CtrlID');
    var desc = bridge.comGet(ctrl, 'UserDesc');
    console.log('  [' + count + '] CtrlID="' + id + '", UserDesc="' + desc + '"');
    ctrl = bridge.comGet(ctrl, 'Prev');
    count++;
  }
  console.log('  총 ' + count + '개');
});

// ================================================================
// 11-5: 특정 컨트롤 Properties 상세
// ================================================================
run('\n11-5 표 컨트롤 Properties', function() {
  var ctrl = bridge.comGet(h, 'HeadCtrl');
  while (ctrl && typeof ctrl === 'object') {
    var id = bridge.comGet(ctrl, 'CtrlID');
    if (id === 'tbl') {
      var props = bridge.comGet(ctrl, 'Properties');
      var members = bridge.comListMembers(props);
      var names = members.filter(function(m) {
        return ['QueryInterface','AddRef','Release','GetTypeInfoCount','GetTypeInfo','GetIDsOfNames','Invoke'].indexOf(m.name) === -1;
      }).map(function(m) { return m.name; });
      console.log('표 Properties 멤버: ' + names.join(', '));

      // Item으로 값 읽기
      var tryItems = ['RowCount', 'ColCount', 'Rows', 'Cols', 'Width', 'Height',
                      'CellSpacing', 'TableWidth', 'TableHeight'];
      tryItems.forEach(function(n) {
        try {
          var exists = bridge.comCallWith(props, 'ItemExist', [n]);
          if (exists) {
            var v = bridge.comCallWith(props, 'Item', [n]);
            console.log('  ' + n + '=' + v);
          }
        } catch(e) {}
      });
      break;
    }
    ctrl = bridge.comGet(ctrl, 'Next');
  }
});

// ================================================================
// 11-6: GetAnchorPos
// ================================================================
run('\n11-6 GetAnchorPos', function() {
  var ctrl = bridge.comGet(h, 'HeadCtrl');
  while (ctrl && typeof ctrl === 'object') {
    var id = bridge.comGet(ctrl, 'CtrlID');
    if (id === 'tbl' || id === 'gso') {
      try {
        var anchorPos = bridge.comCallWith(ctrl, 'GetAnchorPos', [0]);
        if (anchorPos && typeof anchorPos === 'object') {
          var list = bridge.comCallWith(anchorPos, 'Item', ['List']);
          var para = bridge.comCallWith(anchorPos, 'Item', ['Para']);
          var pos = bridge.comCallWith(anchorPos, 'Item', ['Pos']);
          console.log(id + ': L=' + list + ', P=' + para + ', pos=' + pos);
        }
      } catch(e) { console.log(id + ' GetAnchorPos: ❌ ' + e.message); }
    }
    ctrl = bridge.comGet(ctrl, 'Next');
  }
});

// ================================================================
// 11-7: DeleteCtrl
// ================================================================
run('\n11-7 DeleteCtrl (그림 삭제)', function() {
  // 그림 컨트롤 찾기
  var ctrl = bridge.comGet(h, 'HeadCtrl');
  var gso = null;
  while (ctrl && typeof ctrl === 'object') {
    var id = bridge.comGet(ctrl, 'CtrlID');
    if (id === 'gso') { gso = ctrl; break; }
    ctrl = bridge.comGet(ctrl, 'Next');
  }

  if (gso) {
    var ret = bridge.comCallWith(h, 'DeleteCtrl', [gso]);
    console.log('DeleteCtrl ret=' + ret);
    // 확인
    var hasGso = false;
    ctrl = bridge.comGet(h, 'HeadCtrl');
    while (ctrl && typeof ctrl === 'object') {
      if (bridge.comGet(ctrl, 'CtrlID') === 'gso') { hasGso = true; break; }
      ctrl = bridge.comGet(ctrl, 'Next');
    }
    console.log('  삭제 후 gso 존재: ' + hasGso);
  } else {
    console.log('gso 없음');
  }
});

console.log('\n완료');
bridge.comCallWith(h, 'Clear', [1]);
bridge.comCallWith(h, 'Quit', []);
