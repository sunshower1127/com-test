/**
 * InsertPicture 후 반환된 ctrl의 Properties로 크기 조절
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

var imgPath = 'C:/Users/seng1/Documents/Projects/com-test/data/예시이미지.jpg';

insert('이미지 크기 조절 테스트\r\n\r\n');

// 이미지 삽입
var ctrl = bridge.comCallWith(h, 'InsertPicture', [imgPath, 1, 0, 0, 0, 0, 0, 0]);
console.log('InsertPicture 완료\n');

// 반환된 ctrl의 Properties 확인
var props = bridge.comGet(ctrl, 'Properties');
var sid = bridge.comGet(props, 'SetID');
console.log('Properties SetID: ' + sid);

// 어떤 Item이 있는지
var tryNames = ['Width', 'Height', 'TreatAsChar', 'TextWrap',
                'CurWidth', 'CurHeight', 'OriWidth', 'OriHeight',
                'InstWidth', 'InstHeight', 'ScaleX', 'ScaleY',
                'XPos', 'YPos', 'HorzOffset', 'VertOffset'];
console.log('\nProperties Items:');
tryNames.forEach(function(n) {
  try {
    if (bridge.comCallWith(props, 'ItemExist', [n])) {
      var v = bridge.comCallWith(props, 'Item', [n]);
      console.log('  ' + n + ' = ' + (typeof v === 'object' ? '[obj]' : v));
    }
  } catch(e) {}
});

// Width/Height 수정 시도
console.log('\n크기 수정 시도:');
try {
  bridge.comCallWith(props, 'SetItem', ['Width', 8000]);
  bridge.comCallWith(props, 'SetItem', ['Height', 10000]);
  console.log('SetItem 완료');

  // 적용 — ctrl의 Properties를 다시 세팅?
  bridge.comPut(ctrl, 'Properties', props);
  console.log('Properties 재설정 완료');
} catch(e) {
  console.log('❌ ' + e.message);
}

// 다른 방법: ShapeObjDialog로 크기 조절?
console.log('\n다른 방법들:');

// 방법 2: 이미지 선택 후 ShapeObject 액션
try {
  hwp.HAction.GetDefault('ShapeObjDialog', hwp.HParameterSet.HShapeObject.HSet);
  var hso = hwp.HParameterSet.HShapeObject;

  var hps = bridge.comGet(h, 'HParameterSet');
  var hsoHandle = bridge.comGet(hps, 'HShapeObject');
  var members = bridge.comListMembers(hsoHandle);
  var putProps = members.filter(function(m) {
    return (m.kind === 'put' || m.kind === 'get/put') && m.name !== 'HSet';
  });
  console.log('HShapeObject put 프로퍼티:');
  putProps.forEach(function(m) {
    try {
      var v = bridge.comGet(hsoHandle, m.name);
      if (typeof v !== 'object' && v !== 0) console.log('  ' + m.name + ' = ' + v);
    } catch(e) {}
  });

  // 크기 관련 프로퍼티 전체 출력
  console.log('\n전체 프로퍼티:');
  putProps.forEach(function(m) {
    console.log('  ' + m.name);
  });
} catch(e) {
  console.log('ShapeObjDialog: ❌ ' + e.message);
}

console.log('\n확인해주세요!');
