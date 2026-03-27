/**
 * 이미지 크기 조절 — HShapeObject 사용
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

insert('크기 조절 테스트\r\n');

// 이미지 삽입
var ctrl = bridge.comCallWith(h, 'InsertPicture', [imgPath, 1, 0, 0, 0, 0, 0, 0]);
console.log('원본 삽입 완료');

// 현재 크기 확인
var props = bridge.comGet(ctrl, 'Properties');
var origW = bridge.comCallWith(props, 'Item', ['Width']);
var origH = bridge.comCallWith(props, 'Item', ['Height']);
console.log('원본 크기: W=' + origW + ', H=' + origH);

// 방법 1: Properties SetItem → 이미지 선택 → ShapeObjDialog Execute
console.log('\n=== 방법 1: SelectCtrlFront → ShapeObjDialog ===');
try {
  // 이미지 바로 앞에 커서가 있을 것 — 선택
  bridge.comCallWith(h, 'Run', ['MoveLeft']);
  bridge.comCallWith(h, 'Run', ['SelectCtrlFront']);

  hwp.HAction.GetDefault('ShapeObjDialog', hwp.HParameterSet.HShapeObject.HSet);
  hwp.HParameterSet.HShapeObject.Width = 8000;  // 약 28mm
  hwp.HParameterSet.HShapeObject.Height = 10000; // 약 35mm
  hwp.HParameterSet.HShapeObject.TreatAsChar = 1;
  var ret = hwp.HAction.Execute('ShapeObjDialog', hwp.HParameterSet.HShapeObject.HSet);
  console.log('ret=' + ret);
  bridge.comCallWith(h, 'Run', ['Cancel']);
} catch(e) { console.log('❌ ' + e.message); }

// 크기 변했는지 확인
try {
  // ctrl 다시 읽기
  var newCtrl = bridge.comGet(h, 'LastCtrl');
  var id = bridge.comGet(newCtrl, 'CtrlID');
  if (id === 'gso') {
    var p2 = bridge.comGet(newCtrl, 'Properties');
    var newW = bridge.comCallWith(p2, 'Item', ['Width']);
    var newH = bridge.comCallWith(p2, 'Item', ['Height']);
    console.log('변경 후: W=' + newW + ', H=' + newH);
  }
} catch(e) {}

// 방법 2: ctrl.Properties에 직접 SetItem 후 ModifyProperties?
console.log('\n=== 방법 2: Properties 직접 수정 ===');
insert('\r\n\r\n두번째 이미지:\r\n');
var ctrl2 = bridge.comCallWith(h, 'InsertPicture', [imgPath, 1, 0, 0, 0, 0, 0, 0]);
var props2 = bridge.comGet(ctrl2, 'Properties');
console.log('삽입 후: W=' + bridge.comCallWith(props2, 'Item', ['Width']));

bridge.comCallWith(props2, 'SetItem', ['Width', 8000]);
bridge.comCallWith(props2, 'SetItem', ['Height', 10000]);
try {
  bridge.comPut(ctrl2, 'Properties', props2);
  console.log('Properties put 완료');
  // 다시 읽기
  var p3 = bridge.comGet(ctrl2, 'Properties');
  console.log('변경 후: W=' + bridge.comCallWith(p3, 'Item', ['Width']));
} catch(e) { console.log('❌ ' + e.message); }

console.log('\n확인해주세요!');
