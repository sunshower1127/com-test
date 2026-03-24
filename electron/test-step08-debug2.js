/**
 * InsertPicture 재시도 — 반환값(Dispatch) 처리
 */

var bridge = require('./native/com_bridge_node.node');
var { initBridge, createComProxy } = require('./dist/worker/proxy');
var fs = require('fs');
var path = require('path');

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

var testImg = 'C:/Windows/Web/Wallpaper/Windows/img0.jpg';

insert('Before image\r\n');

// InsertPicture — 반환값 Dispatch를 제대로 받기
console.log('=== InsertPicture ===');
try {
  // comCallWith로 호출 — 반환값은 Dispatch (그림 컨트롤 객체)
  var ctrl = bridge.comCallWith(h, 'InsertPicture', [testImg, 1, 0, 0, 0, 0, 0, 0]);
  console.log('반환 타입: ' + typeof ctrl);

  if (ctrl && typeof ctrl === 'object') {
    console.log('✅ InsertPicture 성공! Dispatch 객체 반환');

    // 컨트롤 멤버 확인
    try {
      var members = bridge.comListMembers(ctrl);
      var names = members.filter(function(m) {
        return ['QueryInterface','AddRef','Release','GetTypeInfoCount','GetTypeInfo','GetIDsOfNames','Invoke'].indexOf(m.name) === -1;
      }).map(function(m) { return m.name + '(' + m.kind + ')'; });
      console.log('컨트롤 멤버: ' + names.join(', '));
    } catch(e) { console.log('멤버 조회 실패: ' + e.message); }

    // 프로퍼티 읽기
    try {
      var ctrlID = bridge.comGet(ctrl, 'CtrlID');
      console.log('CtrlID: ' + ctrlID);
    } catch(e) {}
    try {
      var userDesc = bridge.comGet(ctrl, 'UserDesc');
      console.log('UserDesc: ' + userDesc);
    } catch(e) {}
  } else {
    console.log('⚠️ 반환값: ' + String(ctrl));
  }
} catch(e) {
  console.log('❌ ' + e.message);
}

insert('\r\nAfter image');

console.log('\n확인해주세요! 그림이 보이나요?');
