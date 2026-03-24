/**
 * 단계 2: 텍스트 입출력 — 자동 테스트 (비대화형)
 */

var bridge = require('./native/com_bridge_node.node');
var { initBridge, createComProxy } = require('./dist/worker/proxy');

bridge.comInit();
initBridge(bridge);

var h = bridge.comCreate('HWPFrame.HwpObject');
var hwp = createComProxy(h);
var win0 = bridge.comCallWith(bridge.comGet(h, 'XHwpWindows'), 'Item', [0]);
bridge.comPut(win0, 'Visible', true);

console.log('HWP Version: ' + String(hwp.Version));
console.log('');

function run(label, fn) {
  process.stdout.write('── ' + label + ': ');
  try {
    fn();
  } catch(e) {
    console.log('❌ ' + e.message);
  }
}

// 2-1: InsertText (HParameterSet)
run('2-1 InsertText (HParameterSet)', function() {
  hwp.HAction.GetDefault('InsertText', hwp.HParameterSet.HInsertText.HSet);
  hwp.HParameterSet.HInsertText.Text = '안녕하세요 텍스트 테스트입니다.';
  var ret = hwp.HAction.Execute('InsertText', hwp.HParameterSet.HInsertText.HSet);
  console.log('✅ ret=' + String(ret));
});

// 2-2: InsertText (CreateAction)
run('2-2 InsertText (CreateAction)', function() {
  bridge.comCallWith(h, 'Run', ['BreakPara']);
  var act = hwp.CreateAction('InsertText');
  var set = act.CreateSet();
  act.GetDefault(set);
  set.SetItem('Text', 'CreateAction 패턴으로 삽입!');
  var ret = act.Execute(set);
  console.log('✅ ret=' + String(ret));
});

// 2-3: 여러 줄 삽입
run('2-3 여러 줄 삽입 (\\r\\n)', function() {
  bridge.comCallWith(h, 'Run', ['BreakPara']);
  hwp.HAction.GetDefault('InsertText', hwp.HParameterSet.HInsertText.HSet);
  hwp.HParameterSet.HInsertText.Text = '첫째줄\r\n둘째줄\r\n셋째줄';
  var ret = hwp.HAction.Execute('InsertText', hwp.HParameterSet.HInsertText.HSet);
  console.log('✅ ret=' + String(ret));
});

// 2-4: GetTextFile("UNICODE")
run('2-4 GetTextFile("UNICODE")', function() {
  var text = bridge.comCallWith(h, 'GetTextFile', ['UNICODE', '']);
  console.log('✅ len=' + String(text).length + ' | "' + String(text).substring(0, 80) + '..."');
});

// 2-5: GetTextFile("TEXT")
run('2-5 GetTextFile("TEXT")', function() {
  var text = bridge.comCallWith(h, 'GetTextFile', ['TEXT', '']);
  if (text === null || text === undefined || text === false) {
    console.log('⚠️ 반환=' + String(text));
  } else {
    console.log('✅ len=' + String(text).length + ' | "' + String(text).substring(0, 80) + '..."');
  }
});

// 2-6: GetTextFile("UTF8")
run('2-6 GetTextFile("UTF8")', function() {
  var text = bridge.comCallWith(h, 'GetTextFile', ['UTF8', '']);
  if (text === null || text === undefined || text === false) {
    console.log('⚠️ 반환=' + String(text));
  } else {
    console.log('✅ len=' + String(text).length + ' | "' + String(text).substring(0, 80) + '..."');
  }
});

// 2-7: GetTextFile("HTML")
run('2-7 GetTextFile("HTML")', function() {
  var text = bridge.comCallWith(h, 'GetTextFile', ['HTML', '']);
  if (text === null || text === undefined || text === false) {
    console.log('⚠️ 반환=' + String(text));
  } else {
    console.log('✅ len=' + String(text).length + ' | "' + String(text).substring(0, 120) + '..."');
  }
});

// 2-8: GetTextFile("HWPML2X")
run('2-8 GetTextFile("HWPML2X")', function() {
  var text = bridge.comCallWith(h, 'GetTextFile', ['HWPML2X', '']);
  if (text === null || text === undefined || text === false) {
    console.log('⚠️ 반환=' + String(text));
  } else {
    console.log('✅ len=' + String(text).length + ' | "' + String(text).substring(0, 120) + '..."');
  }
});

// 2-9: SetTextFile 전체 교체
run('2-9 SetTextFile 전체 교체', function() {
  var ret = bridge.comCallWith(h, 'SetTextFile', ['SetTextFile로 교체된 내용입니다.', 'UNICODE', '']);
  console.log('ret=' + String(ret));
  var text = bridge.comCallWith(h, 'GetTextFile', ['UNICODE', '']);
  console.log('         교체 후: "' + String(text).substring(0, 80) + '"');
});

// 2-10: SetTextFile insertfile
run('2-10 SetTextFile insertfile', function() {
  var ret = bridge.comCallWith(h, 'SetTextFile', ['\r\n추가 삽입된 텍스트!', 'UNICODE', 'insertfile']);
  console.log('ret=' + String(ret));
  var text = bridge.comCallWith(h, 'GetTextFile', ['UNICODE', '']);
  console.log('         삽입 후: "' + String(text).substring(0, 120) + '"');
});

// 2-11: GetPageText
run('2-11 GetPageText(0)', function() {
  var text = bridge.comCallWith(h, 'GetPageText', [0, '']);
  if (text === null || text === undefined || text === false) {
    console.log('⚠️ pgno=0 반환=' + String(text));
    // 1-based 시도
    try {
      var text2 = bridge.comCallWith(h, 'GetPageText', [1, '']);
      console.log('  pgno=1: "' + String(text2).substring(0, 80) + '"');
    } catch(e2) {
      console.log('  pgno=1도 실패: ' + e2.message);
    }
  } else {
    console.log('✅ "' + String(text).substring(0, 80) + '"');
  }
});

// 2-12: GetHeadingString
run('2-12 GetHeadingString()', function() {
  var heading = bridge.comCallWith(h, 'GetHeadingString', []);
  console.log('✅ "' + String(heading) + '" (빈 문자열이면 정상)');
});

// 2-13: BreakPara
run('2-13 BreakPara x2 + 텍스트', function() {
  bridge.comCallWith(h, 'Run', ['BreakPara']);
  bridge.comCallWith(h, 'Run', ['BreakPara']);
  hwp.HAction.GetDefault('InsertText', hwp.HParameterSet.HInsertText.HSet);
  hwp.HParameterSet.HInsertText.Text = '← 위에 빈 줄 2개';
  hwp.HAction.Execute('InsertText', hwp.HParameterSet.HInsertText.HSet);
  var text = bridge.comCallWith(h, 'GetTextFile', ['UNICODE', '']);
  console.log('✅ 전체: "' + String(text).substring(0, 150) + '"');
});

// 2-14: 선택 영역 GetTextFile (프로그래밍으로 선택)
run('2-14 SelectText + GetTextFile saveblock', function() {
  // 문서 처음의 첫 5글자 선택
  bridge.comCallWith(h, 'Run', ['MoveDocBegin']);
  var sel = bridge.comCallWith(h, 'SelectText', [0, 0, 0, 5]);
  console.log('SelectText ret=' + String(sel));
  var text = bridge.comCallWith(h, 'GetTextFile', ['UNICODE', 'saveblock']);
  if (text === null || text === undefined || text === '' || text === false) {
    console.log('         ⚠️ saveblock 반환=' + String(text));
  } else {
    console.log('         ✅ 선택 텍스트: "' + String(text) + '"');
  }
});

// 종료
console.log('\n── 테스트 완료. 종료 중...');
bridge.comCallWith(h, 'Clear', [1]);
bridge.comCallWith(h, 'Quit', []);
console.log('완료');
