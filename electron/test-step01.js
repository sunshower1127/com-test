/**
 * 단계 1: 문서 생명주기 — 한 단계씩 대화형 테스트
 */

var bridge = require('./native/com_bridge_node.node');
var { initBridge, createComProxy } = require('./dist/worker/proxy');
var fs = require('fs');
var path = require('path');
var readline = require('readline');

bridge.comInit();
initBridge(bridge);

var testDir = path.join(require('os').tmpdir(), 'hwp-step01');
if (!fs.existsSync(testDir)) fs.mkdirSync(testDir, { recursive: true });
var hwpPath = path.join(testDir, "test.hwp");

var hwpHandle = bridge.comCreate("HWPFrame.HwpObject");
var hwp = createComProxy(hwpHandle);
var win0 = bridge.comCallWith(bridge.comGet(hwpHandle, "XHwpWindows"), "Item", [0]);
bridge.comPut(win0, "Visible", false);
bridge.comPut(win0, "Visible", true);

console.log("HWP 실행 완료 (Version: " + String(hwp.Version) + ")");
console.log("테스트 경로: " + testDir);

var rl = readline.createInterface({ input: process.stdin, output: process.stdout });

function ask(msg, fn) {
  rl.question("\n" + msg + " (Enter로 실행) ", function() {
    try {
      fn();
    } catch(e) {
      console.log("  ❌ " + e.message);
    }
  });
}

// 순서대로 실행
console.log("\n── 1단계: 텍스트 삽입 ──");
hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet);
hwp.HParameterSet.HInsertText.Text = "단계 1 테스트 문서입니다.";
hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet);
console.log("  ✅ 텍스트 삽입 완료");

rl.question("\n── 2단계: SaveAs 시도. 보안 팝업이 뜨면 허용해주세요. Enter로 실행: ", function() {
  try {
    hwp.SaveAs(hwpPath, "HWP", "");
    console.log("  ✅ SaveAs 성공!");
    console.log("  파일: " + hwpPath);
    console.log("  크기: " + fs.statSync(hwpPath).size + " bytes");
  } catch(e) {
    console.log("  ❌ SaveAs 실패: " + e.message);
    console.log("  (팝업이 떴다면 허용 후 다시 시도해주세요)");
  }

  rl.question("\n── 3단계: Clear(1) 문서 닫기. Enter로 실행: ", function() {
    try {
      hwp.Clear(1);
      console.log("  ✅ Clear 완료. IsEmpty=" + String(hwp.IsEmpty));
    } catch(e) {
      console.log("  ❌ " + e.message);
    }

    rl.question("\n── 4단계: Open 시도. 보안 팝업이 뜨면 허용. Enter로 실행: ", function() {
      try {
        hwp.Open(hwpPath, "HWP", "");
        var text = bridge.comCallWith(hwpHandle, "GetTextFile", ["UNICODE", ""]);
        console.log("  ✅ Open 성공!");
        console.log("  텍스트: \"" + String(text).substring(0, 50) + "\"");
        console.log("  Path: " + String(hwp.Path));
      } catch(e) {
        console.log("  ❌ Open 실패: " + e.message);
      }

      rl.question("\n── 5단계: 종료. Enter를 누르면 한글 종료: ", function() {
        bridge.comCallWith(hwpHandle, "Clear", [1]);
        bridge.comCallWith(hwpHandle, "Quit", []);
        rl.close();
        process.exit(0);
      });
    });
  });
});
