/**
 * HParameterSet 패턴 테스트 — JS Proxy 경유
 * 공식 문서 예제가 그대로 작동하는지 확인
 */

const bridge = require('./native/com_bridge_node.node');
const { initBridge, createComProxy } = require('./dist/worker/proxy');

// COM 초기화
bridge.comInit();
initBridge(bridge);

// HWP 실행
console.log("=== HWP 실행 ===");
const hwpHandle = bridge.comCreate("HWPFrame.HwpObject");
const hwp = createComProxy(hwpHandle);

// 창 표시
const win0Handle = bridge.comCallWith(bridge.comGet(hwpHandle, "XHwpWindows"), "Item", [0]);
bridge.comPut(win0Handle, "Visible", false);
bridge.comPut(win0Handle, "Visible", true);

console.log("Version:", String(hwp.Version));
console.log("");

// ═══════════════════════════════════════
// 테스트 1: HParameterSet 패턴으로 InsertText
// 공식 예제:
//   hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
//   hwp.HParameterSet.HInsertText.Text = "안녕하세요"
//   hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)
// ═══════════════════════════════════════
console.log("=== 테스트 1: HParameterSet 패턴 InsertText ===");
try {
  hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet);
  hwp.HParameterSet.HInsertText.Text = "안녕하세요 HParameterSet!";
  hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet);
  console.log("  ✅ InsertText 성공!");
} catch (e) {
  console.log("  ❌ InsertText 실패:", e.message);
}

// 텍스트 확인
try {
  var text = bridge.comCallWith(hwpHandle, "GetTextFile", ["UNICODE", ""]);
  console.log("  텍스트 확인:", JSON.stringify(text));
} catch (e) {
  console.log("  텍스트 확인 실패:", e.message);
}

// ═══════════════════════════════════════
// 테스트 2: CreateAction 패턴으로 동일 작업 (비교용)
// ═══════════════════════════════════════
console.log("\n=== 테스트 2: CreateAction 패턴 InsertText (비교) ===");
try {
  hwp.Run("BreakPara");
  var act = hwp.CreateAction("InsertText");
  var set = act.CreateSet();
  act.GetDefault(set);
  set.SetItem("Text", "안녕하세요 CreateAction!");
  act.Execute(set);
  console.log("  ✅ InsertText 성공!");
} catch (e) {
  console.log("  ❌ InsertText 실패:", e.message);
}

// 텍스트 확인
try {
  var text = bridge.comCallWith(hwpHandle, "GetTextFile", ["UNICODE", ""]);
  console.log("  텍스트 확인:", JSON.stringify(text));
} catch (e) {
  console.log("  텍스트 확인 실패:", e.message);
}

// ═══════════════════════════════════════
// 테스트 3: HParameterSet.HCharShape로 서식 변경
// ═══════════════════════════════════════
console.log("\n=== 테스트 3: HParameterSet.HCharShape 서식 ===");
try {
  // 전체 선택
  hwp.Run("MoveDocBegin");
  hwp.Run("MoveSelDocEnd");

  // CharShape 적용
  hwp.HAction.GetDefault("CharShape", hwp.HParameterSet.HCharShape.HSet);
  hwp.HParameterSet.HCharShape.Height = 2400; // 24pt
  hwp.HParameterSet.HCharShape.Bold = 1;
  hwp.HAction.Execute("CharShape", hwp.HParameterSet.HCharShape.HSet);
  console.log("  ✅ CharShape 적용 성공! (24pt, Bold)");
} catch (e) {
  console.log("  ❌ CharShape 실패:", e.message);
}

// ═══════════════════════════════════════
// 테스트 4: HCharShape 프로퍼티 읽기
// ═══════════════════════════════════════
console.log("\n=== 테스트 4: HCharShape 프로퍼티 읽기 ===");
try {
  var height = hwp.HParameterSet.HCharShape.Height;
  console.log("  Height:", String(height));
  var bold = hwp.HParameterSet.HCharShape.Bold;
  console.log("  Bold:", String(bold));
  var faceName = hwp.HParameterSet.HCharShape.FaceNameHangul;
  console.log("  FaceNameHangul:", String(faceName));
} catch (e) {
  console.log("  ❌ 읽기 실패:", e.message);
}

console.log("\n=== 모든 테스트 완료 ===");
console.log("한글 창에서 결과를 확인하세요.");
console.log("Enter 키를 누르면 종료합니다...");

const readline = require('readline');
const rl = readline.createInterface({ input: process.stdin });
rl.on('line', () => {
  bridge.comCallWith(hwpHandle, "Clear", [1]);
  bridge.comCallWith(hwpHandle, "Quit", []);
  process.exit(0);
});
