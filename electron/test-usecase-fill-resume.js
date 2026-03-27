/**
 * 이력서 양식에 더미 데이터 채우기
 *
 * 전략: HWPML2X로 구조 파악 → x,y 좌표로 셀 이동 → InsertText
 */

var bridge = require('./native/com_bridge_node.node');
var { initBridge, createComProxy } = require('./dist/worker/proxy');
var readline = require('readline');

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

var rl = readline.createInterface({ input: process.stdin, output: process.stdout });
function ask(msg) {
  return new Promise(function(resolve) {
    rl.question('\n' + msg + ' → Enter ', function() { resolve(); });
  });
}

// ── 표 진입 후 좌상단으로 이동하는 헬퍼 ──
function enterTable() {
  bridge.comCallWith(h, 'Run', ['MoveDocBegin']);
  bridge.comCallWith(h, 'Run', ['MoveDown']);
  // 좌상단으로
  for (var i = 0; i < 30; i++) bridge.comCallWith(h, 'Run', ['TableUpperCell']);
  for (var i = 0; i < 30; i++) bridge.comCallWith(h, 'Run', ['TableLeftCell']);
}

// ── 좌상단에서 x,y만큼 이동 ──
function moveToCell(row, col) {
  // 좌상단으로 먼저
  for (var i = 0; i < 30; i++) bridge.comCallWith(h, 'Run', ['TableUpperCell']);
  for (var i = 0; i < 30; i++) bridge.comCallWith(h, 'Run', ['TableLeftCell']);
  // row만큼 아래
  for (var i = 0; i < row; i++) bridge.comCallWith(h, 'Run', ['TableLowerCell']);
  // col만큼 오른쪽
  for (var i = 0; i < col; i++) bridge.comCallWith(h, 'Run', ['TableRightCell']);
}

// ── 현재 셀에 값 입력 (기존 내용 삭제 후) ──
function fillCell(value) {
  bridge.comCallWith(h, 'Run', ['MoveLineBegin']);
  bridge.comCallWith(h, 'Run', ['MoveSelLineEnd']);
  bridge.comCallWith(h, 'Run', ['Delete']);
  insert(value);
}

// ── 좌상단에서 x,y로 이동 후 값 입력 ──
function fillAt(row, col, value, label) {
  moveToCell(row, col);
  fillCell(value);
  console.log('  ✅ ' + (label || '') + ' [r' + row + ',c' + col + '] = "' + value + '"');
}

(async function() {

  // 이력서 열기
  var filePath = 'C:/Users/seng1/Documents/Projects/com-test/data/이력서_자기소개서_기본.hwp';
  bridge.comCallWith(h, 'Open', [filePath, 'HWP', '']);
  console.log('이력서 열림\n');

  // ── HWPML2X에서 구조 확인 (이미 분석 완료) ──
  // 행0: [지원분야(8열 병합)]
  // 행1: [사진(2x4)] [성명] [이름빈칸(2열)] [생년월일] [생일빈칸(2열)]
  // 행2: [연락처] [전화빈칸(2열)] [비상연락망] [비상빈칸(2열)]
  // 행3: [E-Mail] [이메일빈칸(5열)]
  // 행4: [주소] [주소빈칸(5열)]
  // 행5: [학력사항(1x3)] [재학기간(2열)] [학교명(5열)]
  // 행6~7: 학력 데이터 행
  // 행8: [경력사항(1x4)] [근무기간(2열)] [근무회사] [근무부서(4열)]
  // 행9~11: 경력 데이터 행
  // ...

  // ── 표1(이력서) 진입 ──
  enterTable();
  console.log('=== 이력서 데이터 입력 ===');

  await ask('── 기본 정보 입력');

  // 행0: 지원분야 — 이미 텍스트가 있으므로 뒤에 추가
  moveToCell(0, 0);
  bridge.comCallWith(h, 'Run', ['MoveLineEnd']);
  insert(' 개발팀');
  console.log('  ✅ 지원분야 = "개발팀"');

  // 행1: 성명 옆 빈칸
  // 사진(c0, 2열) → 성명(c2) → 빈칸(c3) → 생년월일(c5) → 빈칸(c6)
  // 병합 때문에 TableRightCell 횟수가 ColAddr와 다를 수 있음
  // 전략: "성명" 셀에서 TableRightCell 1번 → 빈칸
  moveToCell(1, 1); // 성명 셀
  bridge.comCallWith(h, 'Run', ['TableRightCell']); // 빈칸으로
  fillCell('홍길동');
  console.log('  ✅ 성명 = "홍길동"');

  // 생년월일: 빈칸에서 TableRightCell → 생년월일 → TableRightCell → 빈칸
  bridge.comCallWith(h, 'Run', ['TableRightCell']); // 생년월일 레이블
  bridge.comCallWith(h, 'Run', ['TableRightCell']); // 빈칸
  fillCell('1990. 01. 15');
  console.log('  ✅ 생년월일 = "1990. 01. 15"');

  await ask('── 연락처 입력');

  // 행2: 연락처 → 빈칸 → 비상연락망 → 빈칸
  moveToCell(2, 0); // 연락처
  bridge.comCallWith(h, 'Run', ['TableRightCell']); // 빈칸
  fillCell('010-1234-5678');
  console.log('  ✅ 연락처 = "010-1234-5678"');

  bridge.comCallWith(h, 'Run', ['TableRightCell']); // 비상연락망
  bridge.comCallWith(h, 'Run', ['TableRightCell']); // 빈칸
  fillCell('010-9876-5432');
  console.log('  ✅ 비상연락망 = "010-9876-5432"');

  // 행3: E-Mail
  moveToCell(3, 0); // E-Mail
  bridge.comCallWith(h, 'Run', ['TableRightCell']);
  fillCell('hong@example.com');
  console.log('  ✅ E-Mail = "hong@example.com"');

  // 행4: 주소
  moveToCell(4, 0); // 주소
  bridge.comCallWith(h, 'Run', ['TableRightCell']);
  fillCell('서울시 강남구 테헤란로 123');
  console.log('  ✅ 주소 = "서울시 강남구 테헤란로 123"');

  await ask('── 학력 입력');

  // 행6: 학력 첫 줄 (재학기간 | 학교명)
  moveToCell(6, 0);
  bridge.comCallWith(h, 'Run', ['TableRightCell']); // 재학기간 빈칸
  fillCell('2009.03 ~ 2013.02');
  console.log('  ✅ 학력1 기간');

  bridge.comCallWith(h, 'Run', ['TableRightCell']); // 학교명 빈칸
  fillCell('서울대학교 (컴퓨터공학과)');
  console.log('  ✅ 학력1 학교');

  // 행7: 학력 둘째 줄
  moveToCell(7, 0);
  bridge.comCallWith(h, 'Run', ['TableRightCell']);
  fillCell('2013.03 ~ 2015.02');
  console.log('  ✅ 학력2 기간');

  bridge.comCallWith(h, 'Run', ['TableRightCell']);
  fillCell('서울대학교 대학원 (컴퓨터공학)');
  console.log('  ✅ 학력2 학교');

  await ask('── 경력 입력');

  // 행9: 경력 첫 줄 (근무기간 | 근무회사 | 근무부서)
  moveToCell(9, 0);
  bridge.comCallWith(h, 'Run', ['TableRightCell']);
  fillCell('2015.03 ~ 2020.12');
  console.log('  ✅ 경력1 기간');

  bridge.comCallWith(h, 'Run', ['TableRightCell']);
  fillCell('네이버');
  console.log('  ✅ 경력1 회사');

  bridge.comCallWith(h, 'Run', ['TableRightCell']);
  fillCell('검색개발팀 / 백엔드 개발');
  console.log('  ✅ 경력1 부서');

  // 행10: 경력 둘째 줄
  moveToCell(10, 0);
  bridge.comCallWith(h, 'Run', ['TableRightCell']);
  fillCell('2021.01 ~ 현재');

  bridge.comCallWith(h, 'Run', ['TableRightCell']);
  fillCell('카카오');

  bridge.comCallWith(h, 'Run', ['TableRightCell']);
  fillCell('AI팀 / ML 엔지니어');
  console.log('  ✅ 경력2');

  await ask('── 자격/면허, 업무능력 입력');

  // 행15: 자격 첫 줄
  moveToCell(15, 0);
  bridge.comCallWith(h, 'Run', ['TableRightCell']);
  fillCell('2014. 05');
  bridge.comCallWith(h, 'Run', ['TableRightCell']);
  fillCell('정보처리기사');
  bridge.comCallWith(h, 'Run', ['TableRightCell']);
  fillCell('한국산업인력공단');
  console.log('  ✅ 자격1');

  // 행18: 업무능력
  moveToCell(18, 0);
  bridge.comCallWith(h, 'Run', ['TableRightCell']);
  fillCell('Python, Java, JavaScript, SQL, Docker, Kubernetes');
  console.log('  ✅ 업무능력');

  // ── 서명란 ──
  // 문서 끝쪽 "년    월    일" 부분
  bridge.comCallWith(h, 'Run', ['MoveDocEnd']);
  // 전체 텍스트에서 AllReplace로 처리
  bridge.comCallWith(h, 'Run', ['MoveDocBegin']);
  hwp.HAction.GetDefault('AllReplace', hwp.HParameterSet.HFindReplace.HSet);
  hwp.HParameterSet.HFindReplace.FindString = '년    월    일';
  hwp.HParameterSet.HFindReplace.ReplaceString = '2026년  03월  24일';
  hwp.HParameterSet.HFindReplace.IgnoreMessage = 1;
  hwp.HParameterSet.HFindReplace.ReplaceMode = 1;
  hwp.HAction.Execute('AllReplace', hwp.HParameterSet.HFindReplace.HSet);
  console.log('  ✅ 날짜');

  console.log('\n=== 이력서 입력 완료! ===');

  await ask('── 확인 후 종료');
  bridge.comCallWith(h, 'Clear', [1]);
  bridge.comCallWith(h, 'Quit', []);
  rl.close();
  process.exit(0);

})();
