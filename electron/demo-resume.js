/**
 * 시연용: 이력서 자동 채우기 + 사진 삽입
 *
 * 사용법: cd electron && node demo-resume.js
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

var rl = readline.createInterface({ input: process.stdin, output: process.stdout });
function ask(msg) {
  return new Promise(function(resolve) {
    rl.question('\n' + msg + ' → Enter ', function() { resolve(); });
  });
}

// ════════════════════════════════════════
// 유틸
// ════════════════════════════════════════

function escapeXml(s) {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

function fill(xml, tableIdx, rowAddr, colAddr, value) {
  var tableStart = xml.indexOf('<TABLE');
  for (var t = 0; t < tableIdx; t++) {
    tableStart = xml.indexOf('<TABLE', xml.indexOf('</TABLE>', tableStart) + 1);
  }
  var tableEnd = xml.indexOf('</TABLE>', tableStart) + 8;
  var tableXml = xml.substring(tableStart, tableEnd);

  var searchStart = 0;
  while (true) {
    var cellStart = tableXml.indexOf('<CELL', searchStart);
    if (cellStart < 0) break;
    var cellTagEnd = tableXml.indexOf('>', cellStart);
    var cellTag = tableXml.substring(cellStart, cellTagEnd + 1);

    if (cellTag.indexOf('RowAddr="' + rowAddr + '"') >= 0 &&
        cellTag.indexOf('ColAddr="' + colAddr + '"') >= 0) {
      var cellEnd = tableXml.indexOf('</CELL>', cellStart) + 7;
      var cellContent = tableXml.substring(cellStart, cellEnd);
      var modified = cellContent;
      var val = escapeXml(value);

      modified = modified.replace(/<TEXT CharShape="(\d+)"\/>/,
        '<TEXT CharShape="$1"><CHAR>' + val + '</CHAR></TEXT>');
      modified = modified.replace(/<TEXT CharShape="(\d+)"><\/TEXT>/,
        '<TEXT CharShape="$1"><CHAR>' + val + '</CHAR></TEXT>');

      var globalStart = tableStart + cellStart;
      var globalEnd = tableStart + cellEnd;
      xml = xml.substring(0, globalStart) + modified + xml.substring(globalEnd);
      break;
    }
    searchStart = cellStart + 1;
  }
  return xml;
}

// ════════════════════════════════════════
// 실행
// ════════════════════════════════════════

(async function() {

  // ── 시작: 문서 열기 ──
  console.log('이력서 자동 채우기 시연\n');

  await ask('── 문서 열기');
  var filePath = 'C:/Users/seng1/Documents/Projects/com-test/data/이력서_자기소개서_기본.hwp';
  bridge.comCallWith(h, 'Open', [filePath, 'HWP', '']);
  console.log('  문서 열림');

  // ── Enter 1: 1페이지 텍스트 전부 ──
  await ask('── 1페이지 텍스트 채우기');

  var xml = String(bridge.comCallWith(h, 'GetTextFile', ['HWPML2X', '']));

  // 사진 플레이스홀더
  xml = fill(xml, 0, 1, 0, '{{PHOTO}}');

  // 기본 정보
  xml = fill(xml, 0, 1, 3, '홍길동');
  xml = fill(xml, 0, 1, 6, '1990. 01. 15');
  xml = fill(xml, 0, 2, 3, '010-1234-5678');
  xml = fill(xml, 0, 2, 6, '010-9876-5432');
  xml = fill(xml, 0, 3, 3, 'hong@example.com');
  xml = fill(xml, 0, 4, 3, '서울시 강남구 테헤란로 123 비즈타워 5층');

  // 학력
  xml = fill(xml, 0, 6, 1, '2009.03 ~ 2013.02');
  xml = fill(xml, 0, 6, 3, '서울대학교 (컴퓨터공학과)');
  xml = fill(xml, 0, 7, 1, '2013.03 ~ 2015.02');
  xml = fill(xml, 0, 7, 3, '서울대학교 대학원 (인공지능 전공)');

  // 경력
  xml = fill(xml, 0, 9, 1, '2015.03 ~ 2018.12');
  xml = fill(xml, 0, 9, 3, '네이버');
  xml = fill(xml, 0, 9, 4, '검색플랫폼팀 / 백엔드 개발');
  xml = fill(xml, 0, 10, 1, '2019.01 ~ 2022.06');
  xml = fill(xml, 0, 10, 3, '카카오');
  xml = fill(xml, 0, 10, 4, 'AI Lab / ML 엔지니어');
  xml = fill(xml, 0, 11, 1, '2022.07 ~ 현재');
  xml = fill(xml, 0, 11, 3, '토스');
  xml = fill(xml, 0, 11, 4, 'AI팀 / 테크 리드');

  // 교육
  xml = fill(xml, 0, 13, 1, '2020.06 ~ 2020.08');
  xml = fill(xml, 0, 13, 3, 'AWS Solutions Architect');
  xml = fill(xml, 0, 13, 4, 'Amazon Web Services');
  xml = fill(xml, 0, 13, 7, '클라우드 아키텍처 설계');

  // 자격
  xml = fill(xml, 0, 15, 1, '2014. 11');
  xml = fill(xml, 0, 15, 3, '정보처리기사');
  xml = fill(xml, 0, 15, 6, '한국산업인력공단');
  xml = fill(xml, 0, 16, 1, '2020. 09');
  xml = fill(xml, 0, 16, 3, 'AWS Solutions Architect Associate');
  xml = fill(xml, 0, 16, 6, 'Amazon Web Services');

  // 업무능력
  xml = fill(xml, 0, 18, 1, 'Python, Java, Go, TypeScript, React, Node.js, Docker, Kubernetes, AWS, GCP, PostgreSQL, Redis, Kafka, TensorFlow, PyTorch');

  // 1페이지만 먼저 반영
  bridge.comCallWith(h, 'Run', ['SelectAll']);
  bridge.comCallWith(h, 'Run', ['Delete']);
  bridge.comCallWith(h, 'SetTextFile', [xml, 'HWPML2X', '']);

  // 날짜
  bridge.comCallWith(h, 'Run', ['MoveDocBegin']);
  hwp.HAction.GetDefault('AllReplace', hwp.HParameterSet.HFindReplace.HSet);
  hwp.HParameterSet.HFindReplace.FindString = '년    월    일';
  hwp.HParameterSet.HFindReplace.ReplaceString = '2026년  03월  24일';
  hwp.HParameterSet.HFindReplace.IgnoreMessage = 1;
  hwp.HParameterSet.HFindReplace.ReplaceMode = 1;
  hwp.HAction.Execute('AllReplace', hwp.HParameterSet.HFindReplace.HSet);

  console.log('  ✅ 1페이지 텍스트 반영됨');

  // 사진 삽입

  bridge.comCallWith(h, 'Run', ['MoveDocBegin']);
  hwp.HAction.GetDefault('RepeatFind', hwp.HParameterSet.HFindReplace.HSet);
  hwp.HParameterSet.HFindReplace.FindString = '{{PHOTO}}';
  hwp.HParameterSet.HFindReplace.IgnoreMessage = 1;
  hwp.HAction.Execute('RepeatFind', hwp.HParameterSet.HFindReplace.HSet);
  bridge.comCallWith(h, 'Run', ['Delete']);

  var imgPath = 'C:/Users/seng1/Documents/Projects/com-test/data/예시이미지.jpg';
  var ctrl = bridge.comCallWith(h, 'InsertPicture', [imgPath, 1, 0, 0, 0, 0, 0, 0]);
  if (ctrl && typeof ctrl === 'object') {
    var props = bridge.comGet(ctrl, 'Properties');
    bridge.comCallWith(props, 'SetItem', ['Width', 8041]);
    bridge.comCallWith(props, 'SetItem', ['Height', 10440]);
    bridge.comPut(ctrl, 'Properties', props);
  }

  console.log('  ✅ 사진 삽입 + 크기 조절 완료');

  // 2페이지 텍스트 (자기소개서)

  // 현재 문서에서 다시 HWPML2X 추출 (사진 포함된 상태)
  var xml2 = String(bridge.comCallWith(h, 'GetTextFile', ['HWPML2X', '']));

  xml2 = fill(xml2, 1, 0, 1, '어릴 때부터 컴퓨터에 남다른 호기심을 가지고 자랐습니다. 초등학교 때 처음 접한 HTML로 개인 홈페이지를 만들면서 프로그래밍의 즐거움을 알게 되었고, 중학교 때는 C언어를 독학하며 알고리즘의 세계에 빠졌습니다. 대학에서 컴퓨터공학을 전공하며 이론적 기반을 탄탄히 다졌고, 대학원에서는 인공지능 분야를 깊이 연구했습니다.');
  xml2 = fill(xml2, 1, 1, 1, '논리적이고 분석적인 사고방식이 강점입니다. 복잡한 문제를 작은 단위로 분해하여 체계적으로 해결하는 데 능숙합니다. 코드 리뷰에서 동료들이 놓치기 쉬운 엣지 케이스를 잘 찾아내는 편입니다. 반면 완벽주의 성향이 있어 때로는 의사결정에 시간이 걸리는 단점이 있으나, 최근에는 MVP 방법론을 적극 활용하며 개선하고 있습니다.');
  xml2 = fill(xml2, 1, 2, 1, '대학 시절 IT 봉사동아리에서 3년간 활동하며 지역 아동센터에서 코딩 교육을 진행했습니다. 오픈소스 커뮤니티에서 다수의 프로젝트에 기여했으며, 특히 한국어 NLP 라이브러리의 주요 컨트리뷰터로 활동했습니다.');
  xml2 = fill(xml2, 1, 3, 1, 'AI 기술이 사람들의 일상을 더 편리하고 풍요롭게 만들 수 있다고 확신합니다. 귀사가 추구하는 비전에 깊이 공감하며, 지금까지 쌓아온 백엔드 개발과 ML 엔지니어링 역량을 바탕으로 귀사의 AI 서비스 고도화에 기여하고 싶습니다.');

  bridge.comCallWith(h, 'Run', ['SelectAll']);
  bridge.comCallWith(h, 'Run', ['Delete']);
  bridge.comCallWith(h, 'SetTextFile', [xml2, 'HWPML2X', '']);

  console.log('  ✅ 2페이지 자기소개서 반영됨');
  console.log('\n✅ 이력서 자동 채우기 완료!');

  await ask('── 종료');
  bridge.comCallWith(h, 'Clear', [1]);
  bridge.comCallWith(h, 'Quit', []);
  rl.close();
  process.exit(0);

})();
