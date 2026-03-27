/**
 * HWPML2X 라운드트립: XML에서 텍스트 수정 → SetTextFile로 넣기
 */

var bridge = require('./native/com_bridge_node.node');
var { initBridge, createComProxy } = require('./dist/worker/proxy');
var fs = require('fs');

bridge.comInit();
initBridge(bridge);

var h = bridge.comCreate('HWPFrame.HwpObject');
var hwp = createComProxy(h);
var win0 = bridge.comCallWith(bridge.comGet(h, 'XHwpWindows'), 'Item', [0]);
bridge.comPut(win0, 'Visible', true);

// 이력서 열기
var filePath = 'C:/Users/seng1/Documents/Projects/com-test/data/이력서_자기소개서_기본.hwp';
bridge.comCallWith(h, 'Open', [filePath, 'HWP', '']);
console.log('이력서 열림\n');

// 1. HWPML2X 추출
var xml = String(bridge.comCallWith(h, 'GetTextFile', ['HWPML2X', '']));
console.log('원본 XML 길이: ' + xml.length);

// 2. 구조 확인 — 성명 셀 찾기
// 행1의 구조: ... <CHAR> 성       명</CHAR> ... <TEXT CharShape="10"/> ...
// "성       명" 다음에 오는 빈 <TEXT .../> 가 이름 입력칸

// 성명 빈칸 채우기
var modified = xml;

// 패턴: "성       명" 뒤의 빈 TEXT 태그
// <CHAR> 성       명\n</CHAR> ... </CELL> <CELL ...> ... <TEXT CharShape="10"/>
// 빈 TEXT 태그를 찾아서 내용 넣기

// 간단하게: 성명 CELL 다음 CELL의 빈 TEXT를 찾기
var nameLabel = modified.indexOf('성       명');
if (nameLabel >= 0) {
  // 그 다음 CELL의 빈 <TEXT ... /> 찾기
  var nextCell = modified.indexOf('<CELL', nameLabel);
  var emptyText = modified.indexOf('<TEXT CharShape="10"/>', nextCell);
  if (emptyText >= 0 && emptyText < nextCell + 500) {
    modified = modified.substring(0, emptyText) +
      '<TEXT CharShape="10"><CHAR>홍길동</CHAR></TEXT>' +
      modified.substring(emptyText + '<TEXT CharShape="10"/>'.length);
    console.log('✅ 성명: "홍길동" 삽입');
  }
}

// 생년월일 빈칸
var birthLabel = modified.indexOf('생 년 월 일');
if (birthLabel >= 0) {
  var nextCell = modified.indexOf('<CELL', birthLabel);
  // 여러 CharShape 가능 — 빈 TEXT 찾기
  var searchStart = nextCell;
  var emptyText = modified.indexOf('<TEXT CharShape="3"/>', searchStart);
  if (emptyText < 0) emptyText = modified.indexOf('<TEXT CharShape="10"/>', searchStart);
  if (emptyText >= 0 && emptyText < searchStart + 500) {
    modified = modified.substring(0, emptyText) +
      '<TEXT CharShape="3"><CHAR>1990. 01. 15</CHAR></TEXT>' +
      modified.substring(emptyText + modified.substring(emptyText).match(/<TEXT[^/]*\/>/)[0].length);
    console.log('✅ 생년월일: "1990. 01. 15" 삽입');
  }
}

// 연락처 빈칸
var contactLabel = modified.indexOf('연락처');
if (contactLabel >= 0) {
  var nextCell = modified.indexOf('<CELL', contactLabel);
  var searchArea = modified.substring(nextCell, nextCell + 500);
  var emptyMatch = searchArea.match(/<TEXT CharShape="\d+"\/>/);
  if (emptyMatch) {
    var pos = nextCell + searchArea.indexOf(emptyMatch[0]);
    modified = modified.substring(0, pos) +
      emptyMatch[0].replace('/>', '><CHAR>010-1234-5678</CHAR></TEXT>').replace('</TEXT></TEXT>', '</TEXT>') +
      modified.substring(pos + emptyMatch[0].length);
    // 위에서 </TEXT> 중복 방지
    modified = modified.replace(/<CHAR>010-1234-5678<\/CHAR><\/TEXT><\/TEXT>/, '<CHAR>010-1234-5678</CHAR></TEXT>');
    console.log('✅ 연락처: "010-1234-5678" 삽입');
  }
}

// E-Mail 빈칸
var emailLabel = modified.indexOf('E - Mail');
if (emailLabel >= 0) {
  var nextCell = modified.indexOf('<CELL', emailLabel);
  var searchArea = modified.substring(nextCell, nextCell + 500);
  var emptyMatch = searchArea.match(/<TEXT CharShape="\d+"\/>/);
  if (emptyMatch) {
    var pos = nextCell + searchArea.indexOf(emptyMatch[0]);
    var replacement = emptyMatch[0].replace('/>', '><CHAR>hong@example.com</CHAR></TEXT>').replace('</TEXT></TEXT>', '</TEXT>');
    modified = modified.substring(0, pos) + replacement + modified.substring(pos + emptyMatch[0].length);
    console.log('✅ E-Mail: "hong@example.com" 삽입');
  }
}

// 자기소개서 — 성장과정
var growthLabel = modified.indexOf('성\n</CHAR>');
if (growthLabel < 0) growthLabel = modified.indexOf('성</CHAR>');
// 표2의 성장과정 행 옆 빈칸 찾기
var tbl2Start = modified.indexOf('<TABLE', modified.indexOf('</TABLE>') + 1);
if (tbl2Start >= 0) {
  var firstRow = modified.indexOf('<ROW>', tbl2Start);
  var dataCellStart = modified.indexOf('<CELL', modified.indexOf('<CELL', firstRow) + 1); // 두번째 CELL
  if (dataCellStart >= 0) {
    var searchArea = modified.substring(dataCellStart, dataCellStart + 500);
    var emptyMatch = searchArea.match(/<TEXT CharShape="\d+"\/>/);
    if (emptyMatch) {
      var pos = dataCellStart + searchArea.indexOf(emptyMatch[0]);
      var replacement = emptyMatch[0].replace('/>', '><CHAR>어릴 때부터 컴퓨터에 관심이 많았습니다.</CHAR></TEXT>').replace('</TEXT></TEXT>', '</TEXT>');
      modified = modified.substring(0, pos) + replacement + modified.substring(pos + emptyMatch[0].length);
      console.log('✅ 성장과정 삽입');
    }
  }
}

console.log('\n수정 XML 길이: ' + modified.length + ' (원본: ' + xml.length + ')');

// 3. 수정된 XML을 SetTextFile로 넣기
console.log('\n=== SetTextFile("HWPML2X") ===');
bridge.comCallWith(h, 'Run', ['SelectAll']);
bridge.comCallWith(h, 'Run', ['Delete']);
var ret = bridge.comCallWith(h, 'SetTextFile', [modified, 'HWPML2X', '']);
console.log('SetTextFile ret=' + ret);

// 4. 결과 확인
var result = String(bridge.comCallWith(h, 'GetTextFile', ['UNICODE', '']));
console.log('\n전체 텍스트:');
result.split(/\r?\n/).forEach(function(line, i) {
  if (line.trim()) console.log('  [' + i + '] ' + line.trim());
});

console.log('\n확인해주세요! (닫지 않음)');
