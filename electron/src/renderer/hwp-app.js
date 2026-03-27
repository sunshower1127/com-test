var comCode = document.getElementById('com-code');
var comOutput = document.getElementById('com-output');
var xmlInput = document.getElementById('xml-input');
var xmlOutput = document.getElementById('xml-output');
var statusEl = document.getElementById('status');
var launchBtn = document.getElementById('btn-launch');

var reqId = 0;
var hwpRunning = false;
var pendingXmlCallback = null;

// --- Worker 응답 처리 ---
window.comBridge.onResponse(function(msg) {
  if (msg.type === 'status') {
    hwpRunning = !!msg.hwp;
    statusEl.textContent = hwpRunning ? 'HWP running' : 'HWP not running';
    launchBtn.classList.toggle('active', hwpRunning);
    launchBtn.textContent = hwpRunning ? 'HWP Running' : 'Launch HWP';
    return;
  }

  // XML 패널 응답
  if (msg.id && msg.id.startsWith('xml-')) {
    if (msg.type === 'result') {
      xmlOutput.className = 'xml-output output success';
      var lines = [];
      if (msg.logs && msg.logs.length) lines.push(msg.logs.join('\n'));
      if (msg.result !== null && msg.result !== undefined) {
        lines.push(typeof msg.result === 'string' ? msg.result : JSON.stringify(msg.result, null, 2));
      } else {
        lines.push('OK');
      }
      xmlOutput.value = lines.join('\n\n');
    } else if (msg.type === 'error') {
      xmlOutput.className = 'xml-output output error';
      var lines = ['Error: ' + msg.error];
      if (msg.stack) lines.push('\n' + msg.stack);
      xmlOutput.value = lines.join('\n');
    }
    return;
  }

  // COM 패널 응답
  if (msg.type === 'result') {
    comOutput.className = 'output success';
    var lines = [];
    if (msg.logs && msg.logs.length) lines.push(msg.logs.join('\n'));
    if (msg.result !== null && msg.result !== undefined) {
      var r = msg.result;
      if (typeof r === 'string') {
        lines.push(r);
      } else {
        lines.push(JSON.stringify(r, null, 2));
      }
    } else {
      lines.push('OK');
    }
    comOutput.value = lines.join('\n\n');
    return;
  }

  if (msg.type === 'error') {
    comOutput.className = 'output error';
    var lines = ['Error: ' + msg.error];
    if (msg.line) lines.push('Line: ' + msg.line);
    if (msg.restored) lines.push('[Document restored from savepoint]');
    if (msg.logs && msg.logs.length) lines.push('\n--- logs ---\n' + msg.logs.join('\n'));
    if (msg.stack) lines.push('\n--- stack ---\n' + msg.stack);
    comOutput.value = lines.join('\n');
    return;
  }
});

// --- Launch / Quit ---
launchBtn.addEventListener('click', function() {
  if (!hwpRunning) {
    window.comBridge.send({ type: 'launch', id: String(++reqId), app: 'hwp' });
  }
});

document.getElementById('btn-quit').addEventListener('click', function() {
  window.comBridge.send({ type: 'quit', id: String(++reqId), app: 'hwp' });
});

// --- Reset ---
document.getElementById('btn-reset').addEventListener('click', function() {
  comCode.value = '';
  comOutput.value = '';
  comOutput.className = 'output';
  xmlInput.value = '';
  xmlOutput.value = '';
  xmlOutput.className = 'xml-output output';
  if (hwpRunning) {
    window.comBridge.send({ type: 'quit', id: String(++reqId), app: 'hwp' });
    setTimeout(function() {
      window.comBridge.send({ type: 'launch', id: String(++reqId), app: 'hwp' });
    }, 1000);
  }
});

// === COM 패널 ===
document.getElementById('btn-run-com').addEventListener('click', function() {
  var code = comCode.value.trim();
  if (!code) return;
  comOutput.value = 'Running...';
  comOutput.className = 'output';
  window.comBridge.send({ type: 'execute', id: 'com-' + (++reqId), code: code });
});

document.getElementById('btn-copy-com').addEventListener('click', function() {
  comOutput.select();
  document.execCommand('copy');
  statusEl.textContent = 'Copied!';
  setTimeout(function() { statusEl.textContent = hwpRunning ? 'HWP running' : 'HWP not running'; }, 1500);
});

document.getElementById('btn-clear-com').addEventListener('click', function() {
  comCode.value = '';
  comOutput.value = '';
  comOutput.className = 'output';
});

comCode.addEventListener('keydown', function(e) {
  if (e.ctrlKey && e.key === 'Enter') {
    document.getElementById('btn-run-com').click();
  }
});

// === XML 패널 ===

function sendXmlCmd(code) {
  xmlOutput.value = 'Running...';
  xmlOutput.className = 'xml-output output';
  window.comBridge.send({ type: 'execute', id: 'xml-' + (++reqId), code: code });
}

// STRUCTURE 버튼
document.getElementById('btn-xml-structure').addEventListener('click', function() {
  sendXmlCmd('xml.load(); result = xml.structure()');
});

// GET 버튼 — 입력창에서 경로 파싱
document.getElementById('btn-xml-get').addEventListener('click', function() {
  var text = xmlInput.value.trim();
  // "GET t0.r1.c3" 또는 "t0.r1.c3" 또는 "GET p2" 등
  var path = text.replace(/^GET\s+/i, '').trim().split('\n')[0].trim();
  if (!path) {
    xmlOutput.value = 'Error: 경로를 입력하세요. 예: t0.r1.c3 또는 p2';
    xmlOutput.className = 'xml-output output error';
    return;
  }
  sendXmlCmd('xml.load(); result = xml.raw("' + path.replace(/"/g, '\\"') + '")');
});

// SET 버튼 — 첫 줄에서 경로, 나머지가 XML
document.getElementById('btn-xml-set').addEventListener('click', function() {
  var text = xmlInput.value.trim();
  var lines = text.split('\n');
  var firstLine = lines[0].replace(/^SET\s+/i, '').trim();
  var path = firstLine;
  var xmlContent = lines.slice(1).join('\n').trim();

  if (!path) {
    xmlOutput.value = 'Error: 첫 줄에 경로를 입력하세요. 예:\nSET t0.r1.c3\n<CELL>...</CELL>';
    xmlOutput.className = 'xml-output output error';
    return;
  }
  if (!xmlContent) {
    xmlOutput.value = 'Error: XML 내용이 없습니다.';
    xmlOutput.className = 'xml-output output error';
    return;
  }

  // XML을 이스케이프해서 코드로 전달
  var escaped = xmlContent
    .replace(/\\/g, '\\\\')
    .replace(/"/g, '\\"')
    .replace(/\n/g, '\\n')
    .replace(/\r/g, '\\r');
  sendXmlCmd('xml.load(); result = xml.rawSet("' + path.replace(/"/g, '\\"') + '", "' + escaped + '")');
});

// Clear
document.getElementById('btn-xml-clear').addEventListener('click', function() {
  xmlInput.value = '';
  xmlOutput.value = '';
  xmlOutput.className = 'xml-output output';
});

// XML 입력창 Ctrl+Enter → 자동 판별 (GET/SET/STRUCTURE)
xmlInput.addEventListener('keydown', function(e) {
  if (e.ctrlKey && e.key === 'Enter') {
    var text = xmlInput.value.trim();
    if (!text || text.toUpperCase() === 'STRUCTURE') {
      document.getElementById('btn-xml-structure').click();
    } else if (text.toUpperCase().startsWith('SET')) {
      document.getElementById('btn-xml-set').click();
    } else {
      // 기본은 GET
      document.getElementById('btn-xml-get').click();
    }
  }
});
