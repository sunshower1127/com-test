var comCode = document.getElementById('com-code');
var comOutput = document.getElementById('com-output');
var statusEl = document.getElementById('status');
var launchBtn = document.getElementById('btn-launch');

var reqId = 0;
var hwpRunning = false;

// --- Worker 응답 처리 ---
window.comBridge.onResponse(function(msg) {
  if (msg.type === 'status') {
    hwpRunning = !!msg.hwp;
    statusEl.textContent = hwpRunning ? 'HWP running' : 'HWP not running';
    launchBtn.classList.toggle('active', hwpRunning);
    launchBtn.textContent = hwpRunning ? 'HWP Running' : 'Launch HWP';
    return;
  }

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
  if (hwpRunning) {
    window.comBridge.send({ type: 'quit', id: String(++reqId), app: 'hwp' });
    setTimeout(function() {
      window.comBridge.send({ type: 'launch', id: String(++reqId), app: 'hwp' });
    }, 1000);
  }
});

// --- Run ---
document.getElementById('btn-run-com').addEventListener('click', function() {
  var code = comCode.value.trim();
  if (!code) return;
  comOutput.value = 'Running...';
  comOutput.className = 'output';
  window.comBridge.send({ type: 'execute', id: 'com-' + (++reqId), code: code });
});

// --- Copy ---
document.getElementById('btn-copy-com').addEventListener('click', function() {
  comOutput.select();
  document.execCommand('copy');
  statusEl.textContent = 'Copied!';
  setTimeout(function() { statusEl.textContent = hwpRunning ? 'HWP running' : 'HWP not running'; }, 1500);
});

// --- Clear ---
document.getElementById('btn-clear-com').addEventListener('click', function() {
  comCode.value = '';
  comOutput.value = '';
  comOutput.className = 'output';
});

// --- Ctrl+Enter ---
comCode.addEventListener('keydown', function(e) {
  if (e.ctrlKey && e.key === 'Enter') {
    document.getElementById('btn-run-com').click();
  }
});
