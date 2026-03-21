const codeEl = document.getElementById('code');
const outputEl = document.getElementById('output');
const statusEl = document.getElementById('status');

let reqId = 0;

// Worker 응답 처리
window.comBridge.onResponse((msg) => {
  if (msg.type === 'status') {
    const parts = [];
    if (msg.excel) parts.push('Excel');
    if (msg.hwp) parts.push('HWP');
    statusEl.textContent = parts.length ? parts.join(' + ') + ' running' : 'Ready';

    document.getElementById('btn-excel').classList.toggle('active', msg.excel);
    document.getElementById('btn-hwp').classList.toggle('active', msg.hwp);
    return;
  }

  if (msg.type === 'result') {
    outputEl.className = 'success';
    const lines = [];
    if (msg.logs && msg.logs.length) lines.push('--- logs ---\n' + msg.logs.join('\n'));
    if (msg.result !== null && msg.result !== undefined) {
      lines.push('--- result ---\n' + JSON.stringify(msg.result, null, 2));
    } else {
      lines.push('--- OK (no return value) ---');
    }
    outputEl.value = lines.join('\n\n');
    return;
  }

  if (msg.type === 'error') {
    outputEl.className = 'error';
    const lines = ['Error: ' + msg.error];
    if (msg.line) lines.push('Line: ' + msg.line);
    if (msg.restored) lines.push('[Document restored from savepoint]');
    if (msg.logs && msg.logs.length) lines.push('\n--- logs ---\n' + msg.logs.join('\n'));
    if (msg.stack) lines.push('\n--- stack ---\n' + msg.stack);
    outputEl.value = lines.join('\n');
    return;
  }

  if (msg.type === 'members') {
    outputEl.className = '';
    outputEl.value = msg.members.map(function(m) {
      return '[' + m.kind + '] ' + m.name + '(' +
        m.params.map(function(p) { return p.name + ': ' + p.vt; }).join(', ') + ')' +
        (m.returnType !== 'void' ? ' → ' + m.returnType : '');
    }).join('\n');
  }
});

// Execute
document.getElementById('btn-run').addEventListener('click', function() {
  var code = codeEl.value.trim();
  if (!code) return;
  outputEl.value = 'Running...';
  outputEl.className = '';
  window.comBridge.send({ type: 'execute', id: String(++reqId), code: code });
});

// Copy output
document.getElementById('btn-copy').addEventListener('click', function() {
  outputEl.select();
  document.execCommand('copy');
  statusEl.textContent = 'Copied!';
  setTimeout(function() { statusEl.textContent = ''; }, 1500);
});

// Launch / Quit
document.getElementById('btn-excel').addEventListener('click', function() {
  window.comBridge.send({ type: 'launch', id: String(++reqId), app: 'excel' });
});
document.getElementById('btn-hwp').addEventListener('click', function() {
  window.comBridge.send({ type: 'launch', id: String(++reqId), app: 'hwp' });
});
document.getElementById('btn-quit-excel').addEventListener('click', function() {
  window.comBridge.send({ type: 'quit', id: String(++reqId), app: 'excel' });
});
document.getElementById('btn-quit-hwp').addEventListener('click', function() {
  window.comBridge.send({ type: 'quit', id: String(++reqId), app: 'hwp' });
});

// Ctrl+Enter로 실행
codeEl.addEventListener('keydown', function(e) {
  if (e.ctrlKey && e.key === 'Enter') {
    document.getElementById('btn-run').click();
  }
});
