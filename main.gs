/** ===================== MANUAL / AUTO SENDER (PENDING + ON REQUEST) =====================
 * Menu:
 *  - Sent Task → Preview Task - Pending        (view PENDING; used with auto send 07:00)
 *  - Sent Task → Sent Task - On Request       (manual send; includes global ON/OFF for auto daily Pending)
 *  - Sent Task → Test push to group
 * ==================================================================================== */

const CONFIG = {
  HEADER_ROW: 1,
  TIMEZONE: 'Asia/Bangkok',
  SHEET_NAME: '',

  // LINE OA
  OA_CHANNEL_ACCESS_TOKEN: '',
  OA_GROUP_ID: '',

  // Status values
  STATUS_FILTER_VALUE: 'On request', // manual (Sent Task - On Request)
  PENDING_STATUS_VALUE: 'Pending',   // auto daily at 07:00

  // Display
  TITLE: 'On request',
  SEPARATOR_LINE: '------------------',
  LINE_CHAR_LIMIT: 4900,

  HEADERS: {
    date:        'Date',
    requestBy:   'Request By',
    productId:   'Product ID',
    productName: 'Product Name',
    platform:    'Platform',
    changedType: 'Changed Type',
    impact:      'Impact',
    action:      'Action',
    status:      'Status'
  },
  ALT_HEADERS: {
    status:      ['Progress', 'สถานะ'],
    productId:   ['Product Id', 'ProductID', 'รหัสสินค้า'],
    productName: ['Name', 'Tour Name', 'ชื่อสินค้า', 'ชื่อทัวร์'],
    changedType: ['Change Type', 'Type of Change', 'ประเภทการเปลี่ยนแปลง'],
    impact:      ['Impacts'],
    action:      ['Actions', 'รายละเอียดการเปลี่ยนแปลง'],
    date:        ['Acted Date', 'Effective Date', 'วันที่'],
    requestBy:   ['Request by', 'Requested By', 'Requester', 'ผู้ร้องขอ']
  }
};

/** ===================== MENU ===================== */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Sent Task')
    .addItem('▶ Preview Task - Pending', 'openPreviewPending')
    .addItem('▶ Sent Task - On Request', 'openSendPicker')
    .addSeparator()
    .addItem('▶ Test push to group', 'testPushToGroup')
    .addToUi();
}

/** Helpers */
function safeJson_(v){
  return JSON.stringify(v)
    .replace(/<\//g, '<\\/')
    .replace(/\u2028|\u2029/g, '\\n');
}
// join blocks with separator only BETWEEN items (no head/tail lines)
function joinBlocks_(blocks, sep){
  if (!blocks || !blocks.length) return '';
  if (blocks.length === 1) return String(blocks[0]);
  const s = sep || CONFIG.SEPARATOR_LINE;
  return blocks.join('\n' + s + '\n');
}

/** ===================== PREVIEW PENDING (VIEW-ONLY) ===================== */
function openPreviewPending() {
  const rows = getPendingRowsForUi_();
  const head = buildPendingHead_();
  const body = joinBlocks_(rows.map(function(r){ return r.block; }), CONFIG.SEPARATOR_LINE);
  const fullText = body
    ? head + '\n' + CONFIG.SEPARATOR_LINE + '\n' + body
    : head;

  const html = HtmlService.createHtmlOutput(
    [
      '<!doctype html><html><head><meta charset="utf-8"><title>Preview Task - Pending</title>',
      '<style>',
      'body{margin:0;background:#f6f7f9;color:#1f2328;font:13px/1.5 system-ui,-apple-system,Segoe UI,Roboto,Arial,"Noto Sans",sans-serif}',
      '.wrap{padding:14px}.bar{display:flex;gap:8px;align-items:center;flex-wrap:wrap;margin-bottom:10px}.h{font-weight:700}',
      '.pill{padding:2px 8px;border-radius:999px;background:#d1fae5;color:#065f46;font-weight:700;font-size:12px}',
      'button{font:inherit;border:1px solid #e5e7eb;border-radius:10px;padding:7px 10px;background:#fff;cursor:pointer}button:hover{background:#fafafa}',
      'pre{white-space:pre-wrap;border:1px solid #e5e7eb;border-radius:10px;padding:10px;background:#fff;height:68vh;overflow:auto}.hint{color:#64748b;margin-left:auto}',
      '</style></head><body><div class="wrap">',
      '<div class="bar"><span class="h">Preview Task - Pending</span><span class="pill">', String(rows.length),
      '</span><span class="hint">Status = "Pending" • ระบบจะส่งอัตโนมัติทุกวันเวลา 07:00 น. ถ้าเปิด Auto</span>',
      '<button id="copy">Copy All</button><button id="dl">Download .txt</button><button onclick="google.script.host.close()">Close</button></div>',
      '<div class="label">รวมข้อความทั้งหมด</div><pre id="big"></pre>',
      '<script>',
      'var TEXT=', safeJson_(fullText), ';',
      'document.getElementById("big").textContent = TEXT;',
      'document.getElementById("copy").onclick = function(){ navigator.clipboard.writeText(TEXT).then(function(){alert("Copied!");}).catch(function(){alert("Copy failed.");}); };',
      'document.getElementById("dl").onclick = function(){ var blob=new Blob([TEXT],{type:"text/plain"}); var url=URL.createObjectURL(blob); var a=document.createElement("a"); a.href=url; a.download="pending-preview.txt"; document.body.appendChild(a); a.click(); setTimeout(function(){URL.revokeObjectURL(url);a.remove();},800); };',
      '</script></div></body></html>'
    ].join('')
  ).setWidth(980).setHeight(680);

  SpreadsheetApp.getUi().showModalDialog(html, 'Preview Task - Pending');
}

/** ===================== PICK & SEND (ON REQUEST — MANUAL) ===================== */
function openSendPicker() {
  const rows = getOnRequestRowsForUi_();
  const head = buildOnRequestHead_();

  const html = HtmlService.createHtmlOutput(
    [
      '<!doctype html><html><head><meta charset="utf-8"><title>Sent Task - On Request</title>',
      '<style>',
      'body{margin:0;background:#f6f7f9;color:#1f2328;font:13px/1.5 system-ui,-apple-system,Segoe UI,Roboto,Arial,"Noto Sans",sans-serif}',
      '.wrap{padding:14px}.bar{display:flex;gap:8px;align-items:center;flex-wrap:wrap;margin-bottom:10px}.h{font-weight:700}',
      '.pill{padding:2px 8px;border-radius:999px;background:#dbeafe;color:#1e3a8a;font-weight:700;font-size:12px}',
      '.pill-small{padding:2px 6px;font-size:11px}',
      'button{font:inherit;border:1px solid #e5e7eb;border-radius:10px;padding:7px 10px;background:#fff;cursor:pointer}button:hover{background:#fafafa}',
      '.grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(360px,1fr));gap:10px}.card{background:#fff;border:1px solid #e5e7eb;border-radius:14px;padding:10px}',
      '.hdr{display:flex;gap:8px;align-items:center;margin-bottom:6px}.name{font-weight:700;flex:1}',
      '.mono{font:12px ui-monospace,Menlo,Consolas,monospace;white-space:pre-wrap;border:1px solid #e5e7eb;border-radius:10px;padding:8px;background:#fff}',
      '.label{color:#64748b;font-size:12px;margin-top:4px}',
      '</style></head><body><div class="wrap">',
      '<div class="bar"><span class="h">Sent Task - On Request</span><span class="pill">All: ', String(rows.length),
      '</span><span class="pill" id="sel">Selected: 0</span>',
      '<span class="pill pill-small" id="autoState">Auto daily Pending @7:00: OFF</span>',
      '<button id="autoToggle">OFF</button>',
      '<span style="margin-left:auto"></span>',
      '<button id="selAll">Select all</button><button id="clear">Clear</button><button id="send">Send selected</button><button onclick="google.script.host.close()">Close</button></div>',
      '<div id="list" class="grid"></div><div class="label" style="margin-top:10px">รวมข้อความที่จะส่ง</div><pre id="big" class="mono"></pre>',
      '<script>',
      'var HEAD=', safeJson_(head), ';',
      'var SEP=', safeJson_(CONFIG.SEPARATOR_LINE), ';',
      'var DATA=', safeJson_(rows), ';',
      'var picked = {};',
      'var AUTO_ENABLED = false;',
      'function esc(s){',
      '  s = String(s == null ? "" : s);',
      '  s = s.replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;");',
      '  return s;',
      '}',
      'function render(){',
      '  var box = document.getElementById("list");',
      '  var html = "";',
      '  for (var i=0;i<DATA.length;i++){',
      '    var r = DATA[i];',
      '    var checked = picked[r.row] ? " checked" : "";',
      '    html += "<div class=\\"card\\">";',
      '    html += "<div class=\\"hdr\\">";',
      '    html += "<input type=\\"checkbox\\" class=\\"chk\\" data-row=\\"" + r.row + "\\"" + checked + ">";',
      '    html += "<div class=\\"name\\">" + esc(r.productName||"-") + "</div>";',
      '    html += "<a target=\\"_blank\\" href=\\"" + esc(r.openUrl) + "\\">#" + r.row + "</a>";',
      '    html += "</div>";',
      '    html += "<div class=\\"label\\">Action</div>";',
      '    html += "<div class=\\"mono\\">" + esc(r.action||"-") + "</div>";',
      '    html += "</div>";',
      '  }',
      '  box.innerHTML = html;',
      '  var chks = document.querySelectorAll(".chk");',
      '  for (var j=0;j<chks.length;j++){',
      '    chks[j].onchange = function(){',
      '      var row = Number(this.getAttribute("data-row"));',
      '      if (this.checked){ picked[row] = true; } else { delete picked[row]; }',
      '      updateBig();',
      '    };',
      '  }',
      '  updateBig();',
      '}',
      'function updateBig(){',
      '  var sel = document.getElementById("sel");',
      '  var big = document.getElementById("big");',
      '  var selectedRows = [];',
      '  for (var i=0;i<DATA.length;i++){',
      '    var r = DATA[i];',
      '    if (picked[r.row]) selectedRows.push(r);',
      '  }',
      '  sel.textContent = "Selected: " + selectedRows.length;',
      '  var blocks = [];',
      '  for (var k=0;k<selectedRows.length;k++){ blocks.push(selectedRows[k].block); }',
      '  var text = "";',
      '  if (blocks.length === 1){ text = blocks[0] || ""; }',
      '  else if (blocks.length > 1){ text = blocks.join("\\n"+SEP+"\\n"); }',
      '  if (blocks.length > 0){',
      '    text = HEAD + "\\n" + SEP + "\\n" + text;',
      '  }',
      '  big.textContent = text;',
      '}',
      'function updateAutoButton(){',
      '  var btn = document.getElementById("autoToggle");',
      '  var label = document.getElementById("autoState");',
      '  if (!btn || !label) return;',
      '  btn.textContent = AUTO_ENABLED ? "ON" : "OFF";',
      '  label.textContent = "Auto daily Pending @7:00: " + (AUTO_ENABLED ? "ON" : "OFF");',
      '}',
      'document.getElementById("selAll").onclick = function(){',
      '  picked = {};',
      '  for (var i=0;i<DATA.length;i++){ picked[DATA[i].row] = true; }',
      '  render();',
      '};',
      'document.getElementById("clear").onclick = function(){',
      '  picked = {};',
      '  render();',
      '};',
      'document.getElementById("send").onclick = function(){',
      '  try {',
      '    var btn = this;',
      '    var selectedRows = [];',
      '    for (var i=0;i<DATA.length;i++){ if (picked[DATA[i].row]) selectedRows.push(DATA[i]); }',
      '    if (!selectedRows.length){ alert("ยังไม่ได้เลือกแถวที่จะส่ง"); return; }',
      '    btn.disabled = true; var old = btn.textContent; btn.textContent = "Sending...";',
      '    google.script.run',
      '      .withSuccessHandler(function(msg){ btn.disabled=false; btn.textContent=old; alert(msg||"Sent!"); })',
      '      .withFailureHandler(function(err){ btn.disabled=false; btn.textContent=old; alert("Send failed:\\n" + (err && err.message ? err.message : err)); })',
      '      .sendSelectedOnRequest(selectedRows);',
      '  } catch(ex) { alert("Client error: " + ex.message); }',
      '};',
      'document.getElementById("autoToggle").onclick = function(){',
      '  var btn = this;',
      '  var next = !AUTO_ENABLED;',
      '  btn.disabled = true;',
      '  google.script.run',
      '    .withSuccessHandler(function(){ AUTO_ENABLED = next; btn.disabled=false; updateAutoButton(); })',
      '    .withFailureHandler(function(err){ btn.disabled=false; alert("Update failed: " + (err && err.message ? err.message : err)); })',
      '    .setAutoDailyEnabled(next);',
      '};',
      'google.script.run.withSuccessHandler(function(v){ AUTO_ENABLED = !!v; updateAutoButton(); }).getAutoDailyEnabled();',
      'render();',
      '</script>',
      '</div></body></html>'
    ].join('')
  ).setWidth(980).setHeight(680);

  SpreadsheetApp.getUi().showModalDialog(html, 'Sent Task - On Request');
}

/** ===================== SERVER: PICKED SEND HANDLER (ON REQUEST) ===================== */
function sendSelectedOnRequest(pickedRows) {
  if (!pickedRows || !pickedRows.length) return 'No rows selected';
  const head = buildOnRequestHead_();
  const body = joinBlocks_(pickedRows.map(function(r){ return r.block; }), CONFIG.SEPARATOR_LINE);
  const text = body
    ? head + '\n' + CONFIG.SEPARATOR_LINE + '\n' + body
    : head;
  const chunks = splitAndSendToLine_(text, CONFIG.LINE_CHAR_LIMIT);
  return 'Sent ' + pickedRows.length + ' item(s) to LINE in ' + chunks + ' message chunk(s).';
}

/** ===================== AUTO-SEND SETTINGS (GLOBAL, PENDING ONLY) ===================== */
function getAutoDailyEnabled() {
  const val = PropertiesService.getScriptProperties().getProperty('AUTO_SEND_PENDING');
  return val === 'true';
}

function setAutoDailyEnabled(enabled) {
  PropertiesService.getScriptProperties().setProperty('AUTO_SEND_PENDING', enabled ? 'true' : 'false');
}

/** ===================== AUTO-SEND DAILY (07:00) =====================
 * Attach a time-based trigger to autoSendDaily() at 7:00 AM every day.
 * - If AUTO_SEND_PENDING = true → sends all rows with Status = "Pending"
 * - If AUTO_SEND_PENDING = false → sends nothing
 * - ON REQUEST rows are NEVER auto-sent (manual only)
 * =================================================================== */
function autoSendDaily() {
  const enabled = getAutoDailyEnabled();
  if (!enabled) {
    return { pendingChunks: 0, autoEnabled: false };
  }
  const rows = getPendingRowsForUi_();
  if (!rows.length) {
    return { pendingChunks: 0, autoEnabled: true };
  }
  const head = buildPendingHead_();
  const body = joinBlocks_(rows.map(function(r){ return r.block; }), CONFIG.SEPARATOR_LINE);
  const text = head + '\n' + CONFIG.SEPARATOR_LINE + '\n' + body;
  const chunks = splitAndSendToLine_(text, CONFIG.LINE_CHAR_LIMIT);
  return { pendingChunks: chunks, autoEnabled: true };
}

/** ===================== DATA LAYER (USED BY UIs & AUTO) ===================== */
function getOnRequestRowsForUi_() {
  return getRowsByStatusForUi_(CONFIG.STATUS_FILTER_VALUE || 'On request');
}
function getPendingRowsForUi_() {
  return getRowsByStatusForUi_(CONFIG.PENDING_STATUS_VALUE || 'Pending');
}

function getRowsByStatusForUi_(statusText) {
  const sh = getTargetSheet_();
  const hr = CONFIG.HEADER_ROW || 1;
  const lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
  if (lastRow <= hr) return [];

  const header = sh.getRange(hr, 1, 1, lastCol).getDisplayValues()[0];
  const idx = robustIndexMap_(header);
  const start = hr + 1;
  const rows  = sh.getRange(start, 1, lastRow - start + 1, lastCol).getDisplayValues();

  const base = SpreadsheetApp.getActive().getUrl().split('#')[0];
  const gid  = sh.getSheetId();
  const tz   = CONFIG.TIMEZONE || Session.getScriptTimeZone() || 'Asia/Bangkok';

  const out = [];
  rows.forEach(function(r, off) {
    const statusCell = getCell_(r, idx.status);
    if (!hasStatus_(statusCell, statusText)) return;

    const rownum = start + off;
    const obj = {
      row: rownum,
      date: normalizeDateDisp_(getCell_(r, idx.date), tz),
      requestBy: (getCell_(r, idx.requestBy) || '').trim(),
      productId: (getCell_(r, idx.productId) || '').trim(),
      productName: (getCell_(r, idx.productName) || '').trim(),
      platform: (getCell_(r, idx.platform) || '').trim(),
      changedType: (getCell_(r, idx.changedType) || '').trim(),
      impact: (getCell_(r, idx.impact) || '').trim(),
      action: (getCell_(r, idx.action) || '') + ''
    };

    obj.block = [
      'Date: ' + (obj.date || '-'),
      'Product ID: ' + (obj.productId || '-'),
      'Product Name: ' + (obj.productName || '-'),
      'Platform: ' + (obj.platform || '-'),
      'Changed Type: ' + (obj.changedType || '-'),
      'Impact: ' + (obj.impact || '-'),
      formatActionForBlock_(obj.action),
      'Requested By: ' + (obj.requestBy || '-')
    ].join('\n');

    obj.openUrl = base + '#gid=' + gid + '&range=A' + rownum + ':A' + rownum;
    out.push(obj);
  });
  return out;
}

/** ===================== HEADER MAPPING (ROBUST) ===================== */
function robustIndexMap_(headerRow) {
  const normRow = headerRow.map(normHeader_);
  function findIndex(key) {
    const main = CONFIG.HEADERS[key];
    const alts = (CONFIG.ALT_HEADERS[key] || []);
    const candidates = [main].concat(alts).map(normHeader_);
    for (var i = 0; i < normRow.length; i++) {
      if (candidates.indexOf(normRow[i]) !== -1) return i + 1;
    }
    return 0;
  }
  const idx = {};
  Object.keys(CONFIG.HEADERS).forEach(function(k){ idx[k] = findIndex(k); });

  const must = ['status','date','productName'];
  const missing = must.filter(function(k){ return !idx[k]; });
  if (missing.length) {
    const seen = headerRow.map(function(h){ return '"' + String(h) + '"'; }).join(', ');
    throw new Error('Header "' + CONFIG.HEADERS[missing[0]] +
      '" not found at row ' + CONFIG.HEADER_ROW + '. Seen headers: [' + seen + ']');
  }
  return idx;
}

/** ===================== UTILITIES ===================== */
function getTargetSheet_() {
  const ss = SpreadsheetApp.getActive();
  const byName = CONFIG.SHEET_NAME ? ss.getSheetByName(CONFIG.SHEET_NAME) : null;
  return byName || ss.getActiveSheet();
}
function normHeader_(s) {
  return String(s || '')
    .replace(/\u00A0/g, ' ')
    .replace(/[\u200B-\u200D\uFEFF]/g, '')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}
function normText_(s) {
  return String(s == null ? '' : s)
    .replace(/\u00A0/g, ' ')
    .replace(/[\u200B-\u200D\uFEFF]/g, '')
    .trim()
    .toLowerCase();
}
function hasStatus_(cell, expectedRaw) {
  const wantNorm = normText_(expectedRaw);
  const cellNorm = normText_(cell);
  if (!cellNorm || !wantNorm) return false;

  // Special handling for On Request: allow variations like "onrequest", "on  request"
  const onReqNorm = normText_(CONFIG.STATUS_FILTER_VALUE || 'On request');
  if (wantNorm === onReqNorm) {
    return (cellNorm === onReqNorm) || /on\s*request/.test(cellNorm);
  }
  return cellNorm === wantNorm;
}
function getCell_(row, oneBasedIndex) {
  return oneBasedIndex ? row[oneBasedIndex - 1] : '';
}
function normalizeDateDisp_(v, tz) {
  if (!v) return '-';
  var d = null;
  if (v instanceof Date) d = v;
  else if (/^\d{1,2}\/\d{1,2}\/\d{4}/.test(String(v))) {
    var m1 = String(v).match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
    d = new Date(+m1[3], +m1[2]-1, +m1[1]);
  } else if (/^\d{4}-\d{2}-\d{2}/.test(String(v))) {
    var m2 = String(v).match(/^(\d{4})-(\d{2})-(\d{2})/);
    d = new Date(+m2[1], +m2[2]-1, +m2[3]);
  } else if (!isNaN(+v)) {
    var ms = Math.round((+v-25569)*86400000);
    d = new Date(ms);
  }
  return d ? Utilities.formatDate(d, tz, 'd/M/yy') : String(v);
}
function formatActionForBlock_(s){
  var raw = String(s == null ? '' : s)
    .replace(/\r\n/g,'\n')
    .replace(/\u2028|\u2029/g,'\n');
  var lines = raw.split('\n').map(function(x){ return x.replace(/\s+$/,''); });
  if (lines.length === 0 || (lines.length === 1 && !lines[0])) return 'Action: -';
  if (lines.length === 1) return 'Action: ' + lines[0];
  return 'Action: ' + lines[0] + '\n ' + lines.slice(1).join('\n ');
}

/** ===== HEADER TEXT BUILDERS (for LINE messages & preview) ===== */
function buildPendingHead_() {
  const tz = CONFIG.TIMEZONE || Session.getScriptTimeZone() || 'Asia/Bangkok';
  const now = new Date();
  const dateStr = Utilities.formatDate(now, tz, 'dd/MM/yy');
  return 'Pending Task\n[' + dateStr + ']';
}

function buildOnRequestHead_() {
  const tz = CONFIG.TIMEZONE || Session.getScriptTimeZone() || 'Asia/Bangkok';
  const now = new Date();
  const dateStr = Utilities.formatDate(now, tz, 'dd/MM/yy');
  const timeStr = Utilities.formatDate(now, tz, 'hh:mm a').toLowerCase(); // e.g. 07:05 am
  return 'On request Task\n[' + dateStr + ' ' + timeStr + ']';
}

/** ===================== LINE SEND HELPERS ===================== */
function splitAndSendToLine_(text, limit) {
  const cap = limit || 4900;
  let left = String(text || '');
  let chunks = 0;
  while (left.length > cap) {
    let cut = left.lastIndexOf('\n', cap);
    if (cut <= 0) cut = cap;
    sendLineOA_(left.slice(0, cut));
    chunks++;
    left = left.slice(cut);
    if (left.startsWith('\n')) left = left.slice(1);
  }
  if (left) { sendLineOA_(left); chunks++; }
  return chunks;
}
function sendLineOA_(message) {
  const token = CONFIG.OA_CHANNEL_ACCESS_TOKEN;
  if (!token) throw new Error('Please set CONFIG.OA_CHANNEL_ACCESS_TOKEN');

  const headers = {
    Authorization: 'Bearer ' + token,
    'Content-Type': 'application/json; charset=UTF-8'
  };
  const body = { messages: [{ type: 'text', text: String(message) }] };

  let url;
  if (CONFIG.OA_GROUP_ID && /^C/.test(CONFIG.OA_GROUP_ID)) {
    url = 'https://api.line.me/v2/bot/message/push';
    body.to = CONFIG.OA_GROUP_ID;
  } else {
    url = 'https://api.line.me/v2/bot/message/broadcast';
  }

  const res = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json; charset=UTF-8',
    payload: JSON.stringify(body),
    headers: headers,
    muteHttpExceptions: true
  });

  const code = res.getResponseCode();
  const text = res.getContentText();
  if (code < 200 || code >= 300) {
    let msg = text;
    try {
      const j = JSON.parse(text);
      const details = (j.details || []).map(function(d){ return d.message; }).join('; ');
      msg = [j.message, details].filter(Boolean).join(' — ');
    } catch (e) {}
    const hint =
      code === 404 ? 'Hint: bot not in this group OR wrong groupId.' :
      code === 401 ? 'Hint: invalid/expired Channel Access Token.' :
      code === 403 ? 'Hint: permission/plan not allowed.' : '';
    throw new Error('LINE push failed (' + code + '): ' + msg + (hint ? ' ('+hint+')' : ''));
  }
  return { code: code, body: text };
}

// Quick connectivity test
function testPushToGroup(){
  const result = sendLineOA_('[TEST] Kanban Alert connectivity check: ' + new Date());
  SpreadsheetApp.getUi().alert('Push OK: ' + JSON.stringify(result));
}
