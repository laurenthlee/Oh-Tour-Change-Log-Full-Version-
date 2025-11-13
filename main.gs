/** ===================== SENT TASK (MANUAL + AUTO @07:00 for Pending) ===================== */

const CONFIG = {
  HEADER_ROW: 1,
  TIMEZONE: 'Asia/Bangkok',
  SHEET_NAME: '',

  // LINE OA (put your real values)
  OA_CHANNEL_ACCESS_TOKEN: '',
  OA_GROUP_ID: '',

  LINE_CHAR_LIMIT: 4900,
  SEPARATOR_LINE: '------------------',

  TITLES: { onrequest: 'On request', pending: 'Pending' },

  HEADERS: {
    date: 'Date', requestBy: 'Request By', productId: 'Product ID', productName: 'Product Name',
    platform: 'Platform', changedType: 'Changed Type', impact: 'Impact', action: 'Action', status: 'Status'
  },
  ALT_HEADERS: {
    status:['Progress','สถานะ'],
    requestBy:['Request by','Requested By','Requester','ผู้ร้องขอ'],
    productId:['Product Id','ProductID','รหัสสินค้า'],
    productName:['Name','Tour Name','ชื่อสินค้า','ชื่อทัวร์'],
    changedType:['Change Type','Type of Change','ประเภทการเปลี่ยนแปลง'],
    impact:['Impacts'],
    action:['Actions','รายละเอียดการเปลี่ยนแปลง'],
    date:['Acted Date','Effective Date','วันที่']
  },

  PROP_AUTO_PENDING: 'auto_pending_enabled'
};

/** ===================== MENU ===================== */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Sent Task')
    .addItem('▶  Task Preview – Pending', 'openPreviewPending')
    .addItem('▶  Manually sent task', 'openSendPicker')
    .addToUi();
}

/** ===================== PREVIEW (PENDING) ===================== */
function openPreviewPending() {
  const rows = getRowsForUi_('pending');
  const text = buildHead_('pending') + rows.map(r => r.block).join('\n' + CONFIG.SEPARATOR_LINE + '\n');

  const html = HtmlService.createHtmlOutput([
    '<!doctype html><html><head><meta charset="utf-8"><title>Task Preview – Pending</title>',
    '<style>body{margin:0;background:#f6f7f9;color:#1f2328;font:13px/1.5 system-ui,-apple-system,Segoe UI,Roboto,Arial,"Noto Sans",sans-serif}',
    '.wrap{padding:14px}.bar{display:flex;gap:8px;align-items:center;margin-bottom:10px}.h{font-weight:700}',
    '.pill{padding:2px 8px;border-radius:999px;background:#d1fae5;color:#065f46;font-weight:700;font-size:12px}',
    'button{font:inherit;border:1px solid #e5e7eb;border-radius:10px;padding:7px 10px;background:#fff;cursor:pointer}',
    'button:hover{background:#fafafa}pre{white-space:pre-wrap;border:1px solid #e5e7eb;border-radius:10px;padding:10px;background:#fff;height:68vh;overflow:auto}',
    '.hint{color:#64748b;margin-left:auto}</style></head><body><div class="wrap">',
    '<div class="bar"><span class="h">Task Preview – Pending</span><span class="pill">', String(rows.length),
    '</span><span class="hint">View-only</span>',
    '<button id="copy">Copy All</button><button id="dl">Download .txt</button>',
    '<button onclick="google.script.host.close()">Close</button></div>',
    '<pre id="big"></pre>',
    '<script>',
    'var TEXT=', JSON.stringify(text), ';',
    'document.getElementById("big").textContent=TEXT;',
    'document.getElementById("copy").onclick=function(){navigator.clipboard.writeText(TEXT).catch(()=>0)};',
    'document.getElementById("dl").onclick=function(){var b=new Blob([TEXT],{type:"text/plain"});var u=URL.createObjectURL(b);var a=document.createElement("a");a.href=u;a.download="pending-preview.txt";document.body.appendChild(a);a.click();setTimeout(function(){URL.revokeObjectURL(u);a.remove();},800);};',
    '</script></div></body></html>'
  ].join('')).setWidth(980).setHeight(680);

  SpreadsheetApp.getUi().showModalDialog(html, 'Task Preview – Pending');
}

/** ===================== SEND PICKER (MANUAL, ON REQUEST ONLY) ===================== */
function openSendPicker() {
  const state = getAutoPendingState_();
  const rows = getRowsForUi_('onrequest');
  const head = buildHead_('onrequest');

  const html = HtmlService.createHtmlOutput([
    '<!doctype html><html><head><meta charset="utf-8"><title>Manually sent task</title>',
    '<style>',
    'body{margin:0;background:#f6f7f9;color:#1f2328;font:13px/1.5 system-ui,-apple-system,Segoe UI,Roboto,Arial,"Noto Sans",sans-serif}',
    '.wrap{padding:14px}.bar{display:flex;gap:8px;align-items:center;flex-wrap:wrap;margin-bottom:10px}.h{font-weight:700}',
    '.pill{padding:2px 8px;border-radius:999px;background:#dbeafe;color:#1e3a8a;font-weight:700;font-size:12px}',
    '.status{padding:2px 8px;border-radius:999px;font-weight:700;font-size:12px}.on{background:#dcfce7;color:#166534}.off{background:#fee2e2;color:#991b1b}',
    'button{font:inherit;border:1px solid #e5e7eb;border-radius:10px;padding:7px 10px;background:#fff;cursor:pointer}',
    'button:hover{background:#fafafa}.grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(360px,1fr));gap:10px}',
    '.card{background:#fff;border:1px solid #e5e7eb;border-radius:14px;padding:10px}.hdr{display:flex;gap:8px;align-items:center;margin-bottom:6px}',
    '.name{font-weight:700;flex:1}.mono{font:12px ui-monospace,Menlo,Consolas,monospace;white-space:pre-wrap;border:1px solid #e5e7eb;border-radius:10px;padding:8px;background:#fff}',
    '.label{color:#64748b;font-size:12px;margin-top:4px}',
    '</style></head><body><div class="wrap">',
    '<div class="bar">',
    '<span class="h">Manually sent task</span>',
    '<label style="margin-left:12px;display:flex;align-items:center;gap:6px;">',
    '<input type="checkbox" id="auto7" ', (state.enabled ? 'checked' : ''), '>',
    '<span>Auto 07:00 (Pending)</span>',
    '<span id="autoStatus" class="status ', (state.enabled ? 'on' : 'off'), '">', (state.enabled ? 'ON' : 'OFF'), '</span>',
    '</label>',
    '<span class="pill" id="allCount">All: ', String(rows.length), '</span>',
    '<span class="pill" id="selCount">Selected: 0</span>',
    '<span style="margin-left:auto"></span>',
    '<button id="selAll">Select all</button><button id="clear">Clear</button>',
    '<button id="send">Send selected</button><button onclick="google.script.host.close()">Close</button>',
    '</div>',
    '<div id="list" class="grid"></div>',
    '<div class="label" style="margin-top:10px">รวมข้อความที่จะส่ง</div>',
    '<pre id="big" class="mono"></pre>',
    '<script>',
    'var DATA=', JSON.stringify(rows), ';',
    'var HEAD=', JSON.stringify(head), ';',
    'var picked=new Set();',
    'function esc(s){return String(s==null?"":s).replace(/[&<>"\\\']/g,function(c){return {"&":"&amp;","<":"&lt;",">":"&gt;","\\"":"&quot;","\\\'":"&#39;"}[c];});}',
    'function render(){document.getElementById("allCount").textContent="All: "+DATA.length;',
    ' var box=document.getElementById("list");',
    ' box.innerHTML=DATA.map(function(r){return [' ,
    '  \'<div class="card">\',',
    '  \'<div class="hdr">\',',
    '  \'<input type="checkbox" class="chk" data-row="\'+r.row+\'" \'+(picked.has(r.row)?\'checked\':\'\')+\'>\',',
    '  \'<div class="name">\',esc(r.productName||"-"),\'</div>\',',
    '  \'<a target="_blank" href="\',esc(r.openUrl),\'">#\',r.row,\'</a>\',',
    '  \'</div>\',',
    '  \'<div class="label">Action</div>\',',
    '  \'<div class="mono">\',esc(r.action||"-"),\'</div>\',',
    '  \'</div>\'',
    ' ].join("");}).join("");',
    ' document.querySelectorAll(".chk").forEach(function(el){',
    '   el.onchange=function(){var n=Number(this.dataset.row); if(this.checked){picked.add(n);}else{picked.delete(n);} updateBig();};',
    ' });',
    ' updateBig();}',
    'function updateBig(){var rows=DATA.filter(function(r){return picked.has(r.row);});',
    ' document.getElementById("selCount").textContent="Selected: "+rows.length;',
    ' var body=rows.map(function(r){return r.block;}).join("\\n'+CONFIG.SEPARATOR_LINE+'\\n");',
    ' document.getElementById("big").textContent=HEAD+body;}',
    'document.getElementById("selAll").onclick=function(){DATA.forEach(function(r){picked.add(r.row);}); render();};',
    'document.getElementById("clear").onclick=function(){picked.clear(); render();};',
    'document.getElementById("send").onclick=function(){var arr=DATA.filter(function(r){return picked.has(r.row);}); if(!arr.length){return;}',
    ' google.script.run.withSuccessHandler(function(msg){/* silent */}).withFailureHandler(function(err){console.log(err);}).sendSelected_("onrequest", arr);};',
    // auto 07:00 toggle (silent save + inline ON/OFF)
    'document.getElementById("auto7").onchange=function(){var en=this.checked; var self=this;',
    ' self.disabled=true;',
    ' google.script.run.withSuccessHandler(function(res){self.disabled=false; var st=document.getElementById("autoStatus");',
    '   if(res && res.enabled){st.textContent="ON"; st.className="status on";} else {st.textContent="OFF"; st.className="status off";}',
    ' }).withFailureHandler(function(e){self.disabled=false; self.checked=!en; console.log(e);}).setAutoPendingEnabled(en);',
    '};',
    'render();',
    '</script></div></body></html>'
  ].join('')).setWidth(980).setHeight(680);

  SpreadsheetApp.getUi().showModalDialog(html, 'Manually sent task');
}

/** ===================== SERVER SENDERS ===================== */
function sendSelected_(kind, pickedRows) {
  if (!pickedRows || !pickedRows.length) return 'No rows selected';
  const head = buildHead_(kind);
  const body = pickedRows.map(r => r.block).join('\n' + CONFIG.SEPARATOR_LINE + '\n');
  const text = head + body;
  const chunks = splitAndSendToLine_(text, CONFIG.LINE_CHAR_LIMIT);
  return 'Sent ' + pickedRows.length + ' item(s) in ' + chunks + ' chunk(s).';
}

/** ===================== AUTO @07:00 (PENDING) ===================== */
function setAutoPendingEnabled(enabled) {
  const p = PropertiesService.getScriptProperties();
  p.setProperty(CONFIG.PROP_AUTO_PENDING, enabled ? '1' : '0');
  if (enabled) installPendingTrigger_(); else removePendingTrigger_();
  return { enabled: !!enabled };
}
function getAutoPendingState_() {
  const p = PropertiesService.getScriptProperties();
  return { enabled: p.getProperty(CONFIG.PROP_AUTO_PENDING) === '1' };
}
function installPendingTrigger_() {
  removePendingTrigger_();
  ScriptApp.newTrigger('pendingAutoSendRunner_')
    .timeBased()
    .atHour(7).everyDays(1)
    .inTimezone(CONFIG.TIMEZONE || 'Asia/Bangkok')
    .create();
}
function removePendingTrigger_() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction && t.getHandlerFunction() === 'pendingAutoSendRunner_') {
      ScriptApp.deleteTrigger(t);
    }
  });
}
function pendingAutoSendRunner_() {
  const rows = getRowsForUi_('pending');
  if (!rows.length) return;
  const head = buildHead_('pending');
  const body = rows.map(r => r.block).join('\n' + CONFIG.SEPARATOR_LINE + '\n');
  splitAndSendToLine_(head + body, CONFIG.LINE_CHAR_LIMIT);
}

/** ===================== DATA LAYER ===================== */
function getRowsForUi_(kind) {
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
  rows.forEach((r, off) => {
    const statusCell = getCell_(r, idx.status);
    if (!matchStatus_(statusCell, kind)) return;

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
      'Request By: ' + (obj.requestBy || '-'),
      'Product ID: ' + (obj.productId || '-'),
      'Product Name: ' + (obj.productName || '-'),
      'Platform: ' + (obj.platform || '-'),
      'Changed Type: ' + (obj.changedType || '-'),
      'Impact: ' + (obj.impact || '-'),
      formatActionForBlock_(obj.action)
    ].join('\n');

    obj.openUrl = base + '#gid=' + gid + '&range=A' + rownum + ':A' + rownum;
    out.push(obj);
  });
  return out;
}

function buildHead_(kind) {
  const title = CONFIG.TITLES[kind] || 'On request';
  return title + '\n' + CONFIG.SEPARATOR_LINE + '\n';
}
function matchStatus_(cell, kind) {
  const t = normText_(cell);
  if (!t) return false;
  if (kind === 'pending') return /pending/.test(t);
  return /on\s*request/.test(t);
}

/** ===================== HEADER & UTILS ===================== */
function robustIndexMap_(headerRow) {
  const normRow = headerRow.map(normHeader_);
  function findIndex(key) {
    const main = CONFIG.HEADERS[key];
    const alts = (CONFIG.ALT_HEADERS[key] || []);
    const candidates = [main].concat(alts).map(normHeader_);
    for (let i = 0; i < normRow.length; i++) if (candidates.indexOf(normRow[i]) !== -1) return i + 1;
    return 0;
  }
  const idx = {};
  Object.keys(CONFIG.HEADERS).forEach(k => idx[k] = findIndex(k));
  const must = ['status','date','productName'];
  const missing = must.filter(k => !idx[k]);
  if (missing.length) {
    const seen = headerRow.map(h => '"' + String(h) + '"').join(', ');
    throw new Error('Header "' + CONFIG.HEADERS[missing[0]] + `" not found at row ${CONFIG.HEADER_ROW}. Seen: [` + seen + ']');
  }
  return idx;
}
function getTargetSheet_() {
  const ss = SpreadsheetApp.getActive();
  const byName = CONFIG.SHEET_NAME ? ss.getSheetByName(CONFIG.SHEET_NAME) : null;
  return byName || ss.getActiveSheet();
}
function normHeader_(s) { return String(s||'').replace(/\u00A0/g,' ').replace(/[\u200B-\u200D\uFEFF]/g,'').replace(/\s+/g,' ').trim().toLowerCase(); }
function normText_(s) { return String(s==null?'':s).replace(/\u00A0/g,' ').replace(/[\u200B-\u200D\uFEFF]/g,'').trim().toLowerCase(); }
function getCell_(row, i){ return i ? row[i-1] : ''; }
function normalizeDateDisp_(v, tz) {
  if (!v) return '-';
  let d=null;
  if (v instanceof Date) d=v;
  else if (/^\d{1,2}\/\d{1,2}\/\d{4}/.test(String(v))){const m=String(v).match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/); d=new Date(+m[3],+m[2]-1,+m[1]);}
  else if (/^\d{4}-\d{2}-\d{2}/.test(String(v))){const m=String(v).match(/^(\d{4})-(\d{2})-(\d{2})/); d=new Date(+m[1],+m[2]-1,+m[3]);}
  else if (!isNaN(+v)){const ms=Math.round((+v-25569)*86400000); d=new Date(ms);}
  return d ? Utilities.formatDate(d, tz, 'd/M/yy') : String(v);
}
function formatActionForBlock_(s){
  const raw = String(s==null?'':s).replace(/\r\n/g,'\n').replace(/\u2028|\u2029/g,'\n');
  const lines = raw.split('\n').map(x=>x.replace(/\s+$/,''));
  if (lines.length===0 || (lines.length===1 && !lines[0])) return 'Action: -';
  if (lines.length===1) return 'Action: ' + lines[0];
  return 'Action: ' + lines[0] + '\n ' + lines.slice(1).join('\n ');
}

/** ===================== LINE SENDER ===================== */
function splitAndSendToLine_(text, limit) {
  const cap = limit || 4900;
  let left = String(text||''); let chunks=0;
  while (left.length > cap) { let cut = left.lastIndexOf('\n', cap); if (cut<=0) cut=cap; sendLineOA_(left.slice(0,cut)); chunks++; left = left.slice(cut); if (left.startsWith('\n')) left = left.slice(1); }
  if (left){ sendLineOA_(left); chunks++; }
  return chunks;
}
function sendLineOA_(message) {
  const token = CONFIG.OA_CHANNEL_ACCESS_TOKEN;
  if (!token) throw new Error('Please set CONFIG.OA_CHANNEL_ACCESS_TOKEN');
  const headers = { Authorization: 'Bearer '+token, 'Content-Type': 'application/json; charset=UTF-8' };
  const body = { messages: [{ type:'text', text:String(message) }] };
  let url; if (CONFIG.OA_GROUP_ID && /^C/.test(CONFIG.OA_GROUP_ID)){ url='https://api.line.me/v2/bot/message/push'; body.to=CONFIG.OA_GROUP_ID; } else { url='https://api.line.me/v2/bot/message/broadcast'; }
  const res = UrlFetchApp.fetch(url, { method:'post', contentType:'application/json; charset=UTF-8', payload:JSON.stringify(body), headers, muteHttpExceptions:true });
  const code=res.getResponseCode(), text=res.getContentText();
  if (code<200||code>=300){
    let msg=text; try{const j=JSON.parse(text); const details=(j.details||[]).map(d=>d.message).join('; '); msg=[j.message,details].filter(Boolean).join(' — ');}catch(_){}
    const hint = code===404?'Hint: bot not in this group OR wrong groupId.': code===401?'Hint: invalid/expired Channel Access Token.': code===403?'Hint: permission/plan not allowed.':'';
    throw new Error('LINE push failed ('+code+'): '+msg+(hint?' ('+hint+')':''));
  }
  return { code, body:text };
}
