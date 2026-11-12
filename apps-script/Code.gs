/**
 * 喬琳醫美 Call 客工作台：後端（Google Apps Script）
 *
 * 這段程式碼要貼在「你自己的 Google 試算表」→ 擴充功能 → Apps Script 裡。
 * 所有客戶資料、分配和撥打紀錄都存在這份試算表，網頁本身不存任何資料。
 * 密碼只在這裡檢查，網頁拿不到任何人的密碼。
 */

var SHEET = { ACC: '帳號', CUST: '客戶', ASG: '分配', RES: '撥打紀錄', LOG: '操作紀錄' };
var STAGING_PREFIX = '_匯入中_';
var TZ = 'Asia/Taipei';

// 客戶分頁只存這些欄位（刻意不存證件號碼、完整地址）
var CUST_KEYS = ['id','n','g','ph','b','c','d','s','dr','deal','st','vt','fv','lv','lc','la','tot','owe','sv','ch0','ch1','ch2','cb','ln'];
var CUST_LABELS = ['會員號','姓名','性別','手機','生日','縣市','行政區','所屬諮詢師','主診醫生','成交狀態','會員狀態','最近到訪類型','首次到訪','最近到訪','最近消費日','最近消費金額','消費總額','欠費金額','儲值金餘額','渠道分類','一級渠道','二級渠道','最近回訪','LINE綁定'];
var ASG_LABELS = ['會員號','業務','分配日','分配人'];
var RES_LABELS = ['會員號','業務','撥打結果','備註','撥打次數','最近更新'];
var ACC_LABELS = ['名字','密碼','角色','啟用'];
var RESULTS = ['未接通','已接通-有興趣','已接通-考慮中','已接通-無意願','已約診','空號/錯號','請勿再撥'];

var SESSION_SEC = 21600;      // 登入 6 小時有效（每次使用會自動延長）
var MAX_FAILS = 5;            // 密碼連錯 5 次
var LOCKOUT_SEC = 600;        // 鎖 10 分鐘

/* ================= 試算表選單 ================= */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Call 客工作台')
    .addItem('建立／檢查分頁', 'setup')
    .addItem('把帳號分頁的密碼加密', 'hashPasswords')
    .addToUi();
}

function setup() {
  var ss = SpreadsheetApp.getActive();
  var acc = ensureSheet(ss, SHEET.ACC, ACC_LABELS);
  if (acc.getLastRow() < 2) {
    acc.getRange(2, 1, 1, 4).setValues([['（主管名字）', '（設定密碼）', '主管', true]]);
  }
  acc.getRange('B:B').setNumberFormat('@');
  acc.getRange(2, 3, 200, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(['主管', '業務'], true).build());
  acc.getRange(2, 4, 200, 1).insertCheckboxes();
  acc.setColumnWidth(1, 120); acc.setColumnWidth(2, 160);
  ensureSheet(ss, SHEET.CUST, CUST_LABELS);
  ensureSheet(ss, SHEET.ASG, ASG_LABELS);
  ensureSheet(ss, SHEET.RES, RES_LABELS);
  ensureSheet(ss, SHEET.LOG, ['時間', '誰', '動作', '內容']);
  try {
    SpreadsheetApp.getUi().alert('分頁已建立。請到「帳號」分頁填入名字、密碼和角色（主管／業務），再依說明部署成網頁應用程式。');
  } catch (e) {}
}

function ensureSheet(ss, name, labels) {
  var sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, labels.length).setValues([labels]).setFontWeight('bold');
    sh.setFrozenRows(1);
    if (name !== SHEET.ACC) sh.getRange(1, 1, sh.getMaxRows(), labels.length).setNumberFormat('@');
  }
  return sh;
}

function hashPasswords() {
  var sh = SpreadsheetApp.getActive().getSheetByName(SHEET.ACC);
  if (!sh || sh.getLastRow() < 2) return;
  var rng = sh.getRange(2, 2, sh.getLastRow() - 1, 1);
  var v = rng.getValues(), n = 0;
  for (var i = 0; i < v.length; i++) {
    var pw = String(v[i][0]);
    if (pw && pw.indexOf('sha256$') !== 0) {
      var salt = Utilities.getUuid().replace(/-/g, '').slice(0, 16);
      v[i][0] = 'sha256$' + salt + '$' + sha256(salt + pw); n++;
    }
  }
  rng.setValues(v);
  try { SpreadsheetApp.getUi().alert('已加密 ' + n + ' 組密碼。之後要改密碼，直接在格子裡輸入新密碼即可（可再按一次加密）。'); } catch (e) {}
}

/* ================= 網頁入口 ================= */
function doGet() { return out({ ok: true, app: 'callws' }); }

function doPost(e) {
  var req;
  try { req = JSON.parse(e.postData.contents); } catch (x) { return out({ ok: false, err: '請求格式錯誤' }); }
  try {
    return out(route(req || {}));
  } catch (err) {
    if (err && err.userMsg) return out({ ok: false, err: err.userMsg, code: err.code || '' });
    console.error(err && err.stack || err);
    return out({ ok: false, err: '伺服器發生錯誤，請稍後再試' });
  }
}

function out(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function fail(msg, code) { var e = new Error(msg); e.userMsg = msg; e.code = code; throw e; }

function route(req) {
  var a = String(req.action || '');
  if (a === 'login') return login(req.name, req.password);
  if (a === 'public') return publicInfo();
  var me = auth(req.token);
  if (a === 'logout') { CacheService.getScriptCache().remove('s_' + req.token); return { ok: true }; }
  if (a === 'load') return load(me, req);
  if (a === 'saveResult') { needRole(me, '業務'); return withLock(function () { return saveResult(me, req); }); }
  if (a === 'assign') { needRole(me, '主管'); return withLock(function () { return assign(me, req); }); }
  if (a === 'staffAdd') { needRole(me, '主管'); return withLock(function () { return staffAdd(me, req); }); }
  if (a === 'staffRemove') { needRole(me, '主管'); return withLock(function () { return staffRemove(me, req); }); }
  if (a === 'importStart') { needRole(me, '主管'); return importStart(me); }
  if (a === 'importChunk') { needRole(me, '主管'); return importChunk(me, req); }
  if (a === 'importCommit') { needRole(me, '主管'); return withLock(function () { return importCommit(me, req); }); }
  fail('不明的操作');
}

function needRole(me, role) { if (me.role !== role) fail('你的帳號沒有這個權限'); }

function withLock(fn) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(25000)) fail('另一位同事正在更新名單，請幾秒後再試一次', 'busy');
  try { return fn(); } finally { lock.releaseLock(); }
}

/* ================= 帳號與登入 ================= */
function sha256(s) {
  var b = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s, Utilities.Charset.UTF_8);
  return b.map(function (x) { return ('0' + (x & 255).toString(16)).slice(-2); }).join('');
}

function readAccounts() {
  var sh = SpreadsheetApp.getActive().getSheetByName(SHEET.ACC);
  if (!sh || sh.getLastRow() < 2) return [];
  var v = sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues();
  var list = [];
  v.forEach(function (r) {
    var name = String(r[0]).trim(), role = String(r[2]).trim(), on = r[3];
    if (!name || name.charAt(0) === '（') return;
    if (role !== '主管' && role !== '業務') return;
    if (on === false || /^(false|否|停用|n|no|0)$/i.test(String(on).trim())) return;
    list.push({ name: name, pw: String(r[1]), role: role });
  });
  return list;
}

function checkPw(stored, given) {
  if (!stored || !given) return false;
  if (stored.indexOf('sha256$') === 0) {
    var p = stored.split('$');
    return p.length === 3 && sha256(p[1] + given) === p[2];
  }
  return stored === given;
}

function login(name, password) {
  name = String(name || '').trim(); password = String(password || '');
  if (!name) fail('請輸入名字');
  var cache = CacheService.getScriptCache();
  var fk = 'f_' + sha256(name.toLowerCase());
  var fails = +(cache.get(fk) || 0);
  if (fails >= MAX_FAILS) fail('密碼錯太多次，請 10 分鐘後再試');
  var acc = readAccounts().filter(function (a) { return a.name === name; })[0];
  // 業務的密碼欄留空 → 不用密碼；主管一定要密碼
  var open = acc && acc.role === '業務' && !acc.pw;
  if (acc && !open && !password) fail(acc.role === '主管' && !acc.pw ? '主管帳號必須在試算表設定密碼' : '請輸入密碼', 'needpw');
  if (!acc || (!open && !checkPw(acc.pw, password))) {
    cache.put(fk, String(fails + 1), LOCKOUT_SEC);
    Utilities.sleep(600);
    fail('名字或密碼不正確');
  }
  cache.remove(fk);
  var token = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
  cache.put('s_' + token, JSON.stringify({ name: acc.name, role: acc.role }), SESSION_SEC);
  logOp(acc.name, '登入', acc.role);
  return { ok: true, token: token, name: acc.name, role: acc.role };
}

function auth(token) {
  if (!token || typeof token !== 'string' || token.length < 40) fail('請先登入', 'auth');
  var cache = CacheService.getScriptCache();
  var raw = cache.get('s_' + token);
  if (!raw) fail('登入已過期，請重新登入', 'auth');
  var s = JSON.parse(raw);
  // 每次都對照帳號分頁：停用或改角色會立即生效
  var acc = readAccounts().filter(function (a) { return a.name === s.name && a.role === s.role; })[0];
  if (!acc) { cache.remove('s_' + token); fail('這個帳號已停用或權限已變更，請重新登入', 'auth'); }
  cache.put('s_' + token, raw, SESSION_SEC);
  return s;
}

/* ================= 主管：在網頁上新增／移除業務 ================= */
function accRows() {
  var sh = SpreadsheetApp.getActive().getSheetByName(SHEET.ACC);
  var n = Math.max(sh.getLastRow() - 1, 0);
  return { sh: sh, v: n ? sh.getRange(2, 1, n, 4).getValues() : [] };
}

function staffAdd(me, req) {
  var name = cleanCell(req.name).slice(0, 30);
  var pw = String(req.password || '').slice(0, 50);
  if (!name) fail('請輸入名字');
  var a = accRows(), row = -1;
  for (var i = 0; i < a.v.length; i++) if (String(a.v[i][0]).trim() === name) { row = i; break; }
  if (row >= 0) {
    if (String(a.v[row][2]).trim() === '主管') fail('「' + name + '」已經是主管帳號');
    var on = a.v[row][3];
    if (on !== false && !/^(false|否|停用|n|no|0)$/i.test(String(on).trim()) && String(a.v[row][2]).trim() === '業務') fail('名單裡已經有「' + name + '」');
    a.sh.getRange(row + 2, 3).setValue('業務');
    a.sh.getRange(row + 2, 4).insertCheckboxes().setValue(true);
    if (pw) a.sh.getRange(row + 2, 2).setNumberFormat('@').setValue(pw);
  } else {
    var r = a.v.length + 2;
    ensureRows(a.sh, r);
    a.sh.getRange(r, 2).setNumberFormat('@');
    a.sh.getRange(r, 1, 1, 3).setValues([[name, pw, '業務']]);
    a.sh.getRange(r, 4).insertCheckboxes().setValue(true);
  }
  bumpRev();
  logOp(me.name, '新增業務', name + (pw ? '（有密碼）' : ''));
  return load(me, { custAt: req.custAt });
}

function staffRemove(me, req) {
  var name = String(req.name || '').trim();
  var a = accRows(), found = false;
  for (var i = 0; i < a.v.length; i++) {
    if (String(a.v[i][0]).trim() === name && String(a.v[i][2]).trim() === '業務') {
      a.sh.getRange(i + 2, 4).insertCheckboxes().setValue(false); found = true;
    }
  }
  if (!found) fail('找不到這位業務');
  bumpRev();
  logOp(me.name, '移除業務', name);
  return load(me, { custAt: req.custAt });
}

/* ================= 不用登入的首頁資訊：只有人數，沒有客戶資料 ================= */
function publicInfo() {
  var accs = readAccounts();
  var st = readState();
  var staff = accs.filter(function (a) { return a.role === '業務'; }).map(function (a) { return { name: a.name, pw: !!a.pw }; });
  return { ok: true, staff: staff, summary: summarize(st) };
}

/* ================= 讀取資料 ================= */
function props() { return PropertiesService.getScriptProperties(); }
function getRev() { return props().getProperty('rev') || '0'; }
function bumpRev() { var r = String(+getRev() + 1); props().setProperty('rev', r); return r; }
function nowStr() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm'); }
function todayStr() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'); }

function sheetRows(name, width) {
  var sh = SpreadsheetApp.getActive().getSheetByName(name);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, width).getDisplayValues();
}

function staffNames() {
  return readAccounts().filter(function (a) { return a.role === '業務'; }).map(function (a) { return a.name; });
}

/** 讀取目前狀態：客戶、有效分配、撥打紀錄 */
function readState() {
  var staff = staffNames();
  var staffSet = {}; staff.forEach(function (s) { staffSet[s] = 1; });
  var cust = sheetRows(SHEET.CUST, CUST_KEYS.length).filter(function (r) { return r[0]; });
  var custIdx = {}; cust.forEach(function (r, i) { custIdx[r[0]] = i; });
  var asg = {};
  sheetRows(SHEET.ASG, 4).forEach(function (r) {
    // 業務停用或客人已不在報表 → 視為未分配
    if (r[0] && staffSet[r[1]] && (r[0] in custIdx) && !asg[r[0]]) asg[r[0]] = { o: r[1], at: r[2], by: r[3] };
  });
  var res = sheetRows(SHEET.RES, 6).filter(function (r) { return r[0] && r[1]; });
  return { staff: staff, cust: cust, custIdx: custIdx, asg: asg, res: res };
}

function summarize(st) {
  var last = {};
  st.res.forEach(function (r) { if (r[2] && (!last[r[0]] || r[5] > last[r[0]].at)) last[r[0]] = { r: r[2], at: r[5] }; });
  var per = {}; st.staff.forEach(function (s) { per[s] = { name: s, n: 0, called: 0, booked: 0 }; });
  var assigned = 0;
  for (var id in st.asg) {
    var p = per[st.asg[id].o]; if (!p) continue;
    assigned++; p.n++;
    if (last[id]) { p.called++; if (last[id].r === '已約診') p.booked++; }
  }
  return { total: st.cust.length, unassigned: st.cust.length - assigned, staff: st.staff.map(function (s) { return per[s]; }) };
}

function load(me, req) {
  var rev = getRev();
  if (req && req.rev && String(req.rev) === rev) return { ok: true, same: true, rev: rev };
  var st = readState();
  var custAt = props().getProperty('custAt') || '';
  var meta = { custAt: custAt, files: props().getProperty('custFiles') || '' };
  var o = { ok: true, rev: rev, me: me, staff: st.staff, summary: summarize(st), keys: CUST_KEYS, meta: meta };
  if (me.role === '主管') {
    o.cust = (req && req.custAt && req.custAt === custAt) ? null : st.cust;
    o.asg = Object.keys(st.asg).map(function (id) { var a = st.asg[id]; return [id, a.o, a.at]; });
    o.res = st.res;
  } else {
    // 業務只拿得到分配給自己的客人
    var mine = {};
    for (var id in st.asg) if (st.asg[id].o === me.name) mine[id] = 1;
    o.cust = st.cust.filter(function (r) { return mine[r[0]]; });
    o.asg = Object.keys(mine).map(function (id) { return [id, me.name, st.asg[id].at]; });
    o.res = st.res.filter(function (r) { return mine[r[0]]; });
  }
  return o;
}

/* ================= 業務：記錄撥打結果 ================= */
function saveResult(me, req) {
  var id = String(req.id || '');
  var ss = SpreadsheetApp.getActive();
  var owner = '';
  sheetRows(SHEET.ASG, 2).some(function (r) { if (r[0] === id) { owner = r[1]; return true; } });
  if (owner !== me.name) fail('這位客人已經不在你的名單裡（可能被主管收回或改派）', 'gone');
  var sh = ss.getSheetByName(SHEET.RES);
  var rows = sheetRows(SHEET.RES, 6), at = -1;
  for (var i = 0; i < rows.length; i++) if (rows[i][0] === id && rows[i][1] === me.name) { at = i; break; }
  var e = at >= 0 ? rows[at].slice() : [id, me.name, '', '', '0', ''];
  if ('r' in req) {
    var r = String(req.r || '');
    if (r && RESULTS.indexOf(r) < 0) fail('撥打結果不正確');
    e[2] = r;
    if (req.call && r) e[4] = String((+e[4] || 0) + 1);
  }
  if ('n' in req) e[3] = String(req.n || '').slice(0, 200);
  e[5] = nowStr();
  var rowNo = at >= 0 ? at + 2 : Math.max(sh.getLastRow(), 1) + 1;
  ensureRows(sh, rowNo);
  sh.getRange(rowNo, 1, 1, 6).setNumberFormat('@').setValues([e]);
  return { ok: true, entry: e, rev: bumpRev() };
}

function ensureRows(sh, n) { var m = sh.getMaxRows(); if (m < n) sh.insertRowsAfter(m, n - m); }

/* ================= 主管：分配 ================= */
/**
 * req.mode: 'to'（分配給某人）、'spread'（平均分）、'byCons'（分給所屬諮詢師）、'unassign'（收回）
 * req.ids:  客人會員號
 * req.seen: { 會員號: 主管畫面上看到的業務（未分配為空字串） }
 *           伺服器上的現況跟畫面不同（例如另一位主管剛分配過）→ 這位客人不動，避免互相覆蓋。
 */
function assign(me, req) {
  var st = readState();
  var ids = Array.isArray(req.ids) ? req.ids.map(String) : [];
  var seen = req.seen || {};
  var staffSet = {}; st.staff.forEach(function (s) { staffSet[s] = 1; });
  var mode = req.mode;
  if (mode === 'to' && !staffSet[req.target]) fail('找不到這位業務，請重新整理');
  if (mode === 'spread' && !st.staff.length) fail('還沒有任何業務帳號');
  var cIdx = CUST_KEYS.indexOf('s');
  var load = {}; st.staff.forEach(function (s) { load[s] = 0; });
  for (var k in st.asg) load[st.asg[k].o]++;
  var res = { added: 0, moved: 0, removed: 0, skipped: 0, conflict: 0 };
  var changes = {}, uniq = {};
  ids.forEach(function (id) {
    if (uniq[id]) return; uniq[id] = 1;
    if (!(id in st.custIdx)) { res.skipped++; return; }
    var cur = st.asg[id] ? st.asg[id].o : '';
    if (id in seen && String(seen[id] || '') !== cur) { res.conflict++; return; }
    var t;
    if (mode === 'to') t = req.target;
    else if (mode === 'unassign') t = '';
    else if (mode === 'byCons') { var s = st.cust[st.custIdx[id]][cIdx]; if (!staffSet[s]) { res.skipped++; return; } t = s; }
    else if (mode === 'spread') {
      if (cur) load[cur]--;
      t = st.staff[0]; st.staff.forEach(function (s) { if (load[s] < load[t]) t = s; });
      load[t]++;
    } else fail('不明的分配方式');
    if (t === cur) return;
    if (mode !== 'spread') { if (cur) load[cur]--; if (t) load[t]++; }
    changes[id] = t;
    if (t && cur) res.moved++; else if (t) res.added++; else res.removed++;
  });
  if (Object.keys(changes).length) {
    var today = todayStr();
    var sh = SpreadsheetApp.getActive().getSheetByName(SHEET.ASG);
    var all = sheetRows(SHEET.ASG, 4).filter(function (r) { return r[0] && !(r[0] in changes); });
    for (var id in changes) if (changes[id]) all.push([id, changes[id], today, me.name]);
    var old = Math.max(sh.getLastRow() - 1, 0);
    ensureRows(sh, all.length + 1);
    if (old) sh.getRange(2, 1, old, 4).clearContent();
    if (all.length) sh.getRange(2, 1, all.length, 4).setNumberFormat('@').setValues(all);
    bumpRev();
    logOp(me.name, '分配：' + mode, JSON.stringify(res) + (req.target ? ' → ' + req.target : ''));
  }
  var o = load_(me, req.custAt);
  o.result = res;
  return o;
}
function load_(me, custAt) { return load(me, { custAt: custAt }); }

/* ================= 主管：匯入客戶報表（分批上傳） ================= */
function importStart(me) {
  var ss = SpreadsheetApp.getActive();
  // 清掉 1 小時前沒完成的匯入
  ss.getSheets().forEach(function (sh) {
    var n = sh.getName();
    if (n.indexOf(STAGING_PREFIX) === 0) {
      var t = +n.split('_').pop();
      if (!t || Date.now() - t > 3600000) ss.deleteSheet(sh);
    }
  });
  var id = STAGING_PREFIX + Date.now();
  var sh = ss.insertSheet(id);
  sh.getRange(1, 1, 1, CUST_LABELS.length).setValues([CUST_LABELS]).setFontWeight('bold');
  sh.hideSheet();
  return { ok: true, importId: id };
}

function stagingSheet(id) {
  if (typeof id !== 'string' || id.indexOf(STAGING_PREFIX) !== 0) fail('匯入編號不正確');
  var sh = SpreadsheetApp.getActive().getSheetByName(id);
  if (!sh) fail('匯入逾時，請重新匯入');
  return sh;
}

function cleanCell(v) { return v === null || v === undefined ? '' : String(v).replace(/[\r\n\t]+/g, ' ').trim().slice(0, 100); }

function importChunk(me, req) {
  var sh = stagingSheet(req.importId);
  var rows = Array.isArray(req.rows) ? req.rows : [];
  // 只收白名單內的欄位：就算網頁送多了，也不會存進試算表
  var vals = [];
  rows.forEach(function (o) {
    if (!o || !o.id) return;
    vals.push(CUST_KEYS.map(function (k) { return cleanCell(o[k]); }));
  });
  if (vals.length) {
    var start = sh.getLastRow() + 1;
    ensureRows(sh, start + vals.length - 1);
    sh.getRange(start, 1, vals.length, CUST_KEYS.length).setNumberFormat('@').setValues(vals);
  }
  return { ok: true, n: vals.length };
}

function importCommit(me, req) {
  var ss = SpreadsheetApp.getActive();
  var sh = stagingSheet(req.importId);
  var n = Math.max(sh.getLastRow() - 1, 0);
  if (!n) fail('沒有收到客戶資料');
  // 同一會員號只留一筆（以最後出現的為準）
  var v = sh.getRange(2, 1, n, CUST_KEYS.length).getDisplayValues();
  var map = {}; v.forEach(function (r) { if (r[0]) map[r[0]] = r; });
  var ids = Object.keys(map).sort();
  var uniq = ids.map(function (k) { return map[k]; });
  if (uniq.length !== n) {
    sh.getRange(2, 1, n, CUST_KEYS.length).clearContent();
    sh.getRange(2, 1, uniq.length, CUST_KEYS.length).setNumberFormat('@').setValues(uniq);
  }
  var old = ss.getSheetByName(SHEET.CUST);
  var pos = old ? old.getIndex() : ss.getNumSheets();
  if (old) ss.deleteSheet(old);
  sh.setName(SHEET.CUST);
  sh.showSheet();
  sh.setFrozenRows(1);
  try { ss.setActiveSheet(sh); ss.moveActiveSheet(Math.min(pos, ss.getNumSheets())); } catch (e) {}
  var files = Array.isArray(req.files) ? req.files.map(cleanCell).join('、') : '';
  props().setProperties({ custAt: new Date().toISOString(), custFiles: files });
  bumpRev();
  logOp(me.name, '匯入客戶報表', uniq.length + ' 位；' + files);
  var o = load_(me, '');
  o.imported = uniq.length;
  return o;
}

/* ================= 紀錄 ================= */
function logOp(who, what, detail) {
  try {
    var sh = SpreadsheetApp.getActive().getSheetByName(SHEET.LOG);
    if (sh) sh.appendRow([nowStr(), who, what, String(detail || '').slice(0, 500)]);
  } catch (e) {}
}
