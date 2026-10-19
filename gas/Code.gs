/**
 * おりーぶ庵 打刻管理 Ver.2  サーバー側 (Google Apps Script Web App)
 * - 保存先: スプレッドシート (従業員 / 打刻 / 修正申請 / 設定)
 * - 認証: 従業員番号 + 6桁PIN (SHA-256 + ソルト)。管理者は Script Properties の ADMIN_ID のみ
 * - 打刻時刻は必ずサーバー時刻 (端末の時計は使わない)
 * 初回は setup() → issueInitialPins() を実行してからWebアプリとしてデプロイする。
 */
var TZ = 'Asia/Tokyo';
var SH = { EMP: '従業員', PUNCH: '打刻', REQ: '修正申請', CONF: '設定', PIN: '初期PIN配布' };
var HEAD = {
  EMP: ['従業員ID', '氏名', 'カナ', '区分', '拠点', '在籍', 'PINハッシュ', 'ソルト', '要PIN変更', '失敗回数', 'ロック解除', '最終ログイン'],
  PUNCH: ['記録ID', '従業員ID', '氏名', '勤務区分', '勤務先', '出勤日時', '退勤日時', '休憩分', '実労働分', '状態', '備考', '登録日時', '更新日時'],
  REQ: ['申請ID', '申請日時', '従業員ID', '氏名', '記録ID', '修正前出勤', '修正前退勤', '修正後出勤', '修正後退勤', '修正後勤務区分', '理由', '状態', '承認者', '承認日時', '承認コメント']
};
var SHIFTS = ['日勤', '夜勤', '直行直帰'];
var TOKEN_DAYS = 14, MAX_FAIL = 5, LOCK_MIN = 15;

/* ---------- エントリポイント ---------- */
function doGet() { return respond({ ok: true, service: 'olive-attendance', version: 2 }); }

function doPost(e) {
  var out;
  try {
    var req = JSON.parse(e.postData.contents);
    out = route(req);
  } catch (err) {
    if (!err.userMessage) console.error(err && err.stack || err);
    out = { ok: false, error: err.userMessage || '処理に失敗しました。時間をおいてもう一度お試しください', code: err.code || '' };
  }
  return respond(out);
}

function respond(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function fail(msg, code) { var e = new Error(msg); e.userMessage = msg; e.code = code || ''; throw e; }

function route(req) {
  if (req.action === 'login') return login(req);
  var emp = authenticate(req.token);
  switch (req.action) {
    case 'status': return status(emp, req);
    case 'punch': return withLock(function () { return punch(emp, req); });
    case 'changePin': return changePin(emp, req);
    case 'requestCorrection': return withLock(function () { return requestCorrection(emp, req); });
    case 'adminRequests': requireAdmin(emp); return adminRequests(req);
    case 'adminDecide': requireAdmin(emp); return withLock(function () { return adminDecide(emp, req); });
    case 'adminMonth': requireAdmin(emp); return adminMonth(req);
    case 'adminResetPin': requireAdmin(emp); return withLock(function () { return adminResetPin(req); });
    default: fail('不明な操作です');
  }
}

function withLock(fn) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) fail('ほかの処理中です。もう一度お試しください');
  try { return fn(); } finally { lock.releaseLock(); }
}

/* ---------- スプレッドシート ---------- */
function book() {
  var id = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  return id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActiveSpreadsheet();
}

function load(name) {
  var sh = book().getSheetByName(name);
  if (!sh) fail('シート「' + name + '」がありません。setup() を実行してください');
  var v = sh.getDataRange().getValues();
  return { sh: sh, rows: v.slice(1) };
}

function setRow(sh, rowNo, arr) { sh.getRange(rowNo, 1, 1, arr.length).setValues([arr]); }

function setup() {
  var b = book();
  [['EMP', SH.EMP], ['PUNCH', SH.PUNCH], ['REQ', SH.REQ]].forEach(function (p) {
    var sh = b.getSheetByName(p[1]) || b.insertSheet(p[1]);
    var h = HEAD[p[0]];
    if (sh.getLastRow() === 0) sh.getRange(1, 1, 1, h.length).setValues([h]).setFontWeight('bold').setBackground('#d9ece4');
    sh.setFrozenRows(1);
  });
  // 日時列は文字列固定 (スプレッドシートの自動日付変換を避ける)
  b.getSheetByName(SH.PUNCH).getRange('F:G').setNumberFormat('@');
  b.getSheetByName(SH.PUNCH).getRange('L:M').setNumberFormat('@');
  b.getSheetByName(SH.REQ).getRange('B:B').setNumberFormat('@');
  b.getSheetByName(SH.REQ).getRange('F:I').setNumberFormat('@');
  b.getSheetByName(SH.REQ).getRange('N:N').setNumberFormat('@');
  b.getSheetByName(SH.EMP).getRange('K:L').setNumberFormat('@');
  var conf = b.getSheetByName(SH.CONF) || b.insertSheet(SH.CONF);
  if (conf.getLastRow() === 0) {
    conf.getRange(1, 1, 8, 3).setValues([
      ['項目', '値', '説明'],
      ['日勤休憩分', 60, '日勤の休憩控除(分)'],
      ['夜勤休憩分', 60, '夜勤の休憩控除(分)。就業規則に合わせて変更してください'],
      ['直行直帰休憩分', 60, '直行直帰の休憩控除(分)'],
      ['休憩控除の最低勤務分', 360, 'この分数以下の勤務では休憩を控除しない'],
      ['勤務先', '本部,西九条,九条,酉島,新郡山,春日出,出来島', 'カンマ区切り'],
      ['打刻の有効時間', 24, '出勤から何時間までを同一勤務として退勤打刻できるか'],
      ['', '', '']
    ]);
    conf.getRange(1, 1, 1, 3).setFontWeight('bold').setBackground('#d9ece4');
  }
  var props = PropertiesService.getScriptProperties();
  if (!props.getProperty('SECRET')) props.setProperty('SECRET', Utilities.getUuid() + Utilities.getUuid());
  if (!props.getProperty('ADMIN_ID')) props.setProperty('ADMIN_ID', '11'); // 岩久保 博一 (名簿の番号)
}

function getConf() {
  var t = load(SH.CONF), m = {};
  t.rows.forEach(function (r) { if (r[0]) m[r[0]] = r[1]; });
  return {
    breaks: { '日勤': num(m['日勤休憩分'], 60), '夜勤': num(m['夜勤休憩分'], 60), '直行直帰': num(m['直行直帰休憩分'], 60) },
    minWork: num(m['休憩控除の最低勤務分'], 360),
    sites: String(m['勤務先'] || '本部').split(/[,、]/).map(function (s) { return s.trim(); }).filter(String),
    maxHours: num(m['打刻の有効時間'], 24)
  };
}
function num(v, d) { var n = Number(v); return isFinite(n) && v !== '' ? n : d; }

/* ---------- 時刻 ---------- */
function nowStr() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss'); }
function toDate(s) { return new Date(s.replace(' ', 'T') + '+09:00'); }
function normTime(s) {
  s = String(s || '').trim().replace('T', ' ');
  if (/^\d{4}-\d{2}-\d{2} \d{2}:\d{2}$/.test(s)) s += ':00';
  if (!/^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}$/.test(s) || isNaN(toDate(s).getTime())) fail('日時の形式が正しくありません');
  return s;
}
function minutesBetween(a, b) { return Math.round((toDate(b) - toDate(a)) / 60000); }
function calc(inAt, outAt, shift, conf) {
  var el = minutesBetween(inAt, outAt);
  var brk = el > conf.minWork ? (conf.breaks[shift] || 0) : 0;
  return { brk: brk, work: Math.max(0, el - brk) };
}

/* ---------- 従業員・認証 ---------- */
function empFromRow(r, i) {
  return { row: i + 2, id: String(r[0]), name: r[1], kana: r[2], type: r[3], site: r[4], active: r[5] === true || String(r[5]).toUpperCase() === 'TRUE',
    hash: r[6], salt: r[7], mustChange: r[8] === true || String(r[8]).toUpperCase() === 'TRUE', fails: Number(r[9]) || 0, lockUntil: r[10] || '' };
}
function findEmp(id) {
  id = String(id || '').trim();
  var t = load(SH.EMP);
  for (var i = 0; i < t.rows.length; i++) if (String(t.rows[i][0]) === id) { var e = empFromRow(t.rows[i], i); e.sh = t.sh; e.raw = t.rows[i]; return e; }
  return null;
}
function hashPin(pin, salt) {
  var pepper = PropertiesService.getScriptProperties().getProperty('SECRET');
  var d = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + ':' + pin + ':' + pepper, Utilities.Charset.UTF_8);
  return d.map(function (b) { return ('0' + (b & 0xff).toString(16)).slice(-2); }).join('');
}
function sign(payload) {
  var secret = PropertiesService.getScriptProperties().getProperty('SECRET');
  return Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(payload, secret));
}
function makeToken(e) {
  var exp = Date.now() + TOKEN_DAYS * 86400000;
  var payload = e.id + '|' + exp + '|' + String(e.hash).slice(0, 8);
  return Utilities.base64EncodeWebSafe(payload) + '.' + sign(payload);
}
function authenticate(token) {
  var bad = function () { fail('ログインの有効期限が切れました。もう一度ログインしてください', 'AUTH'); };
  if (!token || token.indexOf('.') < 0) bad();
  var parts = token.split('.');
  var payload;
  try { payload = Utilities.newBlob(Utilities.base64DecodeWebSafe(parts[0])).getDataAsString(); } catch (x) { bad(); }
  if (sign(payload) !== parts[1]) bad();
  var p = payload.split('|');
  if (Number(p[1]) < Date.now()) bad();
  var e = findEmp(p[0]);
  if (!e || !e.active || String(e.hash).slice(0, 8) !== p[2]) bad();
  return e;
}
function isAdmin(e) { return e.id === String(PropertiesService.getScriptProperties().getProperty('ADMIN_ID')); }
function requireAdmin(e) { if (!isAdmin(e)) fail('管理者のみ操作できます'); }

function login(req) {
  var generic = '従業員番号またはPINが違います';
  var e = findEmp(req.empId);
  if (!e || !e.active || !e.hash) fail(generic);
  if (e.lockUntil && toDate(e.lockUntil) > new Date()) fail('ロック中です。しばらくしてからお試しください');
  if (hashPin(String(req.pin || ''), e.salt) !== e.hash) {
    var fails = e.fails + 1, lock = '';
    if (fails >= MAX_FAIL) { lock = Utilities.formatDate(new Date(Date.now() + LOCK_MIN * 60000), TZ, 'yyyy-MM-dd HH:mm:ss'); fails = 0; }
    e.sh.getRange(e.row, 10, 1, 2).setValues([[fails, lock]]);
    fail(lock ? MAX_FAIL + '回間違えたため' + LOCK_MIN + '分間ロックしました' : generic);
  }
  e.sh.getRange(e.row, 10, 1, 3).setValues([[0, '', nowStr()]]);
  var conf = getConf();
  return { ok: true, token: makeToken(e), user: { id: e.id, name: e.name, admin: isAdmin(e), mustChange: e.mustChange }, conf: { sites: conf.sites, shifts: SHIFTS } };
}

function validPin(pin) {
  if (!/^\d{6}$/.test(pin)) fail('PINは数字6桁で入力してください');
  if (/^(\d)\1{5}$/.test(pin) || '0123456789012'.indexOf(pin) >= 0 || '9876543210987'.indexOf(pin) >= 0) fail('同じ数字や連番は使えません');
}
function changePin(e, req) {
  if (hashPin(String(req.oldPin || ''), e.salt) !== e.hash) fail('現在のPINが違います');
  validPin(String(req.newPin || ''));
  if (req.newPin === req.oldPin) fail('現在と同じPINは使えません');
  var salt = Utilities.getUuid();
  e.sh.getRange(e.row, 7, 1, 3).setValues([[hashPin(req.newPin, salt), salt, false]]);
  var fresh = findEmp(e.id);
  return { ok: true, token: makeToken(fresh) };
}

function randomPin() {
  for (;;) { var p = ('00000' + Math.floor(Math.random() * 1000000)).slice(-6); try { validPin(p); return p; } catch (x) {} }
}
/** PIN未設定の在籍者に初期PINを発行し「初期PIN配布」シートに書く。配布後はそのシートを削除すること。 */
function issueInitialPins() {
  var t = load(SH.EMP), out = [];
  t.rows.forEach(function (r, i) {
    var e = empFromRow(r, i);
    if (!e.active || e.hash) return;
    var pin = randomPin(), salt = Utilities.getUuid();
    t.sh.getRange(e.row, 7, 1, 3).setValues([[hashPin(pin, salt), salt, true]]);
    out.push([e.id, e.name, pin]);
  });
  var b = book(), sh = b.getSheetByName(SH.PIN) || b.insertSheet(SH.PIN);
  sh.getRange(1, 1, 1, 3).setValues([['従業員ID', '氏名', '初期PIN(配布後このシートを削除)']]);
  if (out.length) sh.getRange(sh.getLastRow() + 1, 1, out.length, 3).setNumberFormat('@').setValues(out);
  return out.length;
}
function adminResetPin(req) {
  var e = findEmp(req.empId);
  if (!e) fail('従業員が見つかりません');
  var pin = randomPin(), salt = Utilities.getUuid();
  e.sh.getRange(e.row, 7, 1, 3).setValues([[hashPin(pin, salt), salt, true]]);
  e.sh.getRange(e.row, 10, 1, 2).setValues([[0, '']]);
  return { ok: true, name: e.name, pin: pin };
}

/* ---------- 打刻 ---------- */
function recFromRow(r, i) {
  return { row: i + 2, id: r[0], empId: String(r[1]), name: r[2], shift: r[3], site: r[4], inAt: r[5], outAt: r[6], brk: r[7], work: r[8], state: r[9], note: r[10] };
}
function pub(r) { return { id: r.id, empId: r.empId, name: r.name, shift: r.shift, site: r.site, inAt: r.inAt, outAt: r.outAt, brk: r.brk, work: r.work, state: r.state, note: r.note }; }
function recId() { return 'R' + Utilities.formatDate(new Date(), TZ, 'yyyyMMddHHmmss') + Math.floor(Math.random() * 1000); }

function myRecords(empId) {
  var t = load(SH.PUNCH), list = [];
  t.rows.forEach(function (r, i) { if (String(r[1]) === empId) list.push(recFromRow(r, i)); });
  return { sh: t.sh, list: list };
}
function openRecord(list) {
  for (var i = list.length - 1; i >= 0; i--) if (list[i].state === '勤務中') return list[i];
  return null;
}

function punch(e, req) {
  var conf = getConf(), now = nowStr();
  var m = myRecords(e.id), open = openRecord(m.list);
  if (open && minutesBetween(open.inAt, now) > conf.maxHours * 60) {
    m.sh.getRange(open.row, 10, 1, 1).setValue('打刻漏れ');
    m.sh.getRange(open.row, 13).setValue(now);
    open.state = '打刻漏れ';
    if (req.type === 'out') fail('前回の出勤から' + conf.maxHours + '時間以上経っています。修正申請をしてください');
    open = null;
  }
  if (req.type === 'in') {
    if (open) fail('すでに出勤中です');
    if (SHIFTS.indexOf(req.shift) < 0) fail('勤務区分を選んでください');
    if (conf.sites.indexOf(req.site) < 0) fail('勤務先を選んでください');
    var row = [recId(), e.id, e.name, req.shift, req.site, now, '', '', '', '勤務中', String(req.note || '').slice(0, 100), now, now];
    m.sh.appendRow(row);
    return { ok: true, at: now, record: pub(recFromRow(row, 0)) };
  }
  if (req.type === 'out') {
    if (!open) fail('出勤の記録がありません');
    var c = calc(open.inAt, now, open.shift, conf);
    m.sh.getRange(open.row, 7, 1, 4).setValues([[now, c.brk, c.work, '完了']]);
    m.sh.getRange(open.row, 13).setValue(now);
    return { ok: true, at: now, work: c.work };
  }
  fail('不明な打刻です');
}

function status(e, req) {
  var month = /^\d{4}-\d{2}$/.test(req.month || '') ? req.month : nowStr().slice(0, 7);
  var m = myRecords(e.id), open = openRecord(m.list);
  var recs = m.list.filter(function (r) { return String(r.inAt).slice(0, 7) === month; }).map(pub).reverse();
  var rq = load(SH.REQ), reqs = [];
  rq.rows.forEach(function (r) { if (String(r[2]) === e.id) reqs.push(reqObj(r)); });
  return { ok: true, month: month, now: nowStr(), open: open ? pub(open) : null, records: recs, requests: reqs.reverse().slice(0, 30) };
}

/* ---------- 修正申請 ---------- */
function reqObj(r) {
  return { id: r[0], at: r[1], empId: String(r[2]), name: r[3], recordId: r[4], oldIn: r[5], oldOut: r[6], newIn: r[7], newOut: r[8], newShift: r[9], reason: r[10], state: r[11], decidedBy: r[12], decidedAt: r[13], comment: r[14] };
}

function requestCorrection(e, req) {
  var conf = getConf(), now = nowStr();
  var newIn = normTime(req.newIn), newOut = normTime(req.newOut);
  if (minutesBetween(newIn, newOut) <= 0) fail('退勤は出勤より後にしてください');
  if (minutesBetween(newIn, newOut) > conf.maxHours * 60) fail('勤務が' + conf.maxHours + '時間を超えています');
  if (minutesBetween(newOut, now) < 0) fail('未来の日時は申請できません');
  var reason = String(req.reason || '').trim();
  if (!reason) fail('修正の理由を入力してください');
  if (reason.length > 200) fail('理由は200文字以内で入力してください');
  var shift = req.shift;
  var oldIn = '', oldOut = '', recordId = String(req.recordId || '');
  if (recordId) {
    var rec = null;
    myRecords(e.id).list.forEach(function (r) { if (r.id === recordId) rec = r; });
    if (!rec) fail('対象の打刻が見つかりません');
    oldIn = rec.inAt; oldOut = rec.outAt; shift = shift || rec.shift;
  }
  if (SHIFTS.indexOf(shift) < 0) fail('勤務区分を選んでください');
  var t = load(SH.REQ);
  for (var i = 0; i < t.rows.length; i++) {
    var r = t.rows[i];
    if (String(r[2]) === e.id && r[11] === '申請中' && (recordId ? r[4] === recordId : !r[4] && r[7] === newIn)) fail('同じ内容の申請がすでにあります');
  }
  var row = ['Q' + Utilities.formatDate(new Date(), TZ, 'yyyyMMddHHmmss') + Math.floor(Math.random() * 1000), now, e.id, e.name, recordId, oldIn, oldOut, newIn, newOut, shift, reason, '申請中', '', '', ''];
  t.sh.appendRow(row);
  return { ok: true, request: reqObj(row) };
}

function adminRequests(req) {
  var t = load(SH.REQ), want = req.state || '申請中', list = [];
  t.rows.forEach(function (r) { if (!want || r[11] === want) list.push(reqObj(r)); });
  return { ok: true, requests: list.reverse().slice(0, 200) };
}

function adminDecide(admin, req) {
  var decision = req.decision;
  if (decision !== '承認' && decision !== '却下') fail('承認か却下を選んでください');
  var comment = String(req.comment || '').trim().slice(0, 200);
  if (decision === '却下' && !comment) fail('却下の理由を入力してください');
  var t = load(SH.REQ), idx = -1;
  t.rows.forEach(function (r, i) { if (r[0] === req.requestId) idx = i; });
  if (idx < 0) fail('申請が見つかりません');
  var q = reqObj(t.rows[idx]);
  if (q.state !== '申請中') fail('この申請はすでに処理済みです');
  var now = nowStr(), conf = getConf();
  if (decision === '承認') {
    var p = load(SH.PUNCH), c = calc(q.newIn, q.newOut, q.newShift, conf);
    if (q.recordId) {
      var found = -1;
      p.rows.forEach(function (r, i) { if (r[0] === q.recordId) found = i; });
      if (found < 0) fail('対象の打刻が見つかりません');
      var cur = p.rows[found].slice();
      cur[3] = q.newShift; cur[5] = q.newIn; cur[6] = q.newOut; cur[7] = c.brk; cur[8] = c.work; cur[9] = '修正済';
      cur[10] = (cur[10] ? cur[10] + ' / ' : '') + '修正承認 ' + q.id; cur[12] = now;
      setRow(p.sh, found + 2, cur);
    } else {
      p.sh.appendRow([recId(), q.empId, q.name, q.newShift, '', q.newIn, q.newOut, c.brk, c.work, '修正済', '追加承認 ' + q.id, now, now]);
    }
  }
  t.sh.getRange(idx + 2, 12, 1, 4).setValues([[decision === '承認' ? '承認' : '却下', admin.name, now, comment]]);
  return { ok: true };
}

/* ---------- 管理者 ---------- */
function adminMonth(req) {
  var month = /^\d{4}-\d{2}$/.test(req.month || '') ? req.month : nowStr().slice(0, 7);
  var p = load(SH.PUNCH), recs = [];
  p.rows.forEach(function (r, i) { if (String(r[5]).slice(0, 7) === month) recs.push(pub(recFromRow(r, i))); });
  var et = load(SH.EMP), emps = et.rows.map(function (r, i) {
    var e = empFromRow(r, i);
    return { id: e.id, name: e.name, type: e.type, site: e.site, active: e.active, pinSet: !!e.hash, mustChange: e.mustChange };
  });
  return { ok: true, month: month, records: recs.reverse(), employees: emps };
}
