/**
 * タスク管理アプリ「タスク実績」読取用 Web アプリ（日報アプリの「Folioから取込」用）
 *
 * 設置方法:
 *   1. https://script.google.com/ で「新しいプロジェクト」を作成
 *      （タスク管理アプリのスプレッドシートに紐付いた Folio 用 GAS とは別プロジェクトにすること。
 *        同じプロジェクトに入れると doGet が衝突する）
 *   2. このファイルの内容を貼り付けて保存
 *   3. 「デプロイ」→「新しいデプロイ」→ 種類「ウェブアプリ」
 *        次のユーザーとして実行: 自分
 *        アクセスできるユーザー: horizon.co.jp 内の全員（組織内）
 *   4. 発行された URL（.../exec）を日報アプリの taskSheetGasUrl に設定
 *   5. エディタで関数 setupTrigger を1回実行（5分おきにキャッシュを作り直すトリガーを作成）
 *
 * コード更新時は「デプロイ」→「デプロイを管理」→ 既存デプロイを編集 → バージョン「新バージョン」で
 * 更新すること（新しいデプロイを作るとURLが変わる）。
 *
 * リクエスト: GET ?action=getCompletedTasks&date=YYYY-MM-DD&name=苗字[&email=...][&nocache=1][&diag=1][&callback=fn]
 * 応答: {ok:true, rows:[{taskId, taskName, hours, start:"HH:MM"|null, startDate, doneDate, note}],
 *        ms, cached, builtAt, t:{open, read, build}}
 *
 * 作業日の判定: 開始予定日時(R)の日付。空なら作業完了日(P)。
 *   Folio(チーム時間割)はR列の日付でタスクを配置し、Folioで完了にしても作業完了日が入らない場合があるため。
 *
 * 高速化: ブックを開いて全行を読むのに数秒かかるため、「完了」行を作業日ごとにまとめて
 *   CacheService に保存し、通常はキャッシュから返す（5分おきのトリガーで作り直し）。
 *
 * 日報中分類: タスクマスタの「日報中分類CD」列を正とし、各行に cdSub として付ける。
 *   ?action=setTaskCd&taskId=K29&cd=K1          … 1件書き込み（列が無ければ末尾に追加）
 *   ?action=setTaskCds&map={"K29":"K1",...}     … 一括書き込み（空欄のタスクだけ。既存値は上書きしない）
 */

var SPREADSHEET_ID = '1IHxotYypyQkGyskunDMrN2i_v_GU2brQvGlZeAS0UgM';
var SHEET_NAME = 'タスク実績';
var MASTER_SHEET = 'タスクマスタ';
var MASTER_ID_HEADER = 'タスクID';
var MASTER_CD_HEADER = '日報中分類CD';   // タスクID→日報の中分類CD（日報アプリ取込用。正はこの列）

// 列番号（1始まり）: タスク実績シートの見出しに合わせる
var COL = {
  email:    4,  // D 担当者メールアドレス
  name:     5,  // E 担当者氏名
  taskId:   6,  // F タスクID
  taskName: 7,  // G タスク名
  est:     10,  // J 工数
  status:  11,  // K 進捗ステータス
  note:    12,  // L 備考
  actual:  15,  // O 実績工数
  doneDate:16,  // P 作業完了日
  start:   18   // R 開始予定日時
};

var CACHE_DAYS_BACK = 120;   // キャッシュ対象: 今日の120日前〜
var CACHE_DAYS_AHEAD = 14;   //               〜14日後
var CACHE_TTL = 21600;       // 6時間（CacheServiceの上限）

function doGet(e) {
  var p = (e && e.parameter) || {};
  var result;
  try {
    var action = p.action || 'getCompletedTasks';
    if (action === 'test') result = {ok: true};
    else if (action === 'getCompletedTasks') result = getCompletedTasks_(p.date, p.name, p.email, p.nocache === '1', p.diag === '1');
    else if (action === 'setTaskCd') result = setTaskCds_(singleMap_(p.taskId, p.cd), true);
    else if (action === 'setTaskCds') result = setTaskCds_(JSON.parse(p.map || '{}'), false);
    else result = {ok: false, error: 'unknown action: ' + action};
  } catch (err) {
    result = {ok: false, error: String(err && err.message || err)};
  }
  var json = JSON.stringify(result);
  if (p.callback && /^[A-Za-z0-9_$.]+$/.test(p.callback)) {
    return ContentService.createTextOutput(p.callback + '(' + json + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

function getCompletedTasks_(date, name, email, nocache, diag) {
  var t0 = Date.now();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date || '')) throw new Error('date は YYYY-MM-DD で指定してください');
  name = String(name || '').trim();
  email = String(email || '').trim().toLowerCase();
  if (!name && !email) throw new Error('name または email を指定してください');

  var cache = CacheService.getScriptCache();
  var list = null, cached = false, builtAt = null, t = null, diagInfo = null;
  if (!nocache && !diag) {
    var got = cache.getAll(['meta', 'd:' + date]);
    if (got['meta'] && got['d:' + date] != null) {
      list = JSON.parse(got['d:' + date]);
      builtAt = JSON.parse(got['meta']).builtAt;
      cached = true;
    }
  }
  if (!list) {
    var idx = buildIndex_(diag ? {name: name, email: email} : null);
    t = idx.t; builtAt = idx.builtAt; diagInfo = idx.diag;
    list = idx.byDate[date] || [];
  }

  var tm = getTaskMap_();
  var rows = list.filter(function (x) {
    return (email && x.e === email) || (name && x.n.indexOf(name) === 0);
  }).map(function (x) {
    return {taskId: x.id, taskName: x.nm, cdSub: tm[x.id] || null, hours: x.h, start: x.s, startDate: x.sd, doneDate: x.dd, note: x.note};
  });
  var res = {ok: true, rows: rows, ms: Date.now() - t0, cached: cached, builtAt: builtAt};
  if (t) res.t = t;
  if (diagInfo) res.diag = diagInfo;
  return res;
}

// シートを読み、「完了」行を作業日ごとにまとめてキャッシュへ保存する。
function buildIndex_(diagFor) {
  var t0 = Date.now();
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) throw new Error('シート「' + SHEET_NAME + '」が見つかりません');
  var tz = ss.getSpreadsheetTimeZone();
  var t1 = Date.now();

  var byDate = {};
  var last = sh.getLastRow();
  var n = Math.max(0, last - 1);
  var vals = [], estDisp = [], actDisp = [];
  if (n > 0) {
    var c0 = COL.email;
    vals = sh.getRange(2, c0, n, COL.start - c0 + 1).getValues();
    // 時間値は表示値("1:30:00")で読む（Date換算の誤差回避）。必要な2列だけ。
    estDisp = sh.getRange(2, COL.est, n, 1).getDisplayValues();
    actDisp = sh.getRange(2, COL.actual, n, 1).getDisplayValues();
  }
  var t2 = Date.now();

  var today = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  var lo = shiftYmd_(today, -CACHE_DAYS_BACK), hi = shiftYmd_(today, CACHE_DAYS_AHEAD);
  var diag = diagFor ? [] : null;
  var ix = function (col) { return col - COL.email; };

  for (var i = 0; i < n; i++) {
    var r = vals[i];
    var rName = String(r[ix(COL.name)]).trim();
    var rEmail = String(r[ix(COL.email)]).trim().toLowerCase();
    var startV = r[ix(COL.start)], doneV = r[ix(COL.doneDate)];
    var sd = toYmd_(startV, tz), dd = toYmd_(doneV, tz);
    var status = String(r[ix(COL.status)]).trim();
    if (diag && diag.length < 3 &&
        ((diagFor.email && rEmail === diagFor.email) || (diagFor.name && rName.indexOf(diagFor.name) === 0)) &&
        (startV !== '' || doneV !== '')) {
      diag.push({row: i + 2, status: status,
                 startType: Object.prototype.toString.call(startV), startRaw: String(startV), startYmd: sd, startHm: toHm_(startV, tz),
                 doneType: Object.prototype.toString.call(doneV), doneRaw: String(doneV), doneYmd: dd});
    }
    if (status !== '完了') continue;
    var wd = sd || dd;
    if (!wd) continue;
    var hours = durToHours_(actDisp[i] && actDisp[i][0]);
    if (hours == null) hours = durToHours_(estDisp[i] && estDisp[i][0]) || 0;
    (byDate[wd] = byDate[wd] || []).push({
      e: rEmail, n: rName,
      id: String(r[ix(COL.taskId)]).trim(),
      nm: String(r[ix(COL.taskName)]).trim(),
      h: Math.round(hours * 10000) / 10000,
      s: toHm_(startV, tz), sd: sd, dd: dd,
      note: String(r[ix(COL.note)] || '').trim()
    });
  }

  // キャッシュへ保存（範囲内の日付はデータが無くても [] を入れて「キャッシュ済み」を表す）
  var builtAt = new Date().toISOString();
  var put = {};
  for (var d = lo; d <= hi; d = shiftYmd_(d, 1)) put['d:' + d] = JSON.stringify(byDate[d] || []);
  try {
    var keys = Object.keys(put), cache = CacheService.getScriptCache();
    for (var k = 0; k < keys.length; k += 50) {
      var chunk = {};
      keys.slice(k, k + 50).forEach(function (key) { chunk[key] = put[key]; });
      cache.putAll(chunk, CACHE_TTL);
    }
    cache.put('tm', JSON.stringify(readTaskMap_(ss)), CACHE_TTL);
    cache.put('meta', JSON.stringify({builtAt: builtAt}), CACHE_TTL);
  } catch (err) {
    // 1キー100KB超などでキャッシュできなくても、今回の応答は返す
  }
  var t3 = Date.now();
  return {byDate: byDate, builtAt: builtAt, diag: diag, t: {open: t1 - t0, read: t2 - t1, build: t3 - t2}};
}

// ── タスクマスタ「日報中分類CD」列 ──
function masterCols_(sh) {
  var lastCol = Math.max(1, sh.getLastColumn());
  var head = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(function (v) { return String(v).trim(); });
  return {idCol: head.indexOf(MASTER_ID_HEADER) + 1, cdCol: head.indexOf(MASTER_CD_HEADER) + 1, lastCol: lastCol};
}

function readTaskMap_(ss) {
  var sh = ss.getSheetByName(MASTER_SHEET);
  if (!sh) return {};
  var c = masterCols_(sh);
  var n = sh.getLastRow() - 1;
  if (!c.idCol || !c.cdCol || n < 1) return {};
  var ids = sh.getRange(2, c.idCol, n, 1).getValues(), cds = sh.getRange(2, c.cdCol, n, 1).getValues();
  var map = {};
  for (var i = 0; i < n; i++) {
    var id = String(ids[i][0]).trim(), cd = String(cds[i][0]).trim();
    if (id && cd) map[id] = cd;
  }
  return map;
}

// キャッシュ優先でタスクマスタの対応表を返す
function getTaskMap_() {
  var cache = CacheService.getScriptCache();
  var raw = cache.get('tm');
  if (raw) return JSON.parse(raw);
  var map = readTaskMap_(SpreadsheetApp.openById(SPREADSHEET_ID));
  cache.put('tm', JSON.stringify(map), CACHE_TTL);
  return map;
}

function singleMap_(taskId, cd) {
  var m = {};
  m[String(taskId || '').trim()] = String(cd || '').trim();
  return m;
}

// 対応を書き込む。overwrite=false なら空欄のタスクだけ書く。
function setTaskCds_(map, overwrite) {
  var keys = Object.keys(map || {});
  if (!keys.length) throw new Error('taskId と cd を指定してください');
  keys.forEach(function (k) {
    if (!k) throw new Error('taskId が空です');
    if (!/^[A-Za-z0-9α-ω]{1,6}$/.test(String(map[k]))) throw new Error('cd の形式が不正です: ' + map[k]);
  });
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sh = ss.getSheetByName(MASTER_SHEET);
    if (!sh) throw new Error('シート「' + MASTER_SHEET + '」が見つかりません');
    var c = masterCols_(sh);
    if (!c.idCol) throw new Error('タスクマスタに「' + MASTER_ID_HEADER + '」列がありません');
    if (!c.cdCol) {                                   // 列が無ければ末尾に見出し付きで追加
      c.cdCol = c.lastCol + 1;
      if (c.cdCol > sh.getMaxColumns()) sh.insertColumnAfter(sh.getMaxColumns());
      sh.getRange(1, c.cdCol).setValue(MASTER_CD_HEADER);
    }
    var n = sh.getLastRow() - 1;
    var ids = n > 0 ? sh.getRange(2, c.idCol, n, 1).getValues() : [];
    var cdRange = n > 0 ? sh.getRange(2, c.cdCol, n, 1) : null;
    var cds = cdRange ? cdRange.getValues() : [];
    var rowOf = {};
    for (var i = 0; i < n; i++) { var id = String(ids[i][0]).trim(); if (id && !(id in rowOf)) rowOf[id] = i; }
    var written = [], skippedExisting = [], notFound = [];
    keys.forEach(function (k) {
      if (!(k in rowOf)) { notFound.push(k); return; }
      var cur = String(cds[rowOf[k]][0]).trim();
      if (cur && !overwrite) { skippedExisting.push(k); return; }
      cds[rowOf[k]][0] = String(map[k]).trim();
      written.push(k);
    });
    if (overwrite && notFound.length) throw new Error('タスクマスタに存在しないタスクIDです: ' + notFound.join(', '));
    if (written.length) cdRange.setValues(cds);
    var tm = readTaskMap_(ss);
    CacheService.getScriptCache().put('tm', JSON.stringify(tm), CACHE_TTL);
    return {ok: true, written: written.length, skippedExisting: skippedExisting.length, notFound: notFound};
  } finally {
    lock.releaseLock();
  }
}

// 時間主導トリガー用: キャッシュを作り直す
function warmCache() {
  buildIndex_(null);
}

// エディタから1回だけ実行: 5分おきに warmCache を動かすトリガーを作成（重複は作らない）
function setupTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (tr) {
    if (tr.getHandlerFunction() === 'warmCache') ScriptApp.deleteTrigger(tr);
  });
  ScriptApp.newTrigger('warmCache').timeBased().everyMinutes(5).create();
  warmCache();
}

function isDate_(v) {
  return Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v.getTime());
}

// セル値(Date または文字列) → "YYYY-MM-DD"
function toYmd_(v, tz) {
  if (isDate_(v)) return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
  return normDate_(v);
}

// セル値(Date または文字列) → "HH:MM"（時刻なし/00:00 は null）
function toHm_(v, tz) {
  if (isDate_(v)) {
    var hm = Utilities.formatDate(v, tz, 'HH:mm');
    return hm === '00:00' ? null : hm;
  }
  return normTime_(v);
}

// "YYYY-MM-DD" を days 日ずらす
function shiftYmd_(ymd, days) {
  var p = ymd.split('-');
  var d = new Date(Date.UTC(+p[0], +p[1] - 1, +p[2] + days));
  return d.getUTCFullYear() + '-' + ('0' + (d.getUTCMonth() + 1)).slice(-2) + '-' + ('0' + d.getUTCDate()).slice(-2);
}

// "2026/10/1" "2026-10-01 9:00:00" → "2026-10-01"
function normDate_(s) {
  var m = String(s || '').match(/(\d{4})[\/\-.](\d{1,2})[\/\-.](\d{1,2})/);
  if (!m) return null;
  return m[1] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[3]).slice(-2);
}

// "2026/10/01 9:00:00" → "09:00"（日付の後ろの時刻部分）
function normTime_(s) {
  var m = String(s || '').match(/\d{4}[\/\-.]\d{1,2}[\/\-.]\d{1,2}\s+(\d{1,2}):(\d{2})/);
  if (!m) return null;
  return ('0' + m[1]).slice(-2) + ':' + m[2];
}

// "1:30:00" / "0:15" → 時間数。空なら null
function durToHours_(s) {
  var m = String(s || '').trim().match(/^(\d+):(\d{2})(?::(\d{2}))?$/);
  if (!m) return null;
  return Number(m[1]) + Number(m[2]) / 60 + Number(m[3] || 0) / 3600;
}
