/**
 * タスク管理アプリ「タスク実績」読取用 Web アプリ（日報アプリの「タスク実績から取込」用）
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
 *
 * リクエスト: GET ?action=getCompletedTasks&date=YYYY-MM-DD&name=苗字[&email=...][&callback=fn]
 * 応答: {ok:true, rows:[{taskId, taskName, hours, start:"HH:MM"|null, startDate:"YYYY-MM-DD"|null, doneDate, note}]}
 *
 * 作業日の判定: 開始予定日時(R)の日付。空なら作業完了日(P)。
 *   Folio(チーム時間割)はR列の日付でタスクを配置し、Folioで完了にしても作業完了日が入らない場合があるため。
 *
 * コード更新時は「デプロイ」→「デプロイを管理」→ 既存デプロイを編集 → バージョン「新バージョン」で
 * 更新すること（新しいデプロイを作るとURLが変わる）。
 */

var SPREADSHEET_ID = '1IHxotYypyQkGyskunDMrN2i_v_GU2brQvGlZeAS0UgM';
var SHEET_NAME = 'タスク実績';

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

function doGet(e) {
  var p = (e && e.parameter) || {};
  var result;
  try {
    var action = p.action || 'getCompletedTasks';
    if (action === 'test') result = {ok: true};
    else if (action === 'getCompletedTasks') result = getCompletedTasks_(p.date, p.name, p.email);
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

function getCompletedTasks_(date, name, email) {
  var t0 = Date.now();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date || '')) throw new Error('date は YYYY-MM-DD で指定してください');
  name = String(name || '').trim();
  email = String(email || '').trim().toLowerCase();
  if (!name && !email) throw new Error('name または email を指定してください');

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) throw new Error('シート「' + SHEET_NAME + '」が見つかりません');
  var tz = ss.getSpreadsheetTimeZone();
  var last = sh.getLastRow();
  if (last < 2) return {ok: true, rows: [], ms: Date.now() - t0};
  var n = last - 1;

  // 高速化: 全列の表示値(getDisplayValues)は遅いので、判定に使う列だけ getValues で読む。
  // 時間値(工数/実績工数)は Date 変換の誤差を避けるため、該当行だけ表示値で読む。
  var c0 = COL.email;                                  // D列から
  var vals = sh.getRange(2, c0, n, COL.start - c0 + 1).getValues();
  var idx = function (col) { return col - c0; };

  var hitRows = [];
  for (var i = 0; i < n; i++) {
    var r = vals[i];
    if (String(r[idx(COL.status)]).trim() !== '完了') continue;
    var rName = String(r[idx(COL.name)]).trim();
    var rEmail = String(r[idx(COL.email)]).trim().toLowerCase();
    if (!((email && rEmail === email) || (name && rName.indexOf(name) === 0))) continue;
    var startD = toYmd_(r[idx(COL.start)], tz);
    var doneD = toYmd_(r[idx(COL.doneDate)], tz);
    if ((startD || doneD) !== date) continue;
    hitRows.push({i: i, startD: startD, doneD: doneD});
  }

  var rows = hitRows.map(function (h) {
    var r = vals[h.i];
    // 該当行の J〜O 列だけ表示値で読む（件数が少ないので速い）
    var disp = sh.getRange(h.i + 2, COL.est, 1, COL.actual - COL.est + 1).getDisplayValues()[0];
    var hours = durToHours_(disp[COL.actual - COL.est]);
    if (hours == null) hours = durToHours_(disp[0]) || 0;
    return {
      taskId: String(r[idx(COL.taskId)]).trim(),
      taskName: String(r[idx(COL.taskName)]).trim(),
      hours: Math.round(hours * 10000) / 10000,
      start: toHm_(r[idx(COL.start)], tz),
      startDate: h.startD,
      doneDate: h.doneD,
      note: String(r[idx(COL.note)] || '').trim()
    };
  });
  return {ok: true, rows: rows, ms: Date.now() - t0};
}

// セル値(Date または文字列) → "YYYY-MM-DD"
function toYmd_(v, tz) {
  if (v instanceof Date) return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
  return normDate_(v);
}

// セル値(Date または文字列) → "HH:MM"（時刻なし/00:00 は null）
function toHm_(v, tz) {
  if (v instanceof Date) {
    var hm = Utilities.formatDate(v, tz, 'HH:mm');
    return hm === '00:00' ? null : hm;
  }
  return normTime_(v);
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
