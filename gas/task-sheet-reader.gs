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
 * 応答: {ok:true, rows:[{taskId, taskName, hours, start:"HH:MM"|null, startDate:"YYYY-MM-DD"|null, note}]}
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
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date || '')) throw new Error('date は YYYY-MM-DD で指定してください');
  name = String(name || '').trim();
  email = String(email || '').trim().toLowerCase();
  if (!name && !email) throw new Error('name または email を指定してください');

  var sh = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
  if (!sh) throw new Error('シート「' + SHEET_NAME + '」が見つかりません');
  var last = sh.getLastRow();
  if (last < 2) return {ok: true, rows: []};
  // 表示値で読む（時間値 "1:30:00" や日時を文字列として安全に解析するため）
  var values = sh.getRange(2, 1, last - 1, COL.start).getDisplayValues();

  var rows = [];
  for (var i = 0; i < values.length; i++) {
    var r = values[i];
    if (String(r[COL.status - 1]).trim() !== '完了') continue;
    if (normDate_(r[COL.doneDate - 1]) !== date) continue;
    var rEmail = String(r[COL.email - 1]).trim().toLowerCase();
    var rName = String(r[COL.name - 1]).trim();
    var hit = (email && rEmail === email) || (name && rName.indexOf(name) === 0);
    if (!hit) continue;
    var hours = durToHours_(r[COL.actual - 1]);
    if (hours == null) hours = durToHours_(r[COL.est - 1]) || 0;
    var startRaw = String(r[COL.start - 1] || '');
    rows.push({
      taskId: String(r[COL.taskId - 1]).trim(),
      taskName: String(r[COL.taskName - 1]).trim(),
      hours: Math.round(hours * 10000) / 10000,
      start: normTime_(startRaw),
      startDate: normDate_(startRaw),
      note: String(r[COL.note - 1] || '').trim()
    });
  }
  return {ok: true, rows: rows};
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
