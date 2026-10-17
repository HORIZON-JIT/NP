/**
 * 日報アプリ 月次AI分析用 Web アプリ（生産管理課のみで使用）
 *
 * 日報アプリの月次ダッシュボードで集計した値（担当別・中分類別の時間、残業、前月期、年度累計）を受け取り、
 * Claude API で分析文を作って返す。APIキーはこのGASのスクリプトプロパティにだけ置き、HTML側には出さない。
 *
 * 設置方法:
 *   1. https://script.google.com/ で「新しいプロジェクト」を作成（既存の日報GAS・Folio GASとは別プロジェクト）
 *   2. このファイルの内容を貼り付けて保存
 *   3. 「プロジェクトの設定」→「スクリプト プロパティ」に次を追加
 *        ANTHROPIC_API_KEY = sk-ant-...（Claude APIキー）
 *        CLAUDE_MODEL      = （任意。省略時は下の DEFAULT_MODEL）
 *   4. エディタで関数 checkSetup を1回実行（外部通信の許可を求められたら許可。ログに「OK」と出れば準備完了）
 *   5. 「デプロイ」→「新しいデプロイ」→ 種類「ウェブアプリ」
 *        次のユーザーとして実行: 自分
 *        アクセスできるユーザー: horizon.co.jp 内の全員（組織内）
 *   6. 発行された URL（.../exec）を日報アプリの config.js（ルート＝生産管理課）の aiAnalysisGasUrl に設定
 *
 * コード更新時は「デプロイ」→「デプロイを管理」→ 既存デプロイを編集 → バージョン「新バージョン」で更新
 * （新しいデプロイを作るとURLが変わる）。
 *
 * 通信（JSONP。callback パラメータ付きで呼ぶ）:
 *   ?action=test
 *   ?action=get&key=期間キー                         … 保存済みの分析を返す（APIは呼ばない）
 *   ?action=part&id=送信ID&i=番号&d=base64断片        … 集計データの分割送信（URL長の制限対策）
 *   ?action=analyze&id=送信ID&n=断片数&key=期間キー&hash=データハッシュ[&force=1]
 *        → 断片を結合して Claude に送り、結果を保存して返す。
 *          force なしで同じハッシュの保存結果があれば、APIを呼ばずにそれを返す。
 * 応答: {ok, found, text, at, model, hash, cached}
 */

var DEFAULT_MODEL = 'claude-sonnet-5-5';
var MAX_TOKENS = 8000;            // 思考分も含む出力上限（思考で使い切ると本文が空になるため余裕を持たせる）
var PART_TTL = 600;               // 分割送信の断片を保持する秒数
var MAX_PAYLOAD_CHARS = 60000;    // 受け付ける集計データの上限（文字数）
var MIN_INTERVAL_MS = 20 * 1000;  // 同じ期間の再分析を受け付ける最短間隔
var STORE_PREFIX = 'ai_';

function doGet(e) {
  var p = (e && e.parameter) || {};
  var result;
  try {
    var action = p.action || 'test';
    if (action === 'test') result = {ok: true, model: model_(), hasKey: !!apiKey_()};
    else if (action === 'get') result = getStored_(p.key);
    else if (action === 'part') result = putPart_(p.id, p.i, p.d);
    else if (action === 'analyze') result = analyze_(p.id, Number(p.n), p.key, p.hash, p.force === '1');
    else result = {ok: false, error: 'unknown action: ' + action};
  } catch (err) {
    result = {ok: false, error: String(err && err.message || err)};
  }
  var json = JSON.stringify(result);
  if (p.callback) {
    return ContentService.createTextOutput(p.callback + '(' + json + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

// ─── 保存済みの分析 ───
function storeKey_(key) {
  if (!key || !/^[0-9A-Za-z_\-]{1,80}$/.test(key)) throw new Error('期間キーが不正です');
  return STORE_PREFIX + key;
}
function getStored_(key) {
  var raw = PropertiesService.getScriptProperties().getProperty(storeKey_(key));
  if (!raw) return {ok: true, found: false};
  var s = JSON.parse(raw);
  return {ok: true, found: true, text: s.text, at: s.at, model: s.model, hash: s.hash, cached: true};
}

// ─── 分割送信 ───
function putPart_(id, i, d) {
  if (!id || !/^[0-9a-z]{6,40}$/.test(id)) throw new Error('送信IDが不正です');
  if (!/^\d{1,2}$/.test(String(i))) throw new Error('断片番号が不正です');
  CacheService.getScriptCache().put('aip_' + id + '_' + i, String(d || ''), PART_TTL);
  return {ok: true};
}
function joinParts_(id, n) {
  if (!id || !/^[0-9a-z]{6,40}$/.test(id)) throw new Error('送信IDが不正です');
  if (!(n >= 1 && n <= 40)) throw new Error('断片数が不正です');
  var keys = [];
  for (var i = 0; i < n; i++) keys.push('aip_' + id + '_' + i);
  var got = CacheService.getScriptCache().getAll(keys);
  var b64 = '';
  for (var j = 0; j < n; j++) {
    if (got[keys[j]] == null) throw new Error('集計データの受信が欠けています（' + (j + 1) + '/' + n + '）。もう一度お試しください');
    b64 += got[keys[j]];
  }
  var json = Utilities.newBlob(Utilities.base64Decode(b64)).getDataAsString('UTF-8');
  if (json.length > MAX_PAYLOAD_CHARS) throw new Error('集計データが大きすぎます');
  return JSON.parse(json);
}

// ─── 分析本体 ───
function analyze_(id, n, key, hash, force) {
  var sk = storeKey_(key);
  var props = PropertiesService.getScriptProperties();
  var prevRaw = props.getProperty(sk);
  var prev = prevRaw ? JSON.parse(prevRaw) : null;
  if (prev && !force && prev.hash && prev.hash === hash) {
    return {ok: true, found: true, text: prev.text, at: prev.at, model: prev.model, hash: prev.hash, cached: true};
  }
  if (prev && Date.now() - Date.parse(prev.at) < MIN_INTERVAL_MS) {
    return {ok: true, found: true, text: prev.text, at: prev.at, model: prev.model, hash: prev.hash, cached: true,
            note: '直前に分析したばかりのため、前回の結果を表示しています'};
  }
  var data = joinParts_(id, n);

  var lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) throw new Error('他の分析を実行中です。少し待ってからお試しください');
  try {
    var text = callClaude_(data);
    var rec = {text: text, at: new Date().toISOString(), model: model_(), hash: String(hash || '')};
    var s = JSON.stringify(rec);
    // スクリプトプロパティは1件9KBまで。超える場合は末尾を切り詰めて保存する
    while (Utilities.newBlob(s).getBytes().length > 8800 && rec.text.length > 200) {
      rec.text = rec.text.slice(0, Math.floor(rec.text.length * 0.9)) + '\n\n（保存容量の都合で以降を省略）';
      s = JSON.stringify(rec);
    }
    props.setProperty(sk, s);
    return {ok: true, found: true, text: rec.text, at: rec.at, model: rec.model, hash: rec.hash, cached: false};
  } finally {
    lock.releaseLock();
  }
}

function apiKey_() { return PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY') || ''; }
function model_() { return PropertiesService.getScriptProperties().getProperty('CLAUDE_MODEL') || DEFAULT_MODEL; }

function systemPrompt_() {
  return [
    'あなたは製造業（株式会社ホリゾン、製本・印刷関連機器メーカー）の生産管理課の業務分析担当です。',
    '課員の日報を月期ごとに集計したJSONを受け取り、課長・課員が月次の振り返りに使う分析レポートを日本語で書いてください。',
    '',
    '【データの前提】',
    '- 月期は前月16日〜当月15日（15日締め）。inProgress=true は期の途中で、workElapsed/workTotal が経過稼働日/期全体の稼働日。',
    '- 各担当の n（通常時間）と ot（残業時間）は退勤時刻から算出した正式な値。残業は otStart 以降。',
    '- subs は日報に入力した中分類ごとの [通常, 残業] 時間。cats で中分類コードを「大分類/中分類」名に引ける。',
    '- miss は平日の日報未入力日数。prev は前月期、fy は年度（10月期始まり）の期初から当月期までの残業累計と月45h超の回数。',
    '- 36協定の管理基準: 月 otMonthWarn h で警告、月 otMonthLimit h が上限、年 otYearLimit h が上限。',
    '- 担当「ヘルプ」は応援者の工数で、残業・稼働の評価対象外。',
    '',
    '【書き方】',
    '- Markdown。見出しは「## 」、箇条書きは「- 」のみ使う。表は使わない。',
    '- 次の見出しをこの順で出す: ## 全体概況 / ## 残業の傾向 / ## 業務構成の変化 / ## 気になる点 / ## 来月期に向けて',
    '- 数値は必ずデータから計算できるものだけを書き、小数第1位までに丸める。データに無いことは書かない。推測は「〜と考えられます」と明記する。',
    '- 前月期との差は増減の大きいものに絞る。期の途中なら、その旨と稼働日ベースの見込みに触れる。',
    '- 特定の個人を責める書き方はしない。偏りは業務の配分・時期の問題として書く。',
    '- 強調のためにかぎ括弧を使わない。です・ます調。全体で900〜1300字程度。'
  ].join('\n');
}

function callClaude_(data) {
  var key = apiKey_();
  if (!key) throw new Error('APIキーが未設定です（GASのスクリプトプロパティ ANTHROPIC_API_KEY）');
  var body = {
    model: model_(),
    max_tokens: MAX_TOKENS,
    system: systemPrompt_(),
    messages: [{role: 'user', content: '次の月期の日報集計データを分析してください。\n\n```json\n' + JSON.stringify(data) + '\n```'}]
  };
  var res = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'post',
    contentType: 'application/json',
    headers: {'x-api-key': key, 'anthropic-version': '2023-06-01'},
    payload: JSON.stringify(body),
    muteHttpExceptions: true
  });
  var code = res.getResponseCode();
  var out = {};
  try { out = JSON.parse(res.getContentText()); } catch (e) {}
  if (code !== 200) {
    var msg = (out && out.error && out.error.message) || res.getContentText().slice(0, 300);
    throw new Error('Claude API エラー（' + code + '）: ' + msg);
  }
  var text = (out.content || []).filter(function (c) { return c.type === 'text'; })
    .map(function (c) { return c.text; }).join('\n').trim();
  if (!text) {
    var types = (out.content || []).map(function (c) { return c.type; }).join(',') || 'なし';
    throw new Error('Claude API から本文が返りませんでした（stop_reason: ' + out.stop_reason + ' / 内容: ' + types + '）');
  }
  return text;
}

// ─── 設置確認（エディタから手動実行） ───
function checkSetup() {
  if (!apiKey_()) { Logger.log('NG: スクリプトプロパティ ANTHROPIC_API_KEY が未設定です'); return; }
  var res = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'post',
    contentType: 'application/json',
    headers: {'x-api-key': apiKey_(), 'anthropic-version': '2023-06-01'},
    payload: JSON.stringify({model: model_(), max_tokens: 20, messages: [{role: 'user', content: '「OK」とだけ返してください'}]}),
    muteHttpExceptions: true
  });
  Logger.log('モデル: ' + model_());
  Logger.log('HTTP ' + res.getResponseCode() + ': ' + res.getContentText().slice(0, 300));
  if (res.getResponseCode() === 200) Logger.log('OK: 準備完了です');
}

// ─── 保存済みの分析をすべて削除（エディタから手動実行） ───
function clearStoredAnalyses() {
  var props = PropertiesService.getScriptProperties();
  var all = props.getProperties(), n = 0;
  Object.keys(all).forEach(function (k) { if (k.indexOf(STORE_PREFIX) === 0) { props.deleteProperty(k); n++; } });
  Logger.log(n + '件削除しました');
}
