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
  var s = readRec_(storeKey_(key));
  if (!s) return {ok: true, found: false};
  return {ok: true, found: true, text: s.text, at: s.at, model: s.model, hash: s.hash, cached: true};
}
// スクリプトプロパティは1件9KBまでなので、長い結果は sk, sk#1, sk#2 … に分けて保存する
function readRec_(sk) {
  var props = PropertiesService.getScriptProperties();
  var raw = props.getProperty(sk);
  if (!raw) return null;
  var head = JSON.parse(raw);
  if (head.parts) {
    var txt = '';
    for (var i = 1; i <= head.parts; i++) txt += props.getProperty(sk + '#' + i) || '';
    head.text = txt;
  }
  return head;
}
function writeRec_(sk, rec) {
  var props = PropertiesService.getScriptProperties();
  var old = props.getProperty(sk);
  if (old) { try { var o = JSON.parse(old); for (var j = 1; j <= (o.parts || 0); j++) props.deleteProperty(sk + '#' + j); } catch (e) {} }
  var CH = 2500;   // 日本語3バイト換算で約7.5KB
  var parts = [];
  for (var i = 0; i < rec.text.length; i += CH) parts.push(rec.text.slice(i, i + CH));
  if (parts.length > 8) throw new Error('分析結果が長すぎて保存できません');
  var head = {at: rec.at, model: rec.model, hash: rec.hash, parts: parts.length};
  var kv = {}; kv[sk] = JSON.stringify(head);
  parts.forEach(function (t, k) { kv[sk + '#' + (k + 1)] = t; });
  props.setProperties(kv);
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
  var prev = readRec_(sk);
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
    writeRec_(sk, rec);
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
    '【出力形式】',
    '- 次の形のJSONだけを出力する。前後に説明文やコードブロック記号を付けない。',
    '{"headline":"…","status":"good|warn|bad","sections":{"overview":[…],"overtime":[…],"workmix":[…],"concerns":[…]},"notes":{"ot":"…","otcat":"…","heat":"…","mix":"…","sub":"…","fy":"…"},"actions":[{"text":"…","who":"…"}]}',
    '- headline: 月期全体を一文で総括（60字以内）。',
    '- status: 月期全体の状態。bad=月上限超過など対応が必要、warn=注意が必要、good=概ね良好。',
    '- notes: 画面の各グラフの下に添える一言（各60字以内）。そのグラフから読み取れる最も重要な点を、理由や含意まで含めて書く。',
    '  ot=担当別の残業（警告・上限との比較）、otcat=残業がどの大分類の業務で発生したか、heat=担当×大分類の時間配分（業務の集中・偏り）、',
    '  mix=大分類の構成比と前月期からの増減、sub=中分類の増減の上位、fy=年度累計の残業（fy.months が2未満なら空文字）。prev が無ければ mix と sub は前月比に触れない。',
    '- sections の各配列の要素は {"tone":"good|warn|bad|info","text":"…"}。text は1項目80字以内、各セクション2〜4項目。',
    '  overview=全体概況、overtime=残業の傾向、workmix=業務構成の変化（前月期との差）、concerns=気になる点。',
    '  tone: good=改善・良い傾向、warn=注意が必要、bad=基準超過など対応が必要、info=事実の説明。',
    '- actions: 来月期に向けた具体的な打ち手を3〜4件。who は主に関わる担当名か「課全体」。',
    '- 画面には担当別残業・残業の発生業務・担当×業務の時間配分・業務構成・中分類の増減・年度累計のグラフが別に表示される。グラフを読めば分かる数値の羅列は避け、増減の理由・偏り・注意点など解釈を中心に書く。',
    '- 数値はデータから計算できるものだけを使い、小数第1位までに丸める。データに無いことは書かない。推測は「〜と考えられます」と明記する。',
    '- 期の途中（inProgress=true）なら、その旨と稼働日ベースの見込みに触れる。',
    '- 特定の個人を責める書き方はしない。偏りは業務の配分・時期の問題として書く。',
    '- 強調のためにかぎ括弧を使わない。です・ます調。'
  ].join('\n');
}

function callClaude_(data) {
  var key = apiKey_();
  if (!key) throw new Error('APIキーが未設定です（GASのスクリプトプロパティ ANTHROPIC_API_KEY）');
  var base = {
    model: model_(),
    system: systemPrompt_(),
    messages: [{role: 'user', content: '次の月期の日報集計データを分析してください。\n\n```json\n' + JSON.stringify(data) + '\n```'}]
  };
  // 1回目: 事前の思考を止めて（Sonnet 5.5 は between_tools）出力枠をすべて本文に使う
  // 2回目: 指定が通らないモデルや枠切れのときは、思考ありのまま出力枠を広げて再試行
  var tries = [
    {max_tokens: MAX_TOKENS, thinking: {type: 'between_tools'}, output_config: {effort: 'medium'}},
    {max_tokens: 16000, output_config: {effort: 'low'}}
  ];
  var lastErr = null;
  for (var t = 0; t < tries.length; t++) {
    var body = JSON.parse(JSON.stringify(base));
    Object.keys(tries[t]).forEach(function (k) { body[k] = tries[t][k]; });
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
      lastErr = new Error('Claude API エラー（' + code + '）: ' + msg);
      if (code === 400) continue;   // パラメータ非対応 → 次の設定で再試行
      throw lastErr;
    }
    var text = (out.content || []).filter(function (c) { return c.type === 'text'; })
      .map(function (c) { return c.text; }).join('\n').trim();
    if (text && out.stop_reason !== 'max_tokens') return text;
    var types = (out.content || []).map(function (c) { return c.type; }).join(',') || 'なし';
    lastErr = new Error('Claude API から本文が最後まで返りませんでした（stop_reason: ' + out.stop_reason + ' / 内容: ' + types + '）');
  }
  throw lastErr;
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
