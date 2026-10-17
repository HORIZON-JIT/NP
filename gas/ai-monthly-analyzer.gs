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
  if (rec.by) head.by = rec.by;
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

function systemPrompt_(kind) {
  if (kind === 'yearly') return systemPromptYearly_();
  if (kind === 'kpi') return systemPromptKpi_();
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

var COMMON_RULES_ = [
  '- summary（KPI は各要素の summary）に入っている合計・累計・平均・回数・見込み・判定は計算済みの正しい値。数値を書くときは必ずここから引用し、自分で足し算や平均をしない。',
  '- 年上限や月上限を超えた担当は summary の一覧（overYearLimit など）だけを根拠に書く。名前を取り違えない。',
  '- 数値はデータから読み取れるものだけを使い、小数第1位までに丸める。データに無いことは書かない。推測は「〜と考えられます」と明記する。',
  '- 特定の個人を責める書き方はしない。偏りは業務の配分・時期の問題として書く。',
  '- 強調のためにかぎ括弧を使わない。です・ます調。',
  '- 担当「ヘルプ」は応援者の工数で、残業・稼働の評価対象外。'
];

function systemPromptYearly_() {
  return [
    'あなたは製造業（株式会社ホリゾン、製本・印刷関連機器メーカー）の生産管理課の業務分析担当です。',
    '課員の日報を年度（10月期〜翌9月期の12か月期）で集計したJSONを受け取り、課長・課員が年度の振り返りに使う分析レポートを日本語で書いてください。',
    '',
    '【データの前提】',
    '- months は月期のラベル（各月期は前月16日〜当月15日）。配列の値は months と同じ順。null は未経過またはデータなし。',
    '- persons[担当].ot / n は月期ごとの残業・通常時間（退勤時刻から算出した正式な値）。persons[担当].subs は年度累計の中分類別 [通常, 残業]（日報入力ベース）。',
    '- mainByMonth は大分類ごとの月期別の作業時間（通常＋残業）。cats で中分類コードを「大分類/中分類」名に引ける。',
    '- prevOt は前年度の同じ月期の全員の残業合計。prevPersonsOt は前年度の同じ経過月期までの担当別残業合計。',
    '- elapsed は経過した月期数、curIdx は進行中の月期の位置（-1 なら無し）。進行中の月期は途中の値。',
    '- 36協定の管理基準: 月 otMonthWarn h で警告、月 otMonthLimit h が上限、年 otYearLimit h が上限、月上限超えは年 otOverMonthsMax 回まで。',
    '- 年間見込みは summary.persons[担当].projYearOt を使う（締まった月期の累計＋平均×残り月期）。remainingMonths が0の年度は確定済みなので「見込み」を書かない。途中から記録のある担当を12か月に換算しない。',
    '',
    '【出力形式】',
    '- 次の形のJSONだけを出力する。前後に説明文やコードブロック記号を付けない。',
    '{"headline":"…","status":"good|warn|bad","sections":{"overview":[…],"overtime":[…],"workmix":[…],"concerns":[…]},"notes":{"trend":"…","heat":"…","fy":"…","mix":"…"},"actions":[{"text":"…","who":"…"}]}',
    '- headline: 年度全体を一文で総括（60字以内）。status: bad=年上限超過や超過見込みなど対応が必要、warn=注意、good=概ね良好。',
    '- sections の各配列の要素は {"tone":"good|warn|bad|info","text":"…"}。text は1項目80字以内、各セクション2〜4項目。',
    '  overview=年度の全体像、overtime=残業の推移と36協定、workmix=業務構成の変化、concerns=気になる点。',
    '- notes: 画面の各グラフに添える一言（各60字以内）。trend=月期ごとの全員残業の推移と前年比、heat=担当×月期の残業（警告・上限の発生状況）、fy=年度累計と年上限に対する見込み、mix=大分類構成の月期推移。',
    '- actions: 残りの期間に向けた具体的な打ち手を3〜4件。who は主に関わる担当名か「課全体」。',
    '- 経過月期が少ない場合は、その旨を明記し、断定を避ける。'
  ].concat(COMMON_RULES_).join('\n');
}

function systemPromptKpi_() {
  return [
    'あなたは製造業（株式会社ホリゾン、製本・印刷関連機器メーカー）の生産管理課の業務分析担当です。',
    '課のKPI（月期ごとの工数目標。全員の工数合計を目標以下に収める）を年度で集計したJSONを受け取り、課長・各KPIの主担当が進捗確認に使うレポートを日本語で書いてください。',
    '',
    '【データの前提】',
    '- periods は月期（KPI の月期は1日〜月末）。配列の値は periods と同じ順。null は未経過またはデータなし。',
    '- kpis[] の各要素: key, name, target（月期ごとの上限h）, owner（主担当）, plan（アクションプラン）, vals（月期ごとの実績h）, byAuthor（年度累計の担当別h）, bySub（年度累計の中分類別h）。kind=ot は全員の残業時間合計の KPI。',
    '- curIdx は進行中の月期の位置（-1 なら無し）。進行中の月期の値は途中経過で、proj を掛けると締め日までの見込みになる。',
    '- 判定: 実績（進行中は見込み）が target 以下=達成、1.2倍まで=超過、1.2倍超=大幅超過。',
    '',
    '【出力形式】',
    '- 次の形のJSONだけを出力する。前後に説明文やコードブロック記号を付けない。',
    '{"headline":"…","status":"good|warn|bad","sections":{"overview":[…],"concerns":[…]},"kpis":{"<key>":{"tone":"good|warn|bad|info","text":"…"}},"actions":[{"text":"…","who":"…"}]}',
    '- headline: KPI全体の状況を一文で総括（60字以内）。status: bad=大幅超過が続くKPIあり、warn=超過あり、good=概ね達成。',
    '- sections の各配列の要素は {"tone":"good|warn|bad|info","text":"…"}。text は1項目80字以内、各2〜4項目。overview=全体の達成状況、concerns=気になる点。',
    '- kpis: すべての KPI の key について1件ずつ、100字以内。達成した月期数と傾向、今期の見込み、要因（担当や中分類の偏り）、アクションプランとの関係を書く。',
    '- actions: KPI 達成に向けた具体的な打ち手を3〜4件。who は主担当名か「課全体」。'
  ].concat(COMMON_RULES_).join('\n');
}

function callClaude_(data) {
  var key = apiKey_();
  if (!key) throw new Error('APIキーが未設定です（GASのスクリプトプロパティ ANTHROPIC_API_KEY）');
  var base = {
    model: model_(),
    system: systemPrompt_(data && data.kind),
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

// ═══════════════════════════════════════════════════════════════
//  自動分析（時間主導トリガーで毎朝実行）
//  - 平日（土日・祝日を除く）: 月次（当月期）と KPI（当年度）
//  - 月期の締め（15日）後の最初の平日: 締まった月期の月次と、その年度の年次も
//  日報データは Firestore（nippou-f191c）をこのスクリプトの実行ユーザーの権限で直接読む。
//  集計は日報アプリの公開コード（index.html / config.js）から同じ関数を取り出して実行するため、
//  画面に出る図表と同じ数値・同じハッシュになる（画面で「日報が更新されています」と誤表示されない）。
//
//  初回だけ: エディタで setupAutoTrigger を実行（毎朝6時台に autoDaily を実行するトリガーを作成）。
//  動作確認: autoDailyDryRun（Claude は呼ばず、集計と判定だけログに出す）/ autoDailyForce（今すぐ分析）。
// ═══════════════════════════════════════════════════════════════
var APP_SRC_BASE = 'https://raw.githubusercontent.com/HORIZON-JIT/NP/main/';
var FIRESTORE_PROJECT = 'nippou-f191c';
var FIRESTORE_COLLECTION = 'dailyReports';
var AUTO_HOUR = 6;

// 日報アプリの index.html から取り出す宣言（関数・定数）
var APP_CORE_NAMES = [
  'pad', '_hmToMin', 'DEPARTMENT_CONFIG', 'JP_HOLIDAYS', 'todayStr', 'periodOfDate', 'kpiPeriodDates', 'getPeriodDates',
  'isHoliday', 'getLunchBreak', '_calcFromLeaveCache', 'calcFromLeave', 'normalizeRows', 'eachDateInRange', '_authorsLast',
  '_otByAuthor', '_scaledRows', '_effStats', '_kpiHoursOf', '_kpiOtOf', '_kpiBreakdown', '_kpiWorkdays', '_kpiJudge',
  '_normKpis', '_kpiHoursItems', '_kpiOtItem', 'getCurrentFiscalYear', 'getFiscalPeriods', 'fiscalKi', '_fiscalYearOf',
  '_aiR', '_aiHash', '_aiSubs', 'buildAiPayload', '_aiYearPayload', '_aiKpiPayload'
];

// ─── ソースから宣言を1つ取り出す（文字列・テンプレート・コメント・正規表現を読み飛ばして括弧を数える） ───
function extractDecl_(src, name) {
  var esc = name.replace(/[$]/g, '\\$');
  var re = new RegExp('(^|\\n)((?:async\\s+)?function\\s+' + esc + '\\s*\\(|(?:const|let|var)\\s+' + esc + '\\s*=)');
  var m = re.exec(src);
  if (!m) throw new Error('日報アプリのコードに ' + name + ' が見つかりません');
  var start = m.index + m[1].length, isFn = /function/.test(m[2]);
  var i = start, depth = 0, opened = false, tpl = [], prev = '';
  while (i < src.length) {
    var c = src[i], n2 = src[i + 1];
    if (c === '/' && n2 === '/') { i = src.indexOf('\n', i); if (i < 0) break; continue; }
    if (c === '/' && n2 === '*') { i = src.indexOf('*/', i + 2) + 2; continue; }
    if (c === '"' || c === "'") { var q = c; i++; while (i < src.length && src[i] !== q) { if (src[i] === '\\') i++; i++; } i++; prev = q; continue; }
    if (c === '`' || (c === '}' && tpl.length && tpl[tpl.length - 1] === depth)) {
      if (c === '}') { tpl.pop(); depth--; }
      i++;
      while (i < src.length && src[i] !== '`') {
        if (src[i] === '\\') { i += 2; continue; }
        if (src[i] === '$' && src[i + 1] === '{') { depth++; tpl.push(depth); i += 2; break; }
        i++;
      }
      if (src[i] === '`') i++;
      prev = '`'; continue;
    }
    if (c === '/' && /[(,=:\[!&|?{};+\-*%<>~^]|^$/.test(prev)) {   // 正規表現リテラル
      i++; var cls = false;
      while (i < src.length) { var r = src[i]; if (r === '\\') { i += 2; continue; } if (r === '[') cls = true; else if (r === ']') cls = false; else if (r === '/' && !cls) break; i++; }
      i++; while (/[a-z]/i.test(src[i])) i++;
      prev = '/'; continue;
    }
    if (c === '(' || c === '[' || c === '{') { depth++; opened = true; }
    else if (c === ')' || c === ']' || c === '}') {
      depth--;
      if (isFn && c === '}' && depth === 0 && opened) return src.slice(start, i + 1);
    } else if (!isFn && c === ';' && depth === 0) return src.slice(start, i + 1);
    if (!/\s/.test(c)) prev = c;
    i++;
  }
  throw new Error(name + ' の終わりを見つけられませんでした');
}

// ─── 日報アプリの集計関数を読み込む ───
function loadAppCore_() {
  var html = UrlFetchApp.fetch(APP_SRC_BASE + 'index.html?t=' + Date.now()).getContentText();
  var cfg = UrlFetchApp.fetch(APP_SRC_BASE + 'config.js?t=' + Date.now()).getContentText();
  var decls = APP_CORE_NAMES.map(function (n) { return extractDecl_(html, n); });
  var src = [
    'var window = {}, document = undefined, localStorage = undefined;',
    cfg,
    'var SUB_MAP = {}, CATS = [], CAT_LABEL = {}, CAT_COLOR = {}, KPIS = [], _kpiCtx = null, _aiState = null, _entryCtx = null, currentDate = "";',
    decls.join('\n\n'),
    'return {',
    '  set: function (k, v) { switch (k) { case "CATS": CATS = v; break; case "KPIS": KPIS = v; break; case "_kpiCtx": _kpiCtx = v; break; case "_aiState": _aiState = v; break; case "_entryCtx": _entryCtx = v; break; case "currentDate": currentDate = v; break; } },',
    '  get KPIS() { return KPIS; }, get CATS() { return CATS; },',
    '  SUB_MAP: SUB_MAP, CAT_LABEL: CAT_LABEL, CAT_COLOR: CAT_COLOR, DEPARTMENT_CONFIG: DEPARTMENT_CONFIG,',
    '  ' + APP_CORE_NAMES.filter(function (n) { return n !== 'DEPARTMENT_CONFIG'; }).map(function (n) { return n + ': ' + n; }).join(', '),
    '};'
  ].join('\n');
  return new Function(src)();
}

// ─── Firestore（REST）───
function fsValue_(v) {
  if (!v) return null;
  if ('stringValue' in v) return v.stringValue;
  if ('integerValue' in v) return Number(v.integerValue);
  if ('doubleValue' in v) return Number(v.doubleValue);
  if ('booleanValue' in v) return v.booleanValue;
  if ('nullValue' in v) return null;
  if ('timestampValue' in v) return v.timestampValue;
  if ('arrayValue' in v) return (v.arrayValue.values || []).map(fsValue_);
  if ('mapValue' in v) { var o = {}, f = v.mapValue.fields || {}; Object.keys(f).forEach(function (k) { o[k] = fsValue_(f[k]); }); return o; }
  return null;
}
function fsDoc_(doc) { var o = {}, f = (doc && doc.fields) || {}; Object.keys(f).forEach(function (k) { o[k] = fsValue_(f[k]); }); return o; }
function fsBase_() { return 'https://firestore.googleapis.com/v1/projects/' + FIRESTORE_PROJECT + '/databases/(default)/documents'; }
function fsFetch_(url, payload) {
  var opt = {headers: {Authorization: 'Bearer ' + ScriptApp.getOAuthToken()}, muteHttpExceptions: true};
  if (payload) { opt.method = 'post'; opt.contentType = 'application/json'; opt.payload = JSON.stringify(payload); }
  var res = UrlFetchApp.fetch(url, opt), code = res.getResponseCode();
  if (code !== 200) throw new Error('Firestore 読み取りエラー（' + code + '）: ' + res.getContentText().slice(0, 300));
  return JSON.parse(res.getContentText());
}
function fsReportsBetween_(startDate, endDate) {
  var out = fsFetch_(fsBase_() + ':runQuery', {structuredQuery: {
    from: [{collectionId: FIRESTORE_COLLECTION}],
    where: {compositeFilter: {op: 'AND', filters: [
      {fieldFilter: {field: {fieldPath: 'date'}, op: 'GREATER_THAN_OR_EQUAL', value: {stringValue: startDate}}},
      {fieldFilter: {field: {fieldPath: 'date'}, op: 'LESS_THAN_OR_EQUAL', value: {stringValue: endDate}}}
    ]}}
  }});
  return out.filter(function (r) { return r.document; }).map(function (r) { return fsDoc_(r.document); });
}
function fsMaster_(deptId) {
  try { return fsDoc_(fsFetch_(fsBase_() + '/' + FIRESTORE_COLLECTION + '/' + deptId + '_master_categories')); }
  catch (e) { Logger.log('マスタを読めませんでした（既定値を使用）: ' + e.message); return {}; }
}

// 日報アプリの loadRangeDataFirebase と同じ形に集計: {担当: [{cdSub, hoursNormal, hoursOT}], _leaveMap: {担当: {日付: 退勤}}}
function rangeData_(core, docs, startDate, endDate) {
  var tmp = {}, leaveMap = {};
  docs.forEach(function (d) {
    if (!d.author || !d.date || d.date < startDate || d.date > endDate) return;
    core.normalizeRows(d.rows || []).forEach(function (r) {
      var a = tmp[d.author] || (tmp[d.author] = {});
      var x = a[r.cdSub] || (a[r.cdSub] = {cdSub: r.cdSub, hoursNormal: 0, hoursOT: 0});
      x.hoursNormal += r.hoursNormal || 0; x.hoursOT += r.hoursOT || 0;
    });
    if (d.leaveTime) { (leaveMap[d.author] = leaveMap[d.author] || {})[d.date] = d.leaveTime; }
  });
  var result = {_leaveMap: leaveMap};
  Object.keys(tmp).forEach(function (a) { result[a] = Object.keys(tmp[a]).map(function (k) { return tmp[a][k]; }); });
  return result;
}

// 項目・KPI マスタを日報アプリと同じように反映
function applyMaster_(core, master) {
  var cats = Array.isArray(master.cats) && master.cats.length ? master.cats : core.DEPARTMENT_CONFIG.defaultCats;
  core.set('CATS', cats);
  Object.keys(core.SUB_MAP).forEach(function (k) { delete core.SUB_MAP[k]; });
  cats.forEach(function (c) {
    core.CAT_COLOR[c.cd] = c.color || '#6b7280'; core.CAT_LABEL[c.cd] = c.label || c.cd;
    (c.subs || []).forEach(function (s) {
      var o = {}; Object.keys(s).forEach(function (k) { o[k] = s[k]; });
      o.mainCd = c.cd; o.mainLabel = c.label; o.color = c.color; o.mainActive = c.active !== false; o.active = s.active !== false;
      core.SUB_MAP[s.cd] = o;
    });
  });
  core.set('KPIS', core._normKpis(Array.isArray(master.kpis) ? master.kpis : null));
}

// 日報アプリの「入力状況」と同じ判定で、未入力日数と稼働日を出す
function entryCtx_(core, key, leaveMap, startDate, endDate) {
  var list = (core.DEPARTMENT_CONFIG.defaultAuthors || []).filter(function (a) { return a && a !== 'ヘルプ'; });
  Object.keys(leaveMap || {}).forEach(function (a) { if (a !== 'ヘルプ' && list.indexOf(a) < 0) list.push(a); });
  var authors = core._authorsLast(list), today = core.todayStr();
  var anyEntry = function (d) { return authors.some(function (a) { return leaveMap[a] && leaveMap[a][d]; }); };
  var dayInfo = core.eachDateInRange(startDate, endDate).map(function (d) {
    var dow = new Date(d + 'T00:00:00').getDay(), weekend = dow === 0 || dow === 6, future = d > today;
    var holiday = !weekend && !future && d < today && !anyEntry(d);
    return {d: d, weekend: weekend, future: future, off: weekend || holiday};
  });
  var missing = {};
  authors.forEach(function (a) {
    var miss = 0;
    dayInfo.forEach(function (x) { var t = leaveMap[a] && leaveMap[a][x.d]; if (!t && !x.future && !x.off) miss++; });
    missing[a] = miss;
  });
  var elapsed = dayInfo.filter(function (x) { return !x.off && !x.future && (x.d < today || anyEntry(x.d)); }).length;
  var total = elapsed + dayInfo.filter(function (x) { return !x.weekend && (x.future || (x.d === today && !anyEntry(x.d))); }).length;
  return {key: key, missing: missing, workElapsed: elapsed, workTotal: total, inProgress: startDate <= today && today <= endDate};
}

// ─── 各分析の送るデータ（画面の buildAiPayload / _aiKpiPayload / _aiYearPayload と同じ関数で作る） ───
function monthlyPayload_(core, docs, period) {
  var pd = core.getPeriodDates(period.year, period.month), key = pd.startDate + '_' + pd.endDate;
  var cur = rangeData_(core, docs, pd.startDate, pd.endDate);
  if (!Object.keys(cur._leaveMap).length) return null;
  var prevP = period.month === 1 ? {year: period.year - 1, month: 12} : {year: period.year, month: period.month - 1};
  var pp = core.getPeriodDates(prevP.year, prevP.month), prev = rangeData_(core, docs, pp.startDate, pp.endDate);
  var fyPeriods = core.getFiscalPeriods(core._fiscalYearOf(period));
  var fyIdx = fyPeriods.findIndex(function (x) { return x.year === period.year && x.month === period.month; });
  var before = fyPeriods.slice(0, Math.max(0, fyIdx)).map(function (x) { var r = core.getPeriodDates(x.year, x.month); return rangeData_(core, docs, r.startDate, r.endDate); });
  var ctx = {startDate: pd.startDate, endDate: pd.endDate, period: period, cur: core._otByAuthor(cur), curRaw: cur,
    prev: core._otByAuthor(prev), prevRaw: prev, fy: before.map(core._otByAuthor).concat([core._otByAuthor(cur)])};
  core.set('_aiState', {key: key, label: pd.label, ctx: ctx});
  core.set('_entryCtx', entryCtx_(core, key, cur._leaveMap, pd.startDate, pd.endDate));
  return {key: key, payload: core.buildAiPayload()};
}
function kpiPayload_(core, docs, fy) {
  if (!core.KPIS.length) return null;
  var periods = core.getFiscalPeriods(fy), ranges = periods.map(function (x) { return core.kpiPeriodDates(x.year, x.month); });
  var today = core.todayStr();
  var curIdx = ranges.findIndex(function (r) { return r.startDate <= today && today <= r.endDate; });
  var lastIdx = -1; ranges.forEach(function (r, i) { if (r.startDate <= today) lastIdx = i; });
  var data = periods.map(function (x, i) { return i <= lastIdx ? rangeData_(core, docs, ranges[i].startDate, ranges[i].endDate) : null; });
  var wd = curIdx >= 0 && data[curIdx] ? core._kpiWorkdays(data[curIdx], ranges[curIdx].startDate, ranges[curIdx].endDate) : null;
  var proj = wd && wd.elapsed > 0 && wd.total > wd.elapsed ? wd.total / wd.elapsed : 1;
  core.set('_kpiCtx', {fy: fy, periods: periods, ranges: ranges, curIdx: curIdx, lastIdx: lastIdx, data: data, proj: proj});
  var p = core._aiKpiPayload();
  return p ? {key: 'kpi_' + fy, payload: p} : null;
}
function yearlyPayload_(core, docs, fy) {
  var load = function (f) { return core.getFiscalPeriods(f).map(function (x) { var r = core.getPeriodDates(x.year, x.month); return rangeData_(core, docs, r.startDate, r.endDate); }); };
  var p = core._aiYearPayload({fy: fy, periods: core.getFiscalPeriods(fy), data: load(fy), prev: load(fy - 1)});
  return p ? {key: 'fy_' + fy, payload: p} : null;
}

// ─── 1件分析して保存（集計が前回と同じなら Claude を呼ばない） ───
function autoAnalyzeOne_(core, job, dry) {
  if (!job || !job.payload) return 'データなし';
  var hash = core._aiHash(JSON.stringify(job.payload)), sk = storeKey_(job.key), prev = readRec_(sk);
  if (prev && prev.hash === hash) return '変更なし（前回の分析を継続）';
  if (dry) return '分析対象（dry run のため未実行）hash=' + hash;
  var text = callClaude_(job.payload);
  writeRec_(sk, {text: text, at: new Date().toISOString(), model: model_(), hash: hash, by: 'auto'});
  return '分析して保存しました';
}

function autoDailyRun_(force, dry) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) { Logger.log('他の処理が実行中のためスキップ'); return; }
  try {
    var core = loadAppCore_();
    var today = core.todayStr(); core.set('currentDate', today);
    if (!force && core.isHoliday(today)) { Logger.log(today + ' は土日・祝日のためスキップ'); return; }
    var y = Number(today.slice(0, 4)), mo = Number(today.slice(5, 7)), dd = Number(today.slice(8, 10));
    // 締め（15日）後の最初の平日か: 16日〜昨日までに平日が無い
    var closeDay = false;
    if (dd >= 16) {
      closeDay = true;
      for (var d = 16; d < dd; d++) { if (!core.isHoliday(y + '-' + core.pad(mo) + '-' + core.pad(d))) { closeDay = false; break; } }
    }
    applyMaster_(core, fsMaster_(core.DEPARTMENT_CONFIG.deptId));
    var curP = core.periodOfDate(today), closedP = {year: y, month: mo};
    var fyNow = core.getCurrentFiscalYear(), fyClosed = core._fiscalYearOf(closedP);
    // 読む範囲: 前年度の期初（年次の前年比用）〜今日
    var fromFy = Math.min(fyNow, fyClosed) - (closeDay ? 1 : 0);
    var from = core.getPeriodDates(fromFy, 10).startDate;
    var kpiFrom = core.kpiPeriodDates(fyNow, 10).startDate;
    if (kpiFrom < from) from = kpiFrom;
    var docs = fsReportsBetween_(from, today);
    Logger.log('日報 ' + docs.length + '件を読み込み（' + from + '〜' + today + '）');
    var jobs = [['月次 ' + curP.year + '年' + curP.month + '月期', function () { return monthlyPayload_(core, docs, curP); }],
                ['KPI ' + (fyNow + 1) + '年度', function () { return kpiPayload_(core, docs, fyNow); }]];
    if (closeDay) {
      jobs.push(['月次（締め） ' + closedP.year + '年' + closedP.month + '月期', function () { return monthlyPayload_(core, docs, closedP); }]);
      jobs.push(['年次 ' + (fyClosed + 1) + '年度', function () { return yearlyPayload_(core, docs, fyClosed); }]);
    }
    jobs.forEach(function (j) {
      try { Logger.log(j[0] + ': ' + autoAnalyzeOne_(core, j[1](), dry)); }
      catch (e) { Logger.log(j[0] + ': 失敗 ' + (e && e.message || e)); }
    });
  } finally { lock.releaseLock(); }
}
function autoDaily() { autoDailyRun_(false, false); }       // トリガーから呼ばれる
function autoDailyForce() { autoDailyRun_(true, false); }   // 手動: 今すぐ分析（集計が変わったものだけ）
function autoDailyDryRun() { autoDailyRun_(true, true); }   // 手動: 確認だけ（Claude は呼ばない）

function setupAutoTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'autoDaily') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('autoDaily').timeBased().atHour(AUTO_HOUR).everyDays(1).inTimezone('Asia/Tokyo').create();
  Logger.log('毎朝' + AUTO_HOUR + '時台に autoDaily を実行するトリガーを作成しました');
}
