# GAS の自動デプロイ

`gas/task-sheet-reader.gs` を main にマージすると、GitHub Actions（`.github/workflows/deploy-gas.yml`）が
Folio 連携の GAS を「新バージョン」としてデプロイします（URL は変わりません）。

## 一度だけの準備（GitHub の Secrets）

| Secret | 内容 |
|---|---|
| `CLASPRC_JSON` | PC で `clasp login` したときにできる `~/.clasprc.json` の中身（JSON まるごと） |
| `GAS_SCRIPT_ID` | GAS エディタ「プロジェクトの設定」にあるスクリプト ID |

- 事前に https://script.google.com/home/usersettings で「Google Apps Script API」を ON にする。
- clasp は 2.4.2 を使う（`npm i -g @google/clasp@2.4.2`）。認証ファイルの形式が違うため 3 系は不可。
- 手動で動かすときは GitHub の Actions タブ →「Deploy GAS」→ Run workflow。

# 月次AI分析の GAS（生産管理課のみ）

`gas/ai-monthly-analyzer.gs` は月次ダッシュボードの「AI分析」カード用です。自動デプロイの対象外なので、初回は手動で設置します。

1. https://script.google.com/ で新しいプロジェクトを作成し、`gas/ai-monthly-analyzer.gs` の中身を貼り付けて保存
2. 「プロジェクトの設定」→「スクリプト プロパティ」に `ANTHROPIC_API_KEY`（Claude APIキー）を追加
   （モデルを変える場合だけ `CLAUDE_MODEL` も追加。省略時は `claude-sonnet-5-5`）
3. エディタで `checkSetup` を実行し、外部通信を許可。ログに「OK: 準備完了です」と出れば成功
4. 「デプロイ」→「新しいデプロイ」→ ウェブアプリ（実行ユーザー: 自分／アクセス: horizon.co.jp 内の全員）
5. 発行された URL をルートの `config.js` の `aiAnalysisGasUrl` に入れて main にマージ

- APIキーは GAS のスクリプトプロパティにだけ置き、HTML やリポジトリには書かない
- 法務知財部（`legal/config.js`）は `aiAnalysisGasUrl` を設定していないため、カードは表示されない
- 分析結果は GAS のスクリプトプロパティに月期ごとに保存され、画面を開いたときは保存済みを表示するだけ（APIは呼ばない）。API を呼ぶのは「分析する」「再分析」を押したときのみ

## 月次AI分析の自動実行（毎朝6時台）

- 平日（土日・祝日除く）: 月次（当月期）・KPI（当年度）を分析
- 月期の締め（15日）後の最初の平日: 締まった月期の月次と、その年度の年次も分析
- 集計が前回と同じなら Claude を呼ばない（料金なし）
- 日報データは Firestore（nippou-f191c）を実行ユーザー（プロジェクトのオーナー）の権限で読む
- 集計は日報アプリの公開コード（main の index.html / config.js）から同じ関数を取り出して使う。
  そのため画面の図表と同じ数値になり、アプリ側の集計を直すと自動分析にも反映される

設定（初回のみ）:
1. プロジェクトの設定で「appsscript.json」をエディタに表示し、次を設定
   - "timeZone": "Asia/Tokyo"
   - "oauthScopes": script.external_request / datastore / script.scriptapp
2. エディタで autoDailyDryRun を実行（権限の承認。Claude は呼ばずログだけ）
3. setupAutoTrigger を実行（毎朝6時台の時間主導トリガーを作成）
