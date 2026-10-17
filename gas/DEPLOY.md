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
