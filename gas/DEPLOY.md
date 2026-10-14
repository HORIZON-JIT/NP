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
