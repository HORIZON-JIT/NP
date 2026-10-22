# Folio v2（Firebase 保存版）への切替手順

日報アプリは Folio のデータを `folioCall()` 経由で読む。読込元は部署設定 `folio.source` で切り替える。

| 読込元 | 設定 | 読む場所 |
| --- | --- | --- |
| v1（現行） | `folio.source: 'sheet'` | スプレッドシート「タスク管理アプリ v2」を `taskSheetGasUrl` の GAS（gas/task-sheet-reader.gs）経由 |
| v2 | `folio.source: 'firebase'` | Folio v2 の Firestore（プロジェクト folio-73ebf）をブラウザから直接 |

各端末は、設定画面「Folio の読込元」で部署設定と関係なく v1／v2 を切り替えられる（試験用。localStorage に保存）。

## v2 で読むデータ

| 日報アプリの機能 | v2 の読み書き先 |
| --- | --- |
| Folioから取込 | `tasks`（`day` または `doneDate` が期間内、`status == '完了'`） |
| タスク早見表 | `masters/taskMaster` の `list`（id・name・std） |
| 日報中分類の対応（読み書き） | `nippouMap/taskCd`（`{タスクID: 中分類CD}`） |
| 月次の Folio 実績・個人の Folio 完了タスク | `tasks`（同上） |

項目の対応: `email`・`ownerName`・`taskId`・`taskName`・`minutes`（予定、分）・`actual`（実績、分。0 は未入力扱い）・`status`・`note`・`start`（"YYYY-MM-DD HH:MM"）・`day`・`doneDate`。
作業日は `day`（時間割の配置日）、空なら `doneDate`。工数は実績が入っていれば実績、なければ予定。

`nippouMap/taskCd` が空のとき、v1 の GAS が使えれば v1 シートの「日報中分類CD」列を1回だけ引き継ぐ。

## Folio 側で必要な設定（切替前に1回）

1. Firebase コンソール > Authentication > ログイン方法で **Google** を有効にする
2. Authentication > 設定 > 承認済みドメインに **horizon-jit.github.io** を追加
3. Folio の `コード.gs` の `FB_RULES` に次の1行を追加し、`fbRulesDeploy` を実行
   ```
   match /nippouMap/{id} { allow read, write: if inside(); }
   ```
   （`tasks` と `masters` の読み取りは既存ルールの `inside()` で許可済み）

## 切替の流れ

1. 上の Folio 側設定を済ませる
2. 谷さんの端末で設定画面「Folio の読込元」を v2 にし、「v2 の接続テスト」で件数を確認。取込・早見表・ダッシュボードを v1 と見比べる
3. Folio 側で `fbCutover_` を実行（v1 シートから Firestore への最終同期）
4. `config.js` の `folio.source` を `'firebase'` に変えて公開 → 全員が v2 から読む
5. 問題があれば `'sheet'` に戻す（切替後は v1 シートが更新されないので一時しのぎ）

各ユーザーは v2 に切り替わった後、最初に「Folioから取込」を押したときに一度だけ Google ログインが出る。
