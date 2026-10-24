# Folio 連携（Firebase）

日報アプリは Folio のデータを `folioCall()` 経由で、Folio の保存先 Firebase（プロジェクト folio-73ebf）からブラウザで直接読む。
2026/10/23 の Folio 移行で、スプレッドシート＋GAS（v1・`gas/task-sheet-reader.gs`）からの読み込みは廃止した。

Folio 連携は生産管理課だけ。`config.js` に `folio`（Firebase の接続情報）がある部署だけに、「Folioから取込」・タスク早見表・ダッシュボードの Folio 実績が出る。

## 読むデータ

| 日報アプリの機能 | 読み書き先 |
| --- | --- |
| Folioから取込 | `tasks`（`day == 日付`、時間割に未配置のものは `doneDate == 日付`。`status == '完了'`） |
| タスク早見表 | `masters/taskMaster` の `list`（id・name・std） |
| 日報中分類の対応（読み書き） | `nippouMap/taskCd`（`{タスクID: 中分類CD}`） |
| 予実対比・個人の Folio 完了タスク・タスク別 標準工数と実績 | `tasks`（`day` の範囲＋未配置の完了タスク） |

項目の対応: `email`・`ownerName`・`taskId`・`taskName`・`minutes`（予定、分）・`actual`（実績、分。0 は未入力扱い）・`status`・`note`・`start`（"YYYY-MM-DD HH:MM"）・`day`・`doneDate`。
作業日は `day`（時間割の配置日）。`day` が空なら `start` の日付、それも空なら `doneDate`。工数は実績が入っていれば実績、なければ予定。

## 読み取り回数の節約（Firestore 無料枠 1日5万回）

- 完了タスクは日ごとに端末（localStorage `fv2:folio-73ebf:d:YYYY-MM-DD`）へ保存して使い回す。保存期間は日付の古さで変える: 昨日以降 2分／2〜7日前 6時間／それより前 7日。130日より古い日は自動で消す
- 期間の読み込みは配置日 `day` の範囲クエリ1回だけ。古くなった日だけを連続区間にまとめて読む
- 時間割に未配置のタスク（`day` が空）
  - 「Folioから取込」: その1日分だけ `doneDate == 日付` で読む
  - ダッシュボードの集計: `day == '' かつ status == '完了'` をまとめて1回読み、端末に6時間保存（localStorage `fv2:folio-73ebf:unp`）
- 「Folioから取込」を開いたとき・「最新に更新」はその日の分を読み直す
- 設定画面「Folio 連携」に、その日にこの端末で読んだ件数を表示する

## ログイン

各ユーザーは最初に Folio を使うときに一度だけ Google ログインが出る（「Folioから取込」を押したとき、またはダッシュボードの Folio カードの「Folio にログイン」）。以後はブラウザがログイン状態を覚えている。

## Folio 側の設定（設定済み）

1. Authentication > ログイン方法で **Google** を有効
2. Authentication > 設定 > 承認済みドメインに **horizon-jit.github.io**
3. セキュリティルールで `tasks`・`masters` の読み取りと `nippouMap` の読み書きを社内アドレスに許可

## 旧 v1（廃止）

`gas/task-sheet-reader.gs` と `.github/workflows/deploy-gas.yml` は v1 の名残で、日報アプリからは呼ばない。
