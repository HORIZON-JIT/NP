// ═══════════════════════════════════════════════════════════════
//  法務知財部 用 設定（/NP/homu/ に配置）
// ═══════════════════════════════════════════════════════════════
//  index.html は生産管理課と共通（単一ソース）。この config.js だけが
//  法務知財部固有の値を上書きする。未記載の項目は生産管理課の既定値を継承。
//
//  ★★★ 立ち上げに必要な唯一の残作業 ★★★
//   下記 firebase を「法務知財部専用の新規Firebaseプロジェクト」の設定に差し替える。
//   （現在は 'YOUR_...' プレースホルダのため fail-closed = Firebaseに接続せず
//     ローカルモードで動作。生産管理課のデータには一切触れない安全状態）
//
//  【Firebaseプロジェクト作成手順（概要）】
//   1) https://console.firebase.google.com/ で新規プロジェクト作成
//   2) Firestore Database を有効化（本番モード）
//   3) Authentication → Sign-in method → Google を有効化
//   4) Authentication → Settings → 承認済みドメインに horizon-jit.github.io を追加
//   5) プロジェクト設定 → マイアプリ → ウェブアプリを追加 → 表示される
//      firebaseConfig（apiKey等）を下の firebase ブロックに貼り付け
//   6) Firestore のセキュリティルールは生産管理課の既存ルールを流用
//      （@horizon.co.jp / @bp.horizon.co.jp を許可）
//
//  ※ カテゴリ(defaultCats)は「新プロジェクトが空のときの初期の種」。
//    起動後は「管理」タブでいつでも追加/編集/並べ替え/有効無効が可能。
// ═══════════════════════════════════════════════════════════════

window.DEPARTMENT_CONFIG_OVERRIDE = {
  deptId: 'homu',                  // localStorageキー / マスタdoc id のプレフィックス
  collection: 'dailyReports',      // 別Firebaseプロジェクトなので 'dailyReports' のままで衝突なし

  branding: {
    appTitle: '日報アプリ',
    docTitle: '日報アプリ｜ホリゾン法務知財',
    company:  'HORIZON',
    dept:     '法務知財部'
  },

  // GASバックアップは未使用。空にして生産管理課のGASエンドポイント継承を防ぐ（重要）。
  gasUrl: '',

  // ★ 法務知財部用の新Firebaseプロジェクトの設定に差し替える（必須・唯一の残作業）
  //   'YOUR_' が残っている間は Firebase に接続しない（fail-closed / ローカルモード）。
  firebase: {
    apiKey:            'YOUR_API_KEY',
    authDomain:        'YOUR_PROJECT.firebaseapp.com',
    projectId:         'YOUR_PROJECT',
    storageBucket:     'YOUR_PROJECT.firebasestorage.app',
    messagingSenderId: 'YOUR_SENDER_ID',
    appId:             'YOUR_APP_ID'
  },

  allowedDomains: ['horizon.co.jp', 'bp.horizon.co.jp'],

  // 担当者選択の初期リスト。起動後は設定画面から各自追加も可能。
  defaultAuthors: ['中西'],

  // 外部リンクは法務知財部側で未設定。生産管理課のバックアップシート継承を防ぐため空に。
  links: { backupSheetUrl: '' },

  // 勤務ルールは生産管理課と同じ → rules は省略（既定値を継承）。

  // 初期カテゴリ（法務知財部。アップロード資料から作成。管理タブで編集可）
  defaultCats: [
    { cd:'A', label:'管理', color:'#3b82f6', subs:[
      { cd:'A1', label:'社外対応' },
      { cd:'A2', label:'社内' },
      { cd:'A3', label:'その他' }
    ] },
    { cd:'B', label:'調査', color:'#22c55e', subs:[
      { cd:'B1', label:'【特許】SDI' },
      { cd:'B2', label:'【特許】技術動向' },
      { cd:'B3', label:'【特許】侵害予防' },
      { cd:'B4', label:'【特許】先行文献' },
      { cd:'B5', label:'【特許】無効資料' },
      { cd:'B6', label:'【特許】その他' },
      { cd:'B7', label:'【商標】侵害予防' },
      { cd:'B8', label:'【商標】その他' },
      { cd:'B9', label:'【その他】' }
    ] },
    { cd:'C', label:'打合せ', color:'#ef4444', subs:[
      { cd:'C1', label:'打合せ資料準備' },
      { cd:'C2', label:'打合せ' },
      { cd:'C3', label:'問い合わせ対応' }
    ] },
    { cd:'D', label:'教育・研修', color:'#a855f7', subs:[
      { cd:'D1', label:'知財教育・研修' },
      { cd:'D2', label:'その他' }
    ] },
    { cd:'E', label:'改善', color:'#eab308', subs:[
      { cd:'E1', label:'ツール' },
      { cd:'E2', label:'業務' }
    ] },
    { cd:'F', label:'法務', color:'#06b6d4', subs:[
      { cd:'F1', label:'契約' },
      { cd:'F2', label:'その他' }
    ] },
    { cd:'G', label:'グループ会社', color:'#6b7280', subs:[
      { cd:'G1', label:'富士油圧精機' },
      { cd:'G2', label:'東阪電子機器' },
      { cd:'G3', label:'その他' }
    ] },
    { cd:'H', label:'その他', color:'#ec4899', subs:[
      { cd:'H1', label:'その他' }
    ] }
  ]
};
