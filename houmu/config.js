// ═══════════════════════════════════════════════════════════════
//  法務知財部（法務・深谷さん）用 設定（/NP/houmu/ に配置）
// ═══════════════════════════════════════════════════════════════
//  index.html は生産管理課と共通（単一ソース）。この config.js だけが
//  この派生アプリ固有の値を上書きする。未記載の項目は生産管理課の既定値を継承。
//
//  Firebaseプロジェクト: nippou-legal（法務知財部と共用）。
//  コレクションを 'dailyReports_houmu' に分けることで、知的財産室（/legal/）の
//  日報・マスタとは完全に別のデータとして保存する。
//
//  ※ カテゴリ(defaultCats)は「コレクションが空のときの初期の種」。
//    起動後は「管理」タブでいつでも追加/編集/並べ替え/有効無効が可能。
//    （2026/10 深谷さん記入の分類登録シートから作成）
// ═══════════════════════════════════════════════════════════════

window.DEPARTMENT_CONFIG_OVERRIDE = {
  deptId: 'houmu',                 // localStorageキー / マスタdoc id のプレフィックス
  collection: 'dailyReports_houmu',// nippou-legal 内で知的財産室(dailyReports)と分ける

  branding: {
    appTitle: '日報アプリ',
    docTitle: '日報アプリ｜ホリゾン法務',
    company:  'HORIZON',
    dept:     '法務知財部 法務'
  },

  // GASバックアップは未使用。空にして生産管理課のGASエンドポイント継承を防ぐ（重要）。
  gasUrl: '',

  // 法務知財部のFirebaseプロジェクト（nippou-legal）を共用
  firebase: {
    apiKey:            'AIzaSyDLiYOfkKmPZ2rGFaUd1oTQek5XVK7hfMk',
    authDomain:        'nippou-legal.firebaseapp.com',
    projectId:         'nippou-legal',
    storageBucket:     'nippou-legal.firebasestorage.app',
    messagingSenderId: '802361483209',
    appId:             '1:802361483209:web:e4c2bd66666e93d4d45218'
  },

  allowedDomains: ['horizon.co.jp', 'bp.horizon.co.jp'],

  // 担当者選択の初期リスト。起動後は設定画面から各自追加も可能。
  defaultAuthors: ['深谷'],
  authorsLast: [],                 // 常に最後に並べる担当者（指定なし）
  kpis: [],                        // KPI（月次工数目標）は生産管理課のみ

  // 上部ナビの外部アプリリンク（黒板/手順書/資料検索）は生産管理課専用のため非表示。
  showAppLinks: false,

  // 入力時の演出は既定で OFF（設定画面から端末ごとに ON にできる）
  inputFxDefault: false,

  // 外部リンクは未設定。生産管理課のバックアップシート継承を防ぐため空に。
  links: { backupSheetUrl: '' },

  // 勤務ルールは生産管理課と同じ → rules は省略（既定値を継承）。

  // 初期カテゴリ（分類登録シートより。管理タブで編集可）
  defaultCats: [
    { cd:'A', label:'ガバナンス', color:'#3b82f6', subs:[
      { cd:'A1', label:'規則・規程の作成' },
      { cd:'A2', label:'規則・規程のレビュー' },
      { cd:'A3', label:'内部統制・監査の計画・実施' },
      { cd:'A4', label:'内部統制・監査の運営・管理' }
    ] },
    { cd:'B', label:'コンプラ', color:'#22c55e', subs:[
      { cd:'B1', label:'法令・規制調査' },
      { cd:'B2', label:'違反・不備への対応' },
      { cd:'B3', label:'教育・研修' },
      { cd:'B4', label:'通報・相談対応' }
    ] },
    { cd:'C', label:'契約法務', color:'#ef4444', subs:[
      { cd:'C1', label:'契約書作成' },
      { cd:'C2', label:'契約書レビュー' }
    ] },
    { cd:'D', label:'法務相談', color:'#a855f7', subs:[
      { cd:'D1', label:'社内問い合わせ' },
      { cd:'D2', label:'法的検討・回答作成' },
      { cd:'D3', label:'その他' }
    ] },
    { cd:'E', label:'知財', color:'#eab308', subs:[
      { cd:'E1', label:'検討' },
      { cd:'E2', label:'その他' }
    ] },
    { cd:'F', label:'グループ', color:'#06b6d4', subs:[
      { cd:'F1', label:'東阪電子機器' },
      { cd:'F2', label:'富士油圧精機' },
      { cd:'F3', label:'その他' }
    ] },
    { cd:'G', label:'打合せ', color:'#6b7280', subs:[
      { cd:'G1', label:'社内打ち合わせ' },
      { cd:'G2', label:'社外打ち合わせ' },
      { cd:'G3', label:'その他' }
    ] },
    { cd:'H', label:'その他', color:'#ec4899', subs:[
      { cd:'H1', label:'その他雑務' }
    ] }
  ]
};
