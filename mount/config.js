// ═══════════════════════════════════════════════════════════════
//  電装課 用 設定（/NP/mount/ に配置）
// ═══════════════════════════════════════════════════════════════
//  index.html は生産管理課と共通（単一ソース）。この config.js だけが
//  この派生アプリ固有の値を上書きする。未記載の項目は生産管理課の既定値を継承。
//
//  Firebaseプロジェクト: nippou-legal（法務知財部と共用）。
//  コレクションを 'dailyReports_mount' に分けることで、法務知財部（/legal/・/law/）の
//  日報・マスタとは完全に別のデータとして保存する。
//
//  ※ カテゴリ(defaultCats)は「コレクションが空のときの初期の種」。
//    起動後は「管理」タブでいつでも追加/編集/並べ替え/有効無効が可能。
//    （2026/10 電装課記入の分類登録シートから作成）
// ═══════════════════════════════════════════════════════════════

window.DEPARTMENT_CONFIG_OVERRIDE = {
  deptId: 'mount',                  // localStorageキー / マスタdoc id のプレフィックス
  collection: 'dailyReports_mount',  // nippou-legal 内で法務知財部のデータと分ける

  branding: {
    appTitle: '日報アプリ',
    docTitle: '日報アプリ｜ホリゾン電装',
    company:  'HORIZON',
    dept:     '電装課'
  },

  // GASバックアップは未使用。空にして生産管理課のGASエンドポイント継承を防ぐ（重要）。
  gasUrl: '',

  // 法務知財部のFirebaseプロジェクト（nippou-legal）を共用（コレクションは別）
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
  defaultAuthors: ['吉田', '山田', '河波'],
  authorsLast: [],                 // 常に最後に並べる担当者（指定なし）
  kpis: [],                        // KPI（月次工数目標）は生産管理課のみ

  // 上部ナビの外部アプリリンク（黒板/手順書/資料検索）は生産管理課専用のため非表示。
  showAppLinks: false,

  // 入力時の演出は既定で ON・ネオン（設定画面から端末ごとに変更できる）
  inputFxDefault: true,
  inputFxMode: 'neon',

  // 外部リンクは未設定。生産管理課のバックアップシート継承を防ぐため空に。
  links: { backupSheetUrl: '' },

  // 勤務ルールは生産管理課と同じ → rules は省略（既定値を継承）。

  // 初期カテゴリ（分類登録シートより。管理タブで編集可）
  defaultCats: [
    { cd:'A', label:'管理', color:'#3b82f6', subs:[
      { cd:'A1', label:'社外向け業務：事務処理' },
      { cd:'A2', label:'社内向け業務：事務処理：資料作成' },
      { cd:'A3', label:'内外作切替' },
      { cd:'A4', label:'見積り業務' },
      { cd:'A5', label:'手配方式：構成見直し' },
      { cd:'A6', label:'人員計画' },
      { cd:'A7', label:'部材管理' },
      { cd:'A8', label:'HAR予定表管理' },
      { cd:'A9', label:'ACUT予定管理' }
    ] },
    { cd:'B', label:'作業準備', color:'#22c55e', subs:[
      { cd:'B1', label:'工程付け' },
      { cd:'B2', label:'工数計算' },
      { cd:'B3', label:'結線表作成' },
      { cd:'B4', label:'作業用図面作成' },
      { cd:'B5', label:'線切表作成' },
      { cd:'B6', label:'ボード図印刷' }
    ] },
    { cd:'C', label:'打合せ', color:'#6b7280', subs:[
      { cd:'C1', label:'打合せ資料準備' },
      { cd:'C2', label:'打合せ' },
      { cd:'C3', label:'問い合わせ対応' }
    ] },
    { cd:'D', label:'教育・研修', color:'#a855f7', subs:[
      { cd:'D1', label:'電線加工教育' },
      { cd:'D2', label:'社内研修' }
    ] },
    { cd:'E', label:'改善', color:'#eab308', subs:[
      { cd:'E1', label:'QC活動' },
      { cd:'E2', label:'ツール(導入、習得)' },
      { cd:'E3', label:'業務' }
    ] },
    { cd:'F', label:'不適合処理', color:'#ef4444', subs:[
      { cd:'F1', label:'集計' },
      { cd:'F2', label:'追跡' },
      { cd:'F3', label:'外注へ連絡' },
      { cd:'F4', label:'対策処理' },
      { cd:'F5', label:'Notes入力' },
      { cd:'F6', label:'不具合修正：対応' }
    ] },
    { cd:'G', label:'アフター', color:'#f97316', subs:[
      { cd:'G1', label:'対応' }
    ] },
    { cd:'H', label:'その他', color:'#ec4899', subs:[
      { cd:'H1', label:'目標設定' },
      { cd:'H2', label:'全体朝礼' },
      { cd:'H3', label:'GW対応' },
      { cd:'H4', label:'週次昼礼' },
      { cd:'H5', label:'その他' }
    ] },
    { cd:'I', label:'グループ会社', color:'#06b6d4', subs:[
      { cd:'I1', label:'東阪電子機器' }
    ] }
  ]
};
