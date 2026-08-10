// ═══════════════════════════════════════════════════════════════
//  部署別 設定テンプレート（例: 法務知財部）
// ═══════════════════════════════════════════════════════════════
// このアプリ(index.html)は「共通コード + 設定分離」方式です。
// 部署ごとに別Firebaseプロジェクト・別公開URLを用意し、
// この config.js だけを差し替えて横展開します（index.html本体は共通のまま）。
//
// 【使い方】
//  1) このファイルを config.js という名前でコピー
//  2) 下記の値を新部署用に書き換え（新Firebaseプロジェクトの設定など）
//  3) index.html の <head> 内、firebase の <script> 群より後・
//     最初の <script>（メインスクリプト）より前に、次の1行を追加:
//        <script src="config.js"></script>
//
// 【ポイント】
//  ・記載した項目だけが上書きされ、未記載は生産管理課の既定値が使われます。
//    → 勤務ルールが同じなら rules を丸ごと省略してOK。
//  ・カテゴリは起動後「管理」タブでいつでも編集できます。
//    defaultCats は「新Firebaseプロジェクトが空のときの初期の種」です。
//  ・データ分離は "別Firebaseプロジェクト" で担保します（下記 firebase を必ず差し替え）。
// ═══════════════════════════════════════════════════════════════

window.DEPARTMENT_CONFIG_OVERRIDE = {
  deptId: 'legal',                 // 部署ごとに一意。localStorageキー/マスタdoc idのプレフィックス
  collection: 'dailyReports',      // 別Firebaseプロジェクトなら 'dailyReports' のままでOK

  branding: {
    appTitle: '日報アプリ',
    docTitle: '日報アプリ｜ホリゾン法務知財',
    company:  'HORIZON',
    dept:     '法務知財部'
  },

  // GASバックアップを使う場合のみ設定（未使用なら空文字でOK）
  gasUrl: '',

  // ★ 法務知財部用の新Firebaseプロジェクトの設定に差し替える（必須）
  firebase: {
    apiKey:            'YOUR_API_KEY',
    authDomain:        'YOUR_PROJECT.firebaseapp.com',
    projectId:         'YOUR_PROJECT',
    storageBucket:     'YOUR_PROJECT.firebasestorage.app',
    messagingSenderId: 'YOUR_SENDER_ID',
    appId:             'YOUR_APP_ID'
  },

  // 同じ会社ならそのまま。別ドメインを含む場合は追加。
  allowedDomains: ['horizon.co.jp', 'bp.horizon.co.jp'],

  // 担当者選択の初期リスト（この部署のメンバー名）。各自は設定画面から追加も可能。
  defaultAuthors: ['中西'],

  // 上部ナビの外部アプリリンク（黒板/手順書/資料検索）は生産管理課専用。
  // 他部署では非表示にする場合 false（既定は true = 表示）。
  showAppLinks: false,

  // 勤務ルールが生産管理課と同じなら、この rules ブロックは丸ごと削除してOK（既定値が使われます）。
  // 異なる場合のみ、必要な項目を指定:
  // rules: {
  //   workStart: '08:30',   // 始業
  //   normalEnd: '17:20',   // 定時
  //   otStart:   '17:30',   // 残業開始
  //   roundStepMin: 15,     // 丸め単位(分)
  //   roundStepHours: 0.25, // 丸め単位(時間)
  //   lunch: {              // 昼休憩（奇数月/偶数月。固定なら両方同値に）
  //     odd:  { start: '12:40', end: '13:30' },
  //     even: { start: '12:20', end: '13:10' }
  //   }
  // },

  // 初期カテゴリ（例。実際の業務に合わせて「管理」タブで自由に編集できます）
  defaultCats: [
    { cd:'P', label:'特許', color:'#3b82f6', subs:[{cd:'P1',label:'出願準備'},{cd:'P2',label:'中間対応'},{cd:'P3',label:'権利化後管理'}] },
    { cd:'S', label:'商標', color:'#22c55e', subs:[{cd:'S1',label:'出願'},{cd:'S2',label:'更新'},{cd:'S3',label:'調査'}] },
    { cd:'C', label:'契約', color:'#ef4444', subs:[{cd:'C1',label:'審査'},{cd:'C2',label:'作成'},{cd:'C3',label:'交渉'}] },
    { cd:'L', label:'係争', color:'#a855f7', subs:[{cd:'L1',label:'訴訟対応'},{cd:'L2',label:'調査'}] },
    { cd:'R', label:'相談', color:'#eab308', subs:[{cd:'R1',label:'法務相談'}] },
    { cd:'O', label:'その他', color:'#6b7280', subs:[{cd:'O1',label:'その他'}] }
  ]
};
