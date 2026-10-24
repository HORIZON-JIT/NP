// 生産管理課（NP / ルート配置）用の設定。
// 既定値（index.html 内の DEPARTMENT_CONFIG）が生産管理課そのものなので、
// ここでは生産管理課だけに必要な項目のみ上書きする。
// 他部署はサブフォルダ（例 /legal/config.js）に各自の設定を置く。
window.DEPARTMENT_CONFIG_OVERRIDE = {
  // タスク管理アプリ「タスク実績」読取用GAS（gas/task-sheet-reader.gs をデプロイしたもの）。
  // 「タスク実績から取込」ボタンと管理タブのタスク対応表はこのURLがある部署だけに表示。
  taskSheetGasUrl: 'https://script.google.com/a/macros/horizon.co.jp/s/AKfycbwxhtGBpHgb5Iz9FLydd5oxwtYgCQqMEbM_KudCxNTZAo61xL0SLJfokqDE-gW2Ze6w/exec',
  // 月次AI分析用GAS（gas/ai-monthly-analyzer.gs をデプロイしたもの）。生産管理課だけで使う。
  // URL を入れると月次ダッシュボードに「AI分析」カードが出る。空なら非表示。
  aiAnalysisGasUrl: 'https://script.google.com/a/macros/horizon.co.jp/s/AKfycbx_vALcuIZVtAlN5oX4bFbBmcstz-M2HB45UXpSzUdRHPSBHRjeLCBshGbPZFchQcfj/exec',
  // 入力時の演出: 生産管理課の既定は ON・メガフィーバー・効果音（ドン・カッ）あり
  // （設定画面から端末ごとに変更でき、変更した端末ではその設定が優先される）
  inputFxDefault: true,
  inputFxMode: 'mega',
  taikoSoundDefault: true,
  // 全端末で1回だけ演出設定を上書き（id を変えると、もう一度だけ上書きされる）
  inputFxForce: { id: '2026-10-mega', on: true, mode: 'mega', sound: true },

  // Folio v2（Firebase 保存版）。2026/10/23 の Folio 移行に合わせて 'firebase' に切替済み（全員が v2 から読む）。
  // 'sheet' に戻すと v1（スプレッドシート＋上の taskSheetGasUrl。移行後は更新されない）を読む。端末ごとの試験切替は設定画面から。
  // 日報中分類の対応は Folio v2 の nippouMap/taskCd に保存。
  folio: {
    source: 'firebase',
    firebase: {
      apiKey:            'AIzaSyBLN5pxyLzFULeeCzV4dEUBAOZkMuLJBpA',
      authDomain:        'folio-73ebf.firebaseapp.com',
      projectId:         'folio-73ebf',
      messagingSenderId: '380142786013',
      appId:             '1:380142786013:web:cfff637333976eae445234'
    },
    mapCollection: 'nippouMap',
    mapDoc: 'taskCd'
  }
};
