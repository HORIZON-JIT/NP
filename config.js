// 生産管理課（NP / ルート配置）用の設定。
// 既定値（index.html 内の DEPARTMENT_CONFIG）が生産管理課そのものなので、
// ここでは生産管理課だけに必要な項目のみ上書きする。
// 他部署はサブフォルダ（例 /legal/config.js）に各自の設定を置く。
window.DEPARTMENT_CONFIG_OVERRIDE = {
  // タスク管理アプリ「タスク実績」読取用GAS（gas/task-sheet-reader.gs をデプロイしたもの）。
  // 「タスク実績から取込」ボタンと管理タブのタスク対応表はこのURLがある部署だけに表示。
  taskSheetGasUrl: 'https://script.google.com/a/macros/horizon.co.jp/s/AKfycbwxhtGBpHgb5Iz9FLydd5oxwtYgCQqMEbM_KudCxNTZAo61xL0SLJfokqDE-gW2Ze6w/exec'
};
