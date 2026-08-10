// 生産管理課（NP / ルート配置）用の設定。
// 既定値（index.html 内の DEPARTMENT_CONFIG）が生産管理課そのものなので、
// ここでは上書きしない（空オーバーライド）。このファイルは config.js の 404 を防ぐ役割。
// 他部署はサブフォルダ（例 /legal/config.js）に各自の設定を置く。
window.DEPARTMENT_CONFIG_OVERRIDE = {};
