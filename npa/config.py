"""
NPA (日報分析) - 設定ファイル
GAS URL・カテゴリマスタ定義
"""

# GAS Web App URL
GAS_URL = (
    "https://script.google.com/a/macros/horizon.co.jp/s/"
    "AKfycbxx96D5Lz_I3B_L-Bia9WftcUP6Em-l6mTulefZBB6Xu9BEITm0DLroTwB1CHjLkkL6/exec"
)

# 担当者リスト（表示順）
AUTHORS = ["秋永", "吉永", "山田", "須藤", "平山", "小野", "田村", "谷", "ヘルプ"]

# 大分類マスタ
MAIN_CATEGORIES = {
    "K": {"label": "計画",       "color": "#3b82f6"},
    "T": {"label": "手配",       "color": "#22c55e"},
    "J": {"label": "受注",       "color": "#ef4444"},
    "B": {"label": "BOMメンテ",  "color": "#a855f7"},
    "S": {"label": "打合せ",     "color": "#eab308"},
    "Z": {"label": "改善",       "color": "#06b6d4"},
    "O": {"label": "その他",     "color": "#6b7280"},
    "A": {"label": "アフター",   "color": "#ec4899"},
}

# 中分類マスタ
SUB_CATEGORIES = {
    "K1": {"label": "新規計画",       "main": "K"},
    "K2": {"label": "計画調整",       "main": "K"},
    "K3": {"label": "進捗確認",       "main": "K"},
    "K4": {"label": "部品進捗確認",   "main": "K"},
    "T1": {"label": "計画手配",       "main": "T"},
    "T2": {"label": "計画外手配",     "main": "T"},
    "T3": {"label": "試作手配",       "main": "T"},
    "J1": {"label": "受注確認/引当",  "main": "J"},
    "J2": {"label": "OEM商品受注出荷","main": "J"},
    "B1": {"label": "部品表メンテナンス","main": "B"},
    "B2": {"label": "設計変更",       "main": "B"},
    "B3": {"label": "原価集計",       "main": "B"},
    "S1": {"label": "打合せ",         "main": "S"},
    "S2": {"label": "問い合わせ対応", "main": "S"},
    "Z1": {"label": "課内",           "main": "Z"},
    "Z2": {"label": "課外",           "main": "Z"},
    "Z3": {"label": "スタフェス対応", "main": "Z"},
    "O1": {"label": "その他",         "main": "O"},
    "A1": {"label": "アフター改善",   "main": "A"},
    "A2": {"label": "価格登録",       "main": "A"},
    "A3": {"label": "在庫管理",       "main": "A"},
    "A4": {"label": "受注対応",       "main": "A"},
    "A5": {"label": "特注対応",       "main": "A"},
}

# 所定労働時間
REGULAR_HOURS = 8.0

# 締め日（15日締め = 16日〜翌15日が1ヶ月度）
CLOSING_DAY = 15


def get_main_cd(cd_sub: str) -> str:
    """中分類コードから大分類コードを取得"""
    return SUB_CATEGORIES.get(cd_sub, {}).get("main", "?")


def get_main_label(cd_sub: str) -> str:
    """中分類コードから大分類ラベルを取得"""
    main_cd = get_main_cd(cd_sub)
    return MAIN_CATEGORIES.get(main_cd, {}).get("label", "不明")


def get_main_color(cd_sub: str) -> str:
    """中分類コードから大分類カラーを取得"""
    main_cd = get_main_cd(cd_sub)
    return MAIN_CATEGORIES.get(main_cd, {}).get("color", "#999999")
