"""
NPA - グラフ生成 (matplotlib)
"""

import os
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
import pandas as pd
from config import MAIN_CATEGORIES, get_main_color

# 日本語フォント設定
_JP_FONTS = [
    "/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc",
    "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
    "/usr/share/fonts/noto-cjk/NotoSansCJK-Regular.ttc",
]

def _setup_font():
    """利用可能な日本語フォントを探してmatplotlibに設定"""
    for path in _JP_FONTS:
        if os.path.exists(path):
            fm.fontManager.addfont(path)
            prop = fm.FontProperties(fname=path)
            plt.rcParams["font.family"] = prop.get_name()
            return
    # フォールリバック: sans-serif
    plt.rcParams["font.family"] = "sans-serif"

_setup_font()

OUTPUT_DIR = os.path.join(os.path.dirname(__file__), "output")
os.makedirs(OUTPUT_DIR, exist_ok=True)


def _savefig(fig, filename: str) -> str:
    """グラフを保存してパスを返す"""
    path = os.path.join(OUTPUT_DIR, filename)
    fig.savefig(path, dpi=150, bbox_inches="tight")
    plt.close(fig)
    return path


def bar_by_author(auth_df: pd.DataFrame, title: str = "担当者別工数") -> str:
    """
    担当者別の通常/残業 積み上げ棒グラフ。
    """
    if auth_df.empty:
        return ""

    fig, ax = plt.subplots(figsize=(10, 5))
    x = range(len(auth_df))
    ax.bar(x, auth_df["hoursNormal"], label="通常", color="#3b82f6")
    ax.bar(x, auth_df["hoursOT"], bottom=auth_df["hoursNormal"],
           label="残業", color="#ef4444")
    ax.set_xticks(x)
    ax.set_xticklabels(auth_df["author"])
    ax.set_ylabel("時間 (h)")
    ax.set_title(title)
    ax.legend()

    # 合計値をバー上に表示
    for i, row in auth_df.iterrows():
        ax.text(i, row["hoursTotal"] + 0.3, f'{row["hoursTotal"]:.1f}',
                ha="center", va="bottom", fontsize=9)

    return _savefig(fig, "bar_by_author.png")


def pie_by_category(cat_df: pd.DataFrame, title: str = "大分類別工数比率") -> str:
    """
    大分類別の円グラフ。
    """
    if cat_df.empty:
        return ""

    cat_df = cat_df[cat_df["hoursTotal"] > 0].copy()
    if cat_df.empty:
        return ""

    colors = [MAIN_CATEGORIES.get(cd, {}).get("color", "#999")
              for cd in cat_df["mainCd"]]
    labels = [f'{row["mainLabel"]}({row["mainCd"]})' for _, row in cat_df.iterrows()]

    fig, ax = plt.subplots(figsize=(8, 8))
    wedges, texts, autotexts = ax.pie(
        cat_df["hoursTotal"],
        labels=labels,
        colors=colors,
        autopct="%1.1f%%",
        startangle=90,
    )
    ax.set_title(title)
    return _savefig(fig, "pie_by_category.png")


def bar_by_category(cat_df: pd.DataFrame, title: str = "大分類別工数") -> str:
    """
    大分類別の棒グラフ。
    """
    if cat_df.empty:
        return ""

    colors = [MAIN_CATEGORIES.get(cd, {}).get("color", "#999")
              for cd in cat_df["mainCd"]]

    fig, ax = plt.subplots(figsize=(10, 5))
    x = range(len(cat_df))
    ax.bar(x, cat_df["hoursTotal"], color=colors)
    ax.set_xticks(x)
    ax.set_xticklabels([f'{r["mainLabel"]}' for _, r in cat_df.iterrows()])
    ax.set_ylabel("時間 (h)")
    ax.set_title(title)

    for i, row in cat_df.iterrows():
        ax.text(i, row["hoursTotal"] + 0.3, f'{row["hoursTotal"]:.1f}',
                ha="center", va="bottom", fontsize=9)

    return _savefig(fig, "bar_by_category.png")


def bar_overtime(ot_df: pd.DataFrame, title: str = "担当者別残業時間") -> str:
    """
    担当者別の残業時間棒グラフ。
    """
    if ot_df.empty:
        return ""

    fig, ax = plt.subplots(figsize=(10, 5))
    x = range(len(ot_df))
    ax.bar(x, ot_df["hoursOT"], color="#ef4444")
    ax.set_xticks(x)
    ax.set_xticklabels(ot_df["author"])
    ax.set_ylabel("残業時間 (h)")
    ax.set_title(title)

    for i, row in ot_df.iterrows():
        if row["hoursOT"] > 0:
            ax.text(i, row["hoursOT"] + 0.2, f'{row["hoursOT"]:.1f}',
                    ha="center", va="bottom", fontsize=9)

    return _savefig(fig, "bar_overtime.png")


def stacked_author_category(cross_df: pd.DataFrame,
                             title: str = "担当者×大分類 工数内訳") -> str:
    """
    担当者ごとに大分類で積み上げた棒グラフ。
    """
    if cross_df.empty:
        return ""

    pivot = cross_df.pivot_table(
        index="author", columns="mainCd", values="hoursTotal", fill_value=0
    )
    # マスタ順にカラムをソート
    cat_order = [c for c in MAIN_CATEGORIES.keys() if c in pivot.columns]
    pivot = pivot[cat_order]

    fig, ax = plt.subplots(figsize=(12, 6))
    bottom = pd.Series(0.0, index=pivot.index)
    for col in pivot.columns:
        color = MAIN_CATEGORIES.get(col, {}).get("color", "#999")
        label = MAIN_CATEGORIES.get(col, {}).get("label", col)
        ax.bar(pivot.index, pivot[col], bottom=bottom, label=label, color=color)
        bottom += pivot[col]

    ax.set_ylabel("時間 (h)")
    ax.set_title(title)
    ax.legend(bbox_to_anchor=(1.05, 1), loc="upper left")
    return _savefig(fig, "stacked_author_category.png")
