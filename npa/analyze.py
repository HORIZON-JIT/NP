"""
NPA - 集計・分析ロジック
"""

import pandas as pd
from config import MAIN_CATEGORIES, AUTHORS


def by_author(df: pd.DataFrame) -> pd.DataFrame:
    """
    担当者別の工数集計。

    Returns:
        columns: author, hoursNormal, hoursOT, hoursTotal
    """
    if df.empty:
        return pd.DataFrame(columns=["author", "hoursNormal", "hoursOT", "hoursTotal"])

    result = (
        df.groupby("author", as_index=False)
        .agg(hoursNormal=("hoursNormal", "sum"),
             hoursOT=("hoursOT", "sum"))
    )
    result["hoursTotal"] = result["hoursNormal"] + result["hoursOT"]

    # 定義順でソート
    author_order = {a: i for i, a in enumerate(AUTHORS)}
    result["_order"] = result["author"].map(lambda a: author_order.get(a, 99))
    result = result.sort_values("_order").drop(columns="_order").reset_index(drop=True)
    return result


def by_main_category(df: pd.DataFrame) -> pd.DataFrame:
    """
    大分類別の工数集計。

    Returns:
        columns: mainCd, mainLabel, hoursNormal, hoursOT, hoursTotal
    """
    if df.empty:
        return pd.DataFrame(columns=["mainCd", "mainLabel", "hoursNormal", "hoursOT", "hoursTotal"])

    result = (
        df.groupby(["mainCd", "mainLabel"], as_index=False)
        .agg(hoursNormal=("hoursNormal", "sum"),
             hoursOT=("hoursOT", "sum"))
    )
    result["hoursTotal"] = result["hoursNormal"] + result["hoursOT"]

    # マスタ順でソート
    cat_order = list(MAIN_CATEGORIES.keys())
    result["_order"] = result["mainCd"].map(lambda c: cat_order.index(c) if c in cat_order else 99)
    result = result.sort_values("_order").drop(columns="_order").reset_index(drop=True)
    return result


def by_sub_category(df: pd.DataFrame) -> pd.DataFrame:
    """
    中分類別の工数集計。

    Returns:
        columns: cdSub, subLabel, mainCd, mainLabel, hoursNormal, hoursOT, hoursTotal
    """
    if df.empty:
        return pd.DataFrame(columns=["cdSub", "subLabel", "mainCd", "mainLabel",
                                      "hoursNormal", "hoursOT", "hoursTotal"])

    result = (
        df.groupby(["cdSub", "subLabel", "mainCd", "mainLabel"], as_index=False)
        .agg(hoursNormal=("hoursNormal", "sum"),
             hoursOT=("hoursOT", "sum"))
    )
    result["hoursTotal"] = result["hoursNormal"] + result["hoursOT"]
    result = result.sort_values("cdSub").reset_index(drop=True)
    return result


def by_author_and_category(df: pd.DataFrame) -> pd.DataFrame:
    """
    担当者×大分類のクロス集計。

    Returns:
        columns: author, mainCd, mainLabel, hoursTotal
    """
    if df.empty:
        return pd.DataFrame(columns=["author", "mainCd", "mainLabel", "hoursTotal"])

    result = (
        df.groupby(["author", "mainCd", "mainLabel"], as_index=False)
        .agg(hoursTotal=("hoursTotal", "sum"))
    )
    return result


def overtime_by_author(df: pd.DataFrame) -> pd.DataFrame:
    """
    担当者別の残業時間集計。

    Returns:
        columns: author, hoursOT
    """
    if df.empty:
        return pd.DataFrame(columns=["author", "hoursOT"])

    result = (
        df.groupby("author", as_index=False)
        .agg(hoursOT=("hoursOT", "sum"))
    )
    author_order = {a: i for i, a in enumerate(AUTHORS)}
    result["_order"] = result["author"].map(lambda a: author_order.get(a, 99))
    result = result.sort_values("_order").drop(columns="_order").reset_index(drop=True)
    return result


def summary_text(df: pd.DataFrame, start_date: str, end_date: str) -> str:
    """期間サマリーをテキストで生成"""
    if df.empty:
        return f"期間 {start_date} 〜 {end_date}: データなし"

    total_normal = df["hoursNormal"].sum()
    total_ot = df["hoursOT"].sum()
    total = total_normal + total_ot
    n_authors = df["author"].nunique()

    lines = [
        f"=== 期間サマリー: {start_date} 〜 {end_date} ===",
        f"担当者数: {n_authors}名",
        f"通常時間合計: {total_normal:.1f}h",
        f"残業時間合計: {total_ot:.1f}h",
        f"総工数: {total:.1f}h",
        "",
    ]

    # 担当者別
    auth_df = by_author(df)
    lines.append("【担当者別】")
    for _, row in auth_df.iterrows():
        lines.append(f"  {row['author']}: 通常{row['hoursNormal']:.1f}h "
                      f"残業{row['hoursOT']:.1f}h 計{row['hoursTotal']:.1f}h")
    lines.append("")

    # 大分類別
    cat_df = by_main_category(df)
    lines.append("【大分類別】")
    for _, row in cat_df.iterrows():
        pct = row["hoursTotal"] / total * 100 if total > 0 else 0
        lines.append(f"  {row['mainLabel']}({row['mainCd']}): "
                      f"{row['hoursTotal']:.1f}h ({pct:.1f}%)")

    return "\n".join(lines)
