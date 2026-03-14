"""
NPA - GAS APIからデータ取得してDataFrameに変換
"""

import json
import requests
import pandas as pd
from config import GAS_URL, SUB_CATEGORIES, get_main_cd, get_main_label


def fetch_date_range(start_date: str, end_date: str) -> dict:
    """
    GAS getDateRange APIを呼び出し、生JSONを返す。

    Parameters:
        start_date: 開始日 (YYYY-MM-DD)
        end_date:   終了日 (YYYY-MM-DD)

    Returns:
        GASレスポンスのdict
    """
    params = {
        "action": "getDateRange",
        "startDate": start_date,
        "endDate": end_date,
        "noCache": "1",
    }
    resp = requests.get(GAS_URL, params=params, timeout=60)
    resp.raise_for_status()
    data = resp.json()
    if not data.get("ok"):
        raise RuntimeError(f"GAS APIエラー: {data.get('error', '不明')}")
    return data


def to_dataframe(data: dict) -> pd.DataFrame:
    """
    GAS getDateRange レスポンスをDataFrameに変換。

    columns: author, cdSub, subLabel, mainCd, mainLabel, hoursNormal, hoursOT, hoursTotal
    """
    rows = data.get("rows", [])
    if not rows:
        return pd.DataFrame()

    df = pd.DataFrame(rows)

    # カラム名の正規化
    col_map = {}
    if "hoursNormal" not in df.columns and "hours_normal" in df.columns:
        col_map["hours_normal"] = "hoursNormal"
    if "hoursOT" not in df.columns and "hours_ot" in df.columns:
        col_map["hours_ot"] = "hoursOT"
    if col_map:
        df = df.rename(columns=col_map)

    # 数値変換
    for col in ["hoursNormal", "hoursOT"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    # 合計時間
    df["hoursTotal"] = df.get("hoursNormal", 0) + df.get("hoursOT", 0)

    # マスタ情報付与
    df["subLabel"] = df["cdSub"].map(
        lambda cd: SUB_CATEGORIES.get(cd, {}).get("label", cd)
    )
    df["mainCd"] = df["cdSub"].map(get_main_cd)
    df["mainLabel"] = df["cdSub"].map(get_main_label)

    return df


def fetch_as_dataframe(start_date: str, end_date: str) -> pd.DataFrame:
    """データ取得→DataFrame変換をまとめて行う"""
    data = fetch_date_range(start_date, end_date)
    df = to_dataframe(data)
    return df


def get_leave_map(data: dict) -> dict:
    """
    GASレスポンスから退勤時間マップを取得。

    Returns:
        {author: {date: "HH:MM", ...}, ...}
    """
    return data.get("leaveMap", {})
