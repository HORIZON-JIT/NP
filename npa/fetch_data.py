"""
NPA - GAS APIからデータ取得してDataFrameに変換

GAS Web Appは302リダイレクト後にJSONを返す。
requests はリダイレクトを自動追従するが、Google認証が必要な場合は
HTMLログインページが返ることがある。その場合はJSONP方式にフォールバックする。
"""

import json
import re
import requests
import pandas as pd
from config import GAS_URL, SUB_CATEGORIES, get_main_cd, get_main_label


def _try_parse_json(text: str) -> dict | None:
    """JSON or JSONP レスポンスをパースする"""
    text = text.strip()
    # 純粋なJSON
    if text.startswith("{"):
        return json.loads(text)
    # JSONP: callbackName({...})
    m = re.match(r'^[a-zA-Z_]\w*\((.+)\);?\s*$', text, re.DOTALL)
    if m:
        return json.loads(m.group(1))
    return None


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

    # ① 通常のGETリクエスト（リダイレクト自動追従）
    try:
        resp = requests.get(GAS_URL, params=params, timeout=60,
                            allow_redirects=True)
        resp.raise_for_status()
        data = _try_parse_json(resp.text)
        if data:
            if not data.get("ok"):
                raise RuntimeError(f"GAS APIエラー: {data.get('error', '不明')}")
            return data
    except json.JSONDecodeError:
        pass

    # ② JSONP形式で再試行（callbackパラメータ付き）
    params["callback"] = "cb"
    resp = requests.get(GAS_URL, params=params, timeout=60,
                        allow_redirects=True)
    resp.raise_for_status()
    data = _try_parse_json(resp.text)
    if data:
        if not data.get("ok"):
            raise RuntimeError(f"GAS APIエラー: {data.get('error', '不明')}")
        return data

    # レスポンス内容の先頭を表示してデバッグ支援
    preview = resp.text[:300] if resp.text else "(空のレスポンス)"
    raise RuntimeError(
        f"GASからJSONを取得できません。レスポンス先頭:\n{preview}"
    )


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
