"""
NPA - GAS APIからデータ取得してDataFrameに変換

GAS Web App（組織内限定）へのアクセスにはGoogle OAuth2認証が必要。
access_token をクエリパラメータとして付与してリクエストする。
"""

import json
import re
import requests
import pandas as pd
from config import GAS_URL, SUB_CATEGORIES, get_main_cd, get_main_label
from gas_auth import get_access_token


def _try_parse_json(text: str) -> dict | None:
    """JSON or JSONP レスポンスをパースする"""
    text = text.strip()
    if not text:
        return None
    # 純粋なJSON
    if text.startswith("{") or text.startswith("["):
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
    # OAuth2アクセストークンを取得
    token = get_access_token()

    params = {
        "action": "getDateRange",
        "startDate": start_date,
        "endDate": end_date,
        "noCache": "1",
        "access_token": token,
    }

    resp = requests.get(GAS_URL, params=params, timeout=60,
                        allow_redirects=True)
    resp.raise_for_status()

    data = _try_parse_json(resp.text)
    if data is None:
        # レスポンスがHTMLの場合（認証エラーの可能性）
        if "<html" in resp.text.lower()[:200]:
            raise RuntimeError(
                "GAS認証エラー: ログインページが返されました。\n"
                "トークンが期限切れの可能性があります。\n"
                "python npa/gas_auth.py を実行して再認証してください。"
            )
        preview = resp.text[:300] if resp.text else "(空のレスポンス)"
        raise RuntimeError(f"GASからJSONを取得できません。レスポンス先頭:\n{preview}")

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
