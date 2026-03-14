"""
NPA - GAS APIからデータ取得してDataFrameに変換

GAS Web App（組織内限定）へのアクセスにはGoogle OAuth2認証が必要。
Authorization: Bearer ヘッダーで認証し、リダイレクトを手動で処理する。
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


def _fetch_with_auth(url: str, params: dict, token: str, max_redirects: int = 10) -> requests.Response:
    """
    Authorization: Bearer ヘッダー付きでGETリクエストを送信。
    GAS Web Appはリダイレクト（302）するため、手動でリダイレクトを辿り
    各リクエストにAuthorizationヘッダーを再付与する。
    """
    headers = {"Authorization": f"Bearer {token}"}

    resp = requests.get(url, params=params, headers=headers,
                        allow_redirects=False, timeout=60)

    for _ in range(max_redirects):
        if resp.status_code not in (301, 302, 303, 307, 308):
            break
        redirect_url = resp.headers.get("Location")
        if not redirect_url:
            break
        resp = requests.get(redirect_url, headers=headers,
                            allow_redirects=False, timeout=60)

    return resp


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
    }

    resp = _fetch_with_auth(GAS_URL, params, token)
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


def fetch_weekly_breakdown(start_date: str, end_date: str) -> pd.DataFrame:
    """
    期間を週ごとに分割してデータ取得。week_start列を付与。
    推移分析用。
    """
    from datetime import date as dt_date, timedelta

    start = dt_date.fromisoformat(start_date)
    end = dt_date.fromisoformat(end_date)

    all_dfs = []
    # 週の開始を月曜に揃える
    current = start - timedelta(days=start.weekday())

    while current <= end:
        ws = max(current, start)
        we = min(current + timedelta(days=6), end)

        try:
            raw = fetch_date_range(ws.isoformat(), we.isoformat())
            df = to_dataframe(raw)
            if not df.empty:
                df["week_start"] = ws.isoformat()
                all_dfs.append(df)
        except Exception:
            pass

        current += timedelta(days=7)

    if not all_dfs:
        return pd.DataFrame()
    return pd.concat(all_dfs, ignore_index=True)


def fetch_monthly_breakdown(start_date: str, end_date: str) -> pd.DataFrame:
    """
    期間を月ごとに分割してデータ取得。year, month列を付与。
    月別比較・前年比用。
    """
    import calendar
    from datetime import date as dt_date

    start = dt_date.fromisoformat(start_date)
    end = dt_date.fromisoformat(end_date)

    all_dfs = []
    year, month = start.year, start.month

    while (year, month) <= (end.year, end.month):
        _, last_day = calendar.monthrange(year, month)
        ms = dt_date(year, month, 1)
        me = dt_date(year, month, last_day)
        # 期間でクリップ
        ms = max(ms, start)
        me = min(me, end)

        try:
            raw = fetch_date_range(ms.isoformat(), me.isoformat())
            df = to_dataframe(raw)
            if not df.empty:
                df["year"] = year
                df["month"] = month
                all_dfs.append(df)
        except Exception:
            pass

        # 次の月
        if month == 12:
            year += 1
            month = 1
        else:
            month += 1

    if not all_dfs:
        return pd.DataFrame()
    return pd.concat(all_dfs, ignore_index=True)


def get_fiscal_month(d, closing_day: int = 15) -> tuple[int, int]:
    """
    日付が属する月度を返す（15日締めの場合）。

    15日締め: 前月16日〜当月15日 = 当月度
      例: 12/1 → 12月度, 12/16 → 1月度, 1/15 → 1月度, 3/14 → 3月度

    Returns:
        (year, month) の月度
    """
    if d.day <= closing_day:
        return (d.year, d.month)
    else:
        if d.month == 12:
            return (d.year + 1, 1)
        else:
            return (d.year, d.month + 1)


def get_fiscal_month_range(year: int, month: int, closing_day: int = 15):
    """
    月度の開始日・終了日を返す。

    例: 2026年3月度（15日締め） → 2026/2/16 〜 2026/3/15

    Returns:
        (start_date, end_date)
    """
    import calendar
    from datetime import date as dt_date

    # 開始: 前月の closing_day + 1
    if month == 1:
        start = dt_date(year - 1, 12, closing_day + 1)
    else:
        start = dt_date(year, month - 1, closing_day + 1)

    # 終了: 当月の closing_day（月末を超えないよう調整）
    _, last_day = calendar.monthrange(year, month)
    end_day = min(closing_day, last_day)
    end = dt_date(year, month, end_day)

    return (start, end)


def get_fiscal_year(d, closing_day: int = 15, fy_start_month: int = 10) -> int:
    """
    日付が属する年度を返す。

    10月始まり・15日締めの場合:
      2025/9/16〜2026/9/15 = 2025年度
      2025/3/14 → 3月度 → 2024年度
      2025/10/1 → 10月度 → 2025年度

    Returns:
        年度（int）
    """
    _, fm = get_fiscal_month(d, closing_day)
    y = d.year if d.day <= closing_day else (d.year if d.month < 12 else d.year + 1)
    # get_fiscal_month が返す year を使う
    fy_year, fy_month = get_fiscal_month(d, closing_day)
    if fy_month >= fy_start_month:
        return fy_year
    else:
        return fy_year - 1


def get_fiscal_year_range(fiscal_year: int, closing_day: int = 15, fy_start_month: int = 10):
    """
    年度の日付範囲を返す。

    例: 2025年度（10月始まり・15日締め） → 2025/9/16 〜 2026/9/15

    Returns:
        (start_date, end_date)
    """
    # 年度開始 = 開始月度の範囲の開始日
    start = get_fiscal_month_range(fiscal_year, fy_start_month, closing_day)[0]

    # 年度終了 = 最終月度(fy_start_month - 1)の範囲の終了日
    end_month = fy_start_month - 1 if fy_start_month > 1 else 12
    end_year = fiscal_year + 1 if fy_start_month > 1 else fiscal_year
    end = get_fiscal_month_range(end_year, end_month, closing_day)[1]

    return (start, end)


def _fiscal_month_order(fy_start_month: int = 10) -> list[int]:
    """年度の月度順序を返す。例: 10月始まり → [10,11,12,1,2,...,9]"""
    return [(fy_start_month + i - 1) % 12 + 1 for i in range(12)]


def get_fiscal_quarter_range(fiscal_year: int, quarter: int,
                             closing_day: int = 15, fy_start_month: int = 10):
    """
    四半期の日付範囲を返す。

    10月始まり: Q1=10,11,12月度 / Q2=1,2,3月度 / Q3=4,5,6月度 / Q4=7,8,9月度

    Returns:
        (start_date, end_date)
    """
    ordered = _fiscal_month_order(fy_start_month)
    q_months = ordered[(quarter - 1) * 3 : quarter * 3]

    first_m = q_months[0]
    last_m = q_months[-1]

    # 月度の暦年を決定
    first_y = fiscal_year if first_m >= fy_start_month else fiscal_year + 1
    last_y = fiscal_year if last_m >= fy_start_month else fiscal_year + 1

    start = get_fiscal_month_range(first_y, first_m, closing_day)[0]
    end = get_fiscal_month_range(last_y, last_m, closing_day)[1]

    return (start, end)


def get_fiscal_half_range(fiscal_year: int, half: int,
                          closing_day: int = 15, fy_start_month: int = 10):
    """
    上期/下期の日付範囲を返す。

    half=1 (上期): Q1+Q2 / half=2 (下期): Q3+Q4

    Returns:
        (start_date, end_date)
    """
    if half == 1:
        start = get_fiscal_quarter_range(fiscal_year, 1, closing_day, fy_start_month)[0]
        end = get_fiscal_quarter_range(fiscal_year, 2, closing_day, fy_start_month)[1]
    else:
        start = get_fiscal_quarter_range(fiscal_year, 3, closing_day, fy_start_month)[0]
        end = get_fiscal_quarter_range(fiscal_year, 4, closing_day, fy_start_month)[1]

    return (start, end)


def get_current_fiscal_quarter(d, closing_day: int = 15, fy_start_month: int = 10) -> int:
    """
    日付が属する四半期番号を返す（1〜4）。
    """
    _, fm = get_fiscal_month(d, closing_day)
    ordered = _fiscal_month_order(fy_start_month)
    idx = ordered.index(fm)
    return idx // 3 + 1


def fetch_fiscal_monthly_breakdown(start_date: str, end_date: str, closing_day: int = 15) -> pd.DataFrame:
    """
    期間を月度（15日締め）ごとに分割してデータ取得。
    fiscal_year, fiscal_month, fiscal_label 列を付与。
    """
    from datetime import date as dt_date

    start = dt_date.fromisoformat(start_date)
    end = dt_date.fromisoformat(end_date)

    # 開始日・終了日の月度を求める
    fy_s, fm_s = get_fiscal_month(start, closing_day)
    fy_e, fm_e = get_fiscal_month(end, closing_day)

    all_dfs = []
    fy, fm = fy_s, fm_s

    while (fy, fm) <= (fy_e, fm_e):
        ms, me = get_fiscal_month_range(fy, fm, closing_day)
        # 選択期間でクリップ
        ms = max(ms, start)
        me = min(me, end)

        if ms <= me:
            try:
                raw = fetch_date_range(ms.isoformat(), me.isoformat())
                df = to_dataframe(raw)
                if not df.empty:
                    df["fiscal_year"] = fy
                    df["fiscal_month"] = fm
                    df["fiscal_label"] = f"{fm}月度"
                    all_dfs.append(df)
            except Exception:
                pass

        # 次の月度
        if fm == 12:
            fy += 1
            fm = 1
        else:
            fm += 1

    if not all_dfs:
        return pd.DataFrame()
    return pd.concat(all_dfs, ignore_index=True)


def get_leave_map(data: dict) -> dict:
    """
    GASレスポンスから退勤時間マップを取得。

    Returns:
        {author: {date: "HH:MM", ...}, ...}
    """
    return data.get("leaveMap", {})
