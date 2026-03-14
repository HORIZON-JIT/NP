"""
NPA - GAS APIからデータ取得してDataFrameに変換

GAS Web App（組織内限定）へのアクセスにはGoogle OAuth2認証が必要。
Authorization: Bearer ヘッダーで認証し、リダイレクトを手動で処理する。
"""

import json
import math
import re
import requests
import pandas as pd
from config import GAS_URL, SUB_CATEGORIES, get_main_cd, get_main_label
from gas_auth import get_access_token


# ── 退勤時刻 → 通常/残業 計算（index.html calcFromLeave と同一ロジック） ──

def _get_lunch_break(date_str: str | None = None) -> tuple[int, int]:
    """月から昼休憩時刻(分)を返す。奇数月/偶数月で異なる。"""
    if date_str:
        month = int(date_str[5:7])
    else:
        from datetime import date as dt_date
        month = dt_date.today().month
    if month % 2 == 1:
        return (12 * 60 + 40, 13 * 60 + 30)  # 奇数月: 12:40〜13:30
    else:
        return (12 * 60 + 20, 13 * 60 + 10)  # 偶数月: 12:20〜13:10


def calc_from_leave(leave_hhmm: str, date_str: str | None = None) -> dict | None:
    """
    退勤時刻("HH:MM")から通常/残業時間を計算。

    勤務: 8:30〜退勤
    昼休憩: 奇数月 12:40〜13:30 / 偶数月 12:20〜13:10（50分）
    定時: 17:20（夕方休憩前）、残業: 17:30以降
    丸め: 15分未満切り捨て、15分以上は0.5h単位切り上げ

    Returns:
        {"normal": float, "ot": float} or None
    """
    if not leave_hhmm:
        return None
    parts = leave_hhmm.split(":")
    if len(parts) != 2:
        return None
    try:
        hh, mm = int(parts[0]), int(parts[1])
    except ValueError:
        return None

    leave_mins = hh * 60 + mm
    start_mins = 8 * 60 + 30   # 8:30
    normal_end = 17 * 60 + 20  # 17:20
    ot_start = 17 * 60 + 30    # 17:30

    if leave_mins <= start_mins:
        return {"normal": 0, "ot": 0}

    # 残業計算: 15分未満切り捨て、以降30分ごとに0.5h切り上げ
    ot = 0.0
    if leave_mins > ot_start:
        ot_mins = leave_mins - ot_start
        ot = 0.0 if ot_mins < 15 else math.ceil((ot_mins - 14) / 30) * 0.5

    # 通常時間計算
    work_end = min(leave_mins, normal_end)
    lunch_start, lunch_end = _get_lunch_break(date_str)

    lunch_deduct = 0
    if work_end > lunch_start:
        lunch_deduct = max(0, min(work_end, lunch_end) - lunch_start)

    work_mins = max(0, work_end - start_mins - lunch_deduct)
    normal = 0.0 if work_mins < 15 else math.ceil(work_mins / 30) * 0.5
    normal = max(0.0, normal)

    return {"normal": normal, "ot": ot}


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


def _apply_leave_map(df: pd.DataFrame, leave_map: dict) -> pd.DataFrame:
    """
    退勤ログ(leaveMap)がある担当者は、退勤時刻ベースの通常/残業に置換。
    退勤ログがない担当者・ヘルプはそのまま（日報値フォールバック）。

    置換方法: index.html と同一ロジック
    - 退勤ログから担当者の通常/残業合計を算出
    - 日報の作業コード別比率は維持し、合計値をスケーリング
    """
    if not leave_map or df.empty:
        return df

    df = df.copy()

    for author in df["author"].unique():
        if author == "ヘルプ":
            continue

        author_leaves = leave_map.get(author, {})
        if not author_leaves:
            continue

        # 退勤ログから通常/残業合計を算出
        leave_normal = 0.0
        leave_ot = 0.0
        for date_str, hhmm in author_leaves.items():
            calc = calc_from_leave(hhmm, date_str)
            if calc:
                leave_normal += calc["normal"]
                leave_ot += calc["ot"]

        # 日報の通常/残業合計
        mask = df["author"] == author
        log_normal = df.loc[mask, "hoursNormal"].sum()
        log_ot = df.loc[mask, "hoursOT"].sum()

        # 比率でスケーリング（作業コード別の内訳比率は維持）
        scale_n = (leave_normal / log_normal) if log_normal > 0 else 0
        scale_ot = (leave_ot / log_ot) if log_ot > 0 else 0

        df.loc[mask, "hoursNormal"] = df.loc[mask, "hoursNormal"] * scale_n
        df.loc[mask, "hoursOT"] = df.loc[mask, "hoursOT"] * scale_ot

    # 合計再計算
    df["hoursTotal"] = df["hoursNormal"] + df["hoursOT"]
    return df


def to_dataframe(data: dict) -> pd.DataFrame:
    """
    GAS getDateRange レスポンスをDataFrameに変換。
    退勤ログ(leaveMap)がある場合は退勤時刻ベースの残業時間を優先使用。

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

    # 退勤ログベースの通常/残業に置換（あれば）
    leave_map = data.get("leaveMap", {})
    df = _apply_leave_map(df, leave_map)

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
