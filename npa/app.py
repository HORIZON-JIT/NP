"""
NPA (日報分析) - Streamlit ダッシュボード
ダークモダン UI

起動: streamlit run npa/app.py
"""

import streamlit as st
import plotly.express as px
import plotly.graph_objects as go
import pandas as pd
import numpy as np
from datetime import date, timedelta
import calendar

from config import MAIN_CATEGORIES, AUTHORS, SUB_CATEGORIES, REGULAR_HOURS, get_main_cd, get_main_label
from fetch_data import fetch_date_range, to_dataframe, fetch_weekly_breakdown, fetch_monthly_breakdown

# ── ページ設定 ──
st.set_page_config(page_title="NPA Dashboard", page_icon="📊", layout="wide")

# ══════════════════════════════════════════
# カスタム CSS
# ══════════════════════════════════════════
st.markdown("""
<style>
/* --- グローバル --- */
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');

.stApp {
    font-family: 'Inter', sans-serif;
}

/* --- ヘッダー --- */
.main-header {
    background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
    font-size: 2.2rem;
    font-weight: 700;
    letter-spacing: -0.5px;
    margin-bottom: 0;
    padding-bottom: 0;
}
.sub-header {
    color: #8B8FA3;
    font-size: 0.9rem;
    margin-top: -8px;
    margin-bottom: 24px;
}

/* --- KPI カード --- */
.kpi-card {
    background: linear-gradient(145deg, #1E2030 0%, #171926 100%);
    border: 1px solid rgba(108, 99, 255, 0.15);
    border-radius: 16px;
    padding: 20px 24px;
    text-align: center;
    transition: all 0.3s ease;
    position: relative;
    overflow: hidden;
}
.kpi-card::before {
    content: '';
    position: absolute;
    top: 0;
    left: 0;
    right: 0;
    height: 3px;
    background: linear-gradient(90deg, #6C63FF, #4ECDC4);
}
.kpi-card:hover {
    border-color: rgba(108, 99, 255, 0.4);
    transform: translateY(-2px);
    box-shadow: 0 8px 25px rgba(108, 99, 255, 0.15);
}
.kpi-label {
    color: #8B8FA3;
    font-size: 0.78rem;
    font-weight: 500;
    text-transform: uppercase;
    letter-spacing: 1px;
    margin-bottom: 8px;
}
.kpi-value {
    font-size: 1.8rem;
    font-weight: 700;
    color: #E8E8F0;
    line-height: 1.2;
}
.kpi-delta {
    font-size: 0.8rem;
    margin-top: 4px;
    font-weight: 500;
}
.kpi-delta.warn { color: #F59E0B; }
.kpi-delta.good { color: #4ECDC4; }
.kpi-delta.bad  { color: #EF4444; }

/* アクセントカラーバリエーション */
.kpi-card.purple::before { background: linear-gradient(90deg, #6C63FF, #A78BFA); }
.kpi-card.blue::before   { background: linear-gradient(90deg, #3B82F6, #60A5FA); }
.kpi-card.red::before    { background: linear-gradient(90deg, #EF4444, #F87171); }
.kpi-card.teal::before   { background: linear-gradient(90deg, #14B8A6, #4ECDC4); }
.kpi-card.amber::before  { background: linear-gradient(90deg, #F59E0B, #FBBF24); }

/* --- セクションカード --- */
.section-card {
    background: linear-gradient(145deg, #1A1D2E 0%, #151725 100%);
    border: 1px solid rgba(255,255,255,0.06);
    border-radius: 16px;
    padding: 28px;
    margin-bottom: 20px;
}
.section-title {
    color: #C8CDDF;
    font-size: 1.15rem;
    font-weight: 600;
    margin-bottom: 20px;
    display: flex;
    align-items: center;
    gap: 10px;
}
.section-title .icon {
    background: linear-gradient(135deg, #6C63FF 0%, #4ECDC4 100%);
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
    font-size: 1.3rem;
}

/* --- サイドバー ナビゲーション --- */
section[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #0D0F18 0%, #131520 100%);
    border-right: 1px solid rgba(108, 99, 255, 0.1);
}
section[data-testid="stSidebar"] .stRadio > label {
    font-size: 0.85rem;
    font-weight: 600;
    color: #8B8FA3;
    text-transform: uppercase;
    letter-spacing: 0.5px;
}

/* --- テーブル --- */
.stDataFrame {
    border-radius: 12px;
    overflow: hidden;
}

/* --- 区切り線 --- */
hr {
    border-color: rgba(108, 99, 255, 0.15) !important;
    margin: 28px 0 !important;
}

/* --- アラートカード --- */
.alert-card {
    border-radius: 12px;
    padding: 16px 20px;
    margin-bottom: 12px;
}
.alert-card.danger {
    background: rgba(239, 68, 68, 0.1);
    border-left: 4px solid #EF4444;
}
.alert-card.warning {
    background: rgba(245, 158, 11, 0.1);
    border-left: 4px solid #F59E0B;
}
.alert-card.success {
    background: rgba(34, 197, 94, 0.1);
    border-left: 4px solid #22C55E;
}

/* --- プログレスバー --- */
.stProgress > div > div > div {
    background: linear-gradient(90deg, #6C63FF, #4ECDC4);
}

/* --- Streamlit デフォルト非表示 --- */
#MainMenu {visibility: hidden;}
footer {visibility: hidden;}
header {visibility: hidden;}
</style>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════
# Plotly 共通テンプレート
# ══════════════════════════════════════════
PLOTLY_LAYOUT = dict(
    template="plotly_dark",
    paper_bgcolor="rgba(0,0,0,0)",
    plot_bgcolor="rgba(0,0,0,0)",
    font=dict(family="Inter, sans-serif", color="#C8CDDF"),
    title_font=dict(size=16, color="#E8E8F0"),
    legend=dict(
        bgcolor="rgba(0,0,0,0)",
        font=dict(size=11, color="#8B8FA3"),
    ),
    margin=dict(l=40, r=20, t=50, b=40),
    xaxis=dict(gridcolor="rgba(108,99,255,0.08)", zerolinecolor="rgba(108,99,255,0.12)"),
    yaxis=dict(gridcolor="rgba(108,99,255,0.08)", zerolinecolor="rgba(108,99,255,0.12)"),
)

# ダーク用カラーパレット（ネオン系）
NEON_COLORS = ["#6C63FF", "#4ECDC4", "#FF6B9D", "#C084FC", "#FBBF24",
               "#34D399", "#F87171", "#60A5FA", "#A78BFA"]


def apply_dark(fig):
    """Plotly図にダークテーマを適用"""
    fig.update_layout(**PLOTLY_LAYOUT)
    return fig


# ══════════════════════════════════════════
# サイドバー
# ══════════════════════════════════════════
with st.sidebar:
    st.markdown('<p class="main-header" style="font-size:1.5rem;">NPA</p>', unsafe_allow_html=True)
    st.markdown('<p style="color:#8B8FA3;font-size:0.75rem;margin-top:-12px;">Daily Report Analytics</p>', unsafe_allow_html=True)
    st.markdown("---")

    # ナビゲーション
    page = st.radio(
        "NAVIGATION",
        [
            "🏠 ダッシュボード",
            "👤 担当者別",
            "📂 工程別",
            "⏰ 残業・36協定",
            "📈 推移分析",
            "📊 稼働・偏り",
            "📅 月別・前年比",
            "🔍 個人サマリ",
            "📋 データ一覧",
        ],
        label_visibility="collapsed",
    )

    st.markdown("---")

    # 期間設定
    st.markdown("**📅 分析期間**")
    today = date.today()
    default_start = today.replace(day=1)
    default_end = today

    col_s, col_e = st.columns(2)
    start_date = col_s.date_input("開始", value=default_start, label_visibility="collapsed")
    end_date = col_e.date_input("終了", value=default_end, label_visibility="collapsed")

    if start_date > end_date:
        st.error("開始日 > 終了日")
        st.stop()

    st.markdown("---")

    # 担当者フィルタ
    st.markdown("**👥 担当者フィルタ**")
    selected_authors = st.multiselect(
        "担当者（空=全員）", options=AUTHORS, default=[], label_visibility="collapsed"
    )

# ── 共通ユーティリティ ──
author_order = {a: i for i, a in enumerate(AUTHORS)}
cat_colors = {v["label"]: v["color"] for k, v in MAIN_CATEGORIES.items()}
# ダーク向けカラーマップ（明るめ調整）
cat_colors_dark = {
    "計画": "#60A5FA", "手配": "#34D399", "受注": "#F87171",
    "BOMメンテ": "#C084FC", "打合せ": "#FBBF24", "改善": "#22D3EE",
    "その他": "#9CA3AF", "アフター": "#F472B6",
}


def sort_by_author(frame: pd.DataFrame) -> pd.DataFrame:
    frame["_o"] = frame["author"].map(lambda a: author_order.get(a, 99))
    return frame.sort_values("_o").drop(columns="_o")


def business_days(start: date, end: date) -> int:
    return int(np.busday_count(start, end + timedelta(days=1)))


def kpi_card(label: str, value: str, delta: str = "", delta_class: str = "good", accent: str = "purple") -> str:
    delta_html = f'<div class="kpi-delta {delta_class}">{delta}</div>' if delta else ""
    return f"""
    <div class="kpi-card {accent}">
        <div class="kpi-label">{label}</div>
        <div class="kpi-value">{value}</div>
        {delta_html}
    </div>
    """


def section_header(icon: str, title: str):
    st.markdown(f'<div class="section-title"><span class="icon">{icon}</span> {title}</div>', unsafe_allow_html=True)


# ── データ取得 ──
@st.cache_data(ttl=300, show_spinner=False)
def load_data(s: str, e: str):
    raw = fetch_date_range(s, e)
    return to_dataframe(raw), raw


@st.cache_data(ttl=600, show_spinner=False)
def load_weekly(s: str, e: str):
    return fetch_weekly_breakdown(s, e)


@st.cache_data(ttl=600, show_spinner=False)
def load_monthly(s: str, e: str):
    return fetch_monthly_breakdown(s, e)


with st.spinner("データ取得中..."):
    try:
        df, raw_data = load_data(str(start_date), str(end_date))
    except FileNotFoundError as e:
        st.error("Google OAuth2 認証が未設定です。")
        st.info(str(e))
        st.stop()
    except RuntimeError as e:
        err_msg = str(e)
        if "認証" in err_msg or "ログイン" in err_msg:
            st.error(f"認証エラー: {err_msg}")
            st.info("ターミナルで `python npa/gas_auth.py` を実行して再認証してください。")
        else:
            st.error(f"データ取得エラー: {err_msg}")
        st.stop()
    except Exception as e:
        st.error(f"データ取得エラー: {e}")
        st.stop()

if df.empty:
    st.warning("この期間のデータがありません。")
    st.stop()

# 担当者フィルタ適用
if selected_authors:
    df = df[df["author"].isin(selected_authors)].copy()
    if df.empty:
        st.warning("選択した担当者のデータがありません。")
        st.stop()

# 共通集計
total_normal = df["hoursNormal"].sum()
total_ot = df["hoursOT"].sum()
total_all = total_normal + total_ot
n_authors = df["author"].nunique()
bdays = business_days(start_date, end_date)
avg_per_day = total_all / bdays if bdays else 0

# ══════════════════════════════════════════
# ページルーティング
# ══════════════════════════════════════════

# ────────────────────────────────────
# ダッシュボード（ホーム）
# ────────────────────────────────────
if page == "🏠 ダッシュボード":
    st.markdown('<p class="main-header">Daily Report Dashboard</p>', unsafe_allow_html=True)
    st.markdown(f'<p class="sub-header">{start_date.strftime("%Y/%m/%d")} — {end_date.strftime("%Y/%m/%d")}　|　営業日 {bdays}日　|　{n_authors}名</p>', unsafe_allow_html=True)

    # KPI カード
    ot_pct = f"{total_ot/total_all*100:.0f}%" if total_all else "0%"
    avg_val = f"{avg_per_day/n_authors:.1f}h" if n_authors else "-"

    k1, k2, k3, k4, k5 = st.columns(5)
    with k1:
        st.markdown(kpi_card("総工数", f"{total_all:.1f}h", accent="purple"), unsafe_allow_html=True)
    with k2:
        st.markdown(kpi_card("通常時間", f"{total_normal:.1f}h", accent="blue"), unsafe_allow_html=True)
    with k3:
        st.markdown(kpi_card("残業時間", f"{total_ot:.1f}h", f"全体の {ot_pct}",
                             "warn" if total_ot/total_all*100 > 15 else "good" if total_all else "good",
                             "red"), unsafe_allow_html=True)
    with k4:
        st.markdown(kpi_card("担当者数", f"{n_authors}名", accent="teal"), unsafe_allow_html=True)
    with k5:
        st.markdown(kpi_card("1日平均/人", avg_val, accent="amber"), unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # 概要チャート2列
    col_left, col_right = st.columns(2)

    with col_left:
        # 担当者別工数（積み上げ棒）
        auth_df = (
            df.groupby("author", as_index=False)
            .agg(通常=("hoursNormal", "sum"), 残業=("hoursOT", "sum"))
        )
        auth_df["合計"] = auth_df["通常"] + auth_df["残業"]
        auth_df = sort_by_author(auth_df)

        fig = go.Figure()
        fig.add_trace(go.Bar(x=auth_df["author"], y=auth_df["通常"],
                             name="通常", marker_color="#60A5FA"))
        fig.add_trace(go.Bar(x=auth_df["author"], y=auth_df["残業"],
                             name="残業", marker_color="#F87171"))
        fig.update_layout(barmode="stack", title="担当者別工数",
                          yaxis_title="時間 (h)", height=400)
        apply_dark(fig)
        st.plotly_chart(fig, use_container_width=True)

    with col_right:
        # 大分類別 工数比率（ドーナツ）
        cat_df = (
            df.groupby(["mainCd", "mainLabel"], as_index=False)
            .agg(hours=("hoursTotal", "sum"))
        )
        cat_df = cat_df[cat_df["hours"] > 0]

        fig3 = px.pie(cat_df, values="hours", names="mainLabel",
                      color="mainLabel", color_discrete_map=cat_colors_dark,
                      title="大分類別 工数比率", hole=0.45)
        fig3.update_traces(textinfo="label+percent", textfont_size=11)
        apply_dark(fig3)
        st.plotly_chart(fig3, use_container_width=True)

    # 担当者×大分類 積み上げ
    cross = (
        df.groupby(["author", "mainLabel"], as_index=False)
        .agg(hours=("hoursTotal", "sum"))
    )
    fig2 = px.bar(cross, x="author", y="hours", color="mainLabel",
                  color_discrete_map=cat_colors_dark,
                  title="担当者 × 大分類 工数内訳",
                  labels={"hours": "時間 (h)", "author": "担当者", "mainLabel": "大分類"})
    fig2.update_layout(barmode="stack", height=420)
    apply_dark(fig2)
    st.plotly_chart(fig2, use_container_width=True)


# ────────────────────────────────────
# 担当者別
# ────────────────────────────────────
elif page == "👤 担当者別":
    st.markdown('<p class="main-header">担当者別分析</p>', unsafe_allow_html=True)
    st.markdown(f'<p class="sub-header">{start_date.strftime("%Y/%m/%d")} — {end_date.strftime("%Y/%m/%d")}</p>', unsafe_allow_html=True)

    auth_df = (
        df.groupby("author", as_index=False)
        .agg(通常=("hoursNormal", "sum"), 残業=("hoursOT", "sum"))
    )
    auth_df["合計"] = auth_df["通常"] + auth_df["残業"]
    auth_df = sort_by_author(auth_df)

    fig = go.Figure()
    fig.add_trace(go.Bar(x=auth_df["author"], y=auth_df["通常"],
                         name="通常", marker_color="#60A5FA"))
    fig.add_trace(go.Bar(x=auth_df["author"], y=auth_df["残業"],
                         name="残業", marker_color="#F87171"))
    fig.update_layout(barmode="stack", title="担当者別工数",
                      yaxis_title="時間 (h)", height=450)
    apply_dark(fig)
    st.plotly_chart(fig, use_container_width=True)

    cross = (
        df.groupby(["author", "mainLabel"], as_index=False)
        .agg(hours=("hoursTotal", "sum"))
    )
    fig2 = px.bar(cross, x="author", y="hours", color="mainLabel",
                  color_discrete_map=cat_colors_dark,
                  title="担当者×大分類 工数内訳",
                  labels={"hours": "時間 (h)", "author": "担当者", "mainLabel": "大分類"})
    fig2.update_layout(barmode="stack", height=450)
    apply_dark(fig2)
    st.plotly_chart(fig2, use_container_width=True)

# ────────────────────────────────────
# 工程別
# ────────────────────────────────────
elif page == "📂 工程別":
    st.markdown('<p class="main-header">工程別分析</p>', unsafe_allow_html=True)
    st.markdown(f'<p class="sub-header">{start_date.strftime("%Y/%m/%d")} — {end_date.strftime("%Y/%m/%d")}</p>', unsafe_allow_html=True)

    col_a, col_b = st.columns(2)

    cat_df = (
        df.groupby(["mainCd", "mainLabel"], as_index=False)
        .agg(hours=("hoursTotal", "sum"))
    )
    cat_df = cat_df[cat_df["hours"] > 0]

    with col_a:
        fig3 = px.pie(cat_df, values="hours", names="mainLabel",
                      color="mainLabel", color_discrete_map=cat_colors_dark,
                      title="大分類別 工数比率", hole=0.4)
        fig3.update_traces(textinfo="label+percent+value")
        apply_dark(fig3)
        st.plotly_chart(fig3, use_container_width=True)

    with col_b:
        cat_order = list(MAIN_CATEGORIES.keys())
        cat_df["_o"] = cat_df["mainCd"].map(lambda c: cat_order.index(c) if c in cat_order else 99)
        cat_df = cat_df.sort_values("_o")
        fig4 = px.bar(cat_df, x="mainLabel", y="hours",
                      color="mainLabel", color_discrete_map=cat_colors_dark,
                      title="大分類別 工数",
                      labels={"hours": "時間 (h)", "mainLabel": "大分類"})
        fig4.update_layout(showlegend=False)
        apply_dark(fig4)
        st.plotly_chart(fig4, use_container_width=True)

    sub_df = (
        df.groupby(["cdSub", "subLabel", "mainLabel"], as_index=False)
        .agg(通常=("hoursNormal", "sum"), 残業=("hoursOT", "sum"))
    )
    sub_df["合計"] = sub_df["通常"] + sub_df["残業"]
    sub_df = sub_df.sort_values("cdSub")
    sub_df = sub_df.rename(columns={"cdSub": "コード", "subLabel": "作業内容", "mainLabel": "大分類"})

    section_header("📋", "中分類別 工数一覧")
    st.dataframe(sub_df, use_container_width=True, hide_index=True)

# ────────────────────────────────────
# 残業・36協定
# ────────────────────────────────────
elif page == "⏰ 残業・36協定":
    st.markdown('<p class="main-header">残業・36協定管理</p>', unsafe_allow_html=True)
    st.markdown(f'<p class="sub-header">{start_date.strftime("%Y/%m/%d")} — {end_date.strftime("%Y/%m/%d")}</p>', unsafe_allow_html=True)

    ot_df = (
        df.groupby("author", as_index=False)
        .agg(残業=("hoursOT", "sum"))
    )
    ot_df = sort_by_author(ot_df)

    fig5 = go.Figure()
    fig5.add_trace(go.Bar(
        x=ot_df["author"], y=ot_df["残業"],
        marker=dict(
            color=ot_df["残業"],
            colorscale=[[0, "#4ECDC4"], [0.5, "#FBBF24"], [1, "#EF4444"]],
        ),
        text=[f"{v:.1f}h" for v in ot_df["残業"]],
        textposition="outside",
        textfont=dict(color="#C8CDDF"),
    ))
    fig5.update_layout(title="担当者別 残業時間", yaxis_title="残業時間 (h)", height=400)
    fig5.add_hline(y=45, line_dash="dash", line_color="#EF4444",
                   annotation_text="月上限 45h", annotation_font_color="#EF4444")
    apply_dark(fig5)
    st.plotly_chart(fig5, use_container_width=True)

    # 残業比率テーブル
    auth_full = (
        df.groupby("author", as_index=False)
        .agg(通常=("hoursNormal", "sum"), 残業=("hoursOT", "sum"))
    )
    auth_full["合計"] = auth_full["通常"] + auth_full["残業"]
    auth_full["残業率"] = (auth_full["残業"] / auth_full["合計"] * 100).round(1)
    auth_full = sort_by_author(auth_full)

    section_header("📊", "担当者別 残業比率")
    st.dataframe(
        auth_full[["author", "通常", "残業", "合計", "残業率"]].rename(
            columns={"author": "担当者", "残業率": "残業率 (%)"}
        ),
        use_container_width=True, hide_index=True
    )

    # ── 36協定アラート ──
    st.markdown("---")
    section_header("⚡", "36協定 残業上限チェック")
    st.caption("月45h / 年360h が上限（特別条項: 月100h / 年720h）")

    ot_by_author = (
        df.groupby("author", as_index=False)
        .agg(月残業=("hoursOT", "sum"))
    )
    ot_by_author = sort_by_author(ot_by_author)

    year_start = date(end_date.year, 1, 1)
    with st.spinner("年間残業データ取得中..."):
        try:
            year_monthly = load_monthly(year_start.isoformat(), str(end_date))
            if not year_monthly.empty:
                year_ot = (
                    year_monthly.groupby("author", as_index=False)
                    .agg(年間累計=("hoursOT", "sum"))
                )
            else:
                year_ot = pd.DataFrame(columns=["author", "年間累計"])
        except Exception:
            year_ot = pd.DataFrame(columns=["author", "年間累計"])

    limit_df = ot_by_author.merge(year_ot, on="author", how="left").fillna(0)

    for _, row in limit_df.iterrows():
        name = row["author"]
        monthly_ot = row["月残業"]
        yearly_ot = row["年間累計"]

        col_name, col_month, col_year = st.columns([1, 2, 2])
        with col_name:
            st.markdown(f"**{name}**")
        with col_month:
            pct_m = min(monthly_ot / 45 * 100, 100)
            st.markdown(f"月: **{monthly_ot:.1f}h** / 45h")
            st.progress(min(pct_m / 100, 1.0))
            if monthly_ot > 45:
                st.markdown(f'<div class="alert-card danger">月上限超過 (+{monthly_ot - 45:.1f}h)</div>', unsafe_allow_html=True)
            elif monthly_ot > 36:
                st.markdown(f'<div class="alert-card warning">月上限まで残り {45 - monthly_ot:.1f}h</div>', unsafe_allow_html=True)
        with col_year:
            pct_y = min(yearly_ot / 360 * 100, 100)
            st.markdown(f"年: **{yearly_ot:.1f}h** / 360h")
            st.progress(min(pct_y / 100, 1.0))
            if yearly_ot > 360:
                st.markdown(f'<div class="alert-card danger">年上限超過 (+{yearly_ot - 360:.1f}h)</div>', unsafe_allow_html=True)
            elif yearly_ot > 300:
                st.markdown(f'<div class="alert-card warning">年上限まで残り {360 - yearly_ot:.1f}h</div>', unsafe_allow_html=True)

# ────────────────────────────────────
# 推移分析
# ────────────────────────────────────
elif page == "📈 推移分析":
    st.markdown('<p class="main-header">推移分析</p>', unsafe_allow_html=True)
    st.markdown(f'<p class="sub-header">{start_date.strftime("%Y/%m/%d")} — {end_date.strftime("%Y/%m/%d")}</p>', unsafe_allow_html=True)

    with st.spinner("週別データ取得中..."):
        try:
            weekly_df = load_weekly(str(start_date), str(end_date))
        except Exception as e:
            st.error(f"週別データ取得エラー: {e}")
            weekly_df = pd.DataFrame()

    if not weekly_df.empty:
        if selected_authors:
            weekly_df = weekly_df[weekly_df["author"].isin(selected_authors)].copy()

        # 週別×担当者 残業推移
        wk_ot = (
            weekly_df.groupby(["week_start", "author"], as_index=False)
            .agg(残業=("hoursOT", "sum"), 通常=("hoursNormal", "sum"))
        )
        wk_ot["合計"] = wk_ot["通常"] + wk_ot["残業"]

        fig_trend_ot = px.line(
            wk_ot, x="week_start", y="残業", color="author",
            title="週別 残業推移（担当者別）",
            labels={"week_start": "週", "残業": "残業時間 (h)", "author": "担当者"},
            markers=True, color_discrete_sequence=NEON_COLORS
        )
        fig_trend_ot.update_layout(height=400)
        apply_dark(fig_trend_ot)
        st.plotly_chart(fig_trend_ot, use_container_width=True)

        # 週別 合計工数推移
        wk_total = (
            weekly_df.groupby("week_start", as_index=False)
            .agg(通常=("hoursNormal", "sum"), 残業=("hoursOT", "sum"))
        )
        wk_total["合計"] = wk_total["通常"] + wk_total["残業"]

        fig_trend_total = go.Figure()
        fig_trend_total.add_trace(go.Bar(
            x=wk_total["week_start"], y=wk_total["通常"],
            name="通常", marker_color="#60A5FA"
        ))
        fig_trend_total.add_trace(go.Bar(
            x=wk_total["week_start"], y=wk_total["残業"],
            name="残業", marker_color="#F87171"
        ))
        fig_trend_total.update_layout(
            barmode="stack", title="週別 全体工数推移",
            xaxis_title="週", yaxis_title="時間 (h)", height=400
        )
        apply_dark(fig_trend_total)
        st.plotly_chart(fig_trend_total, use_container_width=True)

        # 週別×大分類 推移
        wk_cat = (
            weekly_df.groupby(["week_start", "mainLabel"], as_index=False)
            .agg(hours=("hoursTotal", "sum"))
        )
        fig_trend_cat = px.bar(
            wk_cat, x="week_start", y="hours", color="mainLabel",
            color_discrete_map=cat_colors_dark,
            title="週別 大分類別工数推移",
            labels={"week_start": "週", "hours": "時間 (h)", "mainLabel": "大分類"}
        )
        fig_trend_cat.update_layout(barmode="stack", height=400)
        apply_dark(fig_trend_cat)
        st.plotly_chart(fig_trend_cat, use_container_width=True)
    else:
        st.info("週別データがありません。")

# ────────────────────────────────────
# 稼働・偏り
# ────────────────────────────────────
elif page == "📊 稼働・偏り":
    st.markdown('<p class="main-header">稼働率・偏り分析</p>', unsafe_allow_html=True)
    st.markdown(f'<p class="sub-header">{start_date.strftime("%Y/%m/%d")} — {end_date.strftime("%Y/%m/%d")}　|　所定: {bdays}日 × {REGULAR_HOURS}h = {bdays * REGULAR_HOURS:.0f}h/人</p>', unsafe_allow_html=True)

    expected_hours = bdays * REGULAR_HOURS

    util_df = (
        df.groupby("author", as_index=False)
        .agg(通常=("hoursNormal", "sum"), 残業=("hoursOT", "sum"))
    )
    util_df["合計"] = util_df["通常"] + util_df["残業"]
    util_df["所定"] = expected_hours
    util_df["稼働率"] = (util_df["合計"] / expected_hours * 100).round(1) if expected_hours else 0
    util_df = sort_by_author(util_df)

    # 稼働率ゲージ風棒グラフ
    fig_util = go.Figure()
    fig_util.add_trace(go.Bar(
        x=util_df["author"], y=util_df["稼働率"],
        marker_color=[
            "#EF4444" if r > 120 else "#FBBF24" if r > 105 else "#4ECDC4"
            for r in util_df["稼働率"]
        ],
        text=[f"{r:.0f}%" for r in util_df["稼働率"]],
        textposition="outside",
        textfont=dict(color="#C8CDDF"),
    ))
    fig_util.add_hline(y=100, line_dash="dash", line_color="#8B8FA3",
                       annotation_text="所定100%", annotation_font_color="#8B8FA3")
    fig_util.update_layout(title="担当者別 稼働率", yaxis_title="稼働率 (%)", height=400)
    apply_dark(fig_util)
    st.plotly_chart(fig_util, use_container_width=True)

    st.dataframe(
        util_df[["author", "通常", "残業", "合計", "所定", "稼働率"]].rename(
            columns={"author": "担当者", "稼働率": "稼働率 (%)"}
        ),
        use_container_width=True, hide_index=True
    )

    st.markdown("---")

    # ── ヒートマップ ──
    heat_df = (
        df.groupby(["author", "mainLabel"], as_index=False)
        .agg(hours=("hoursTotal", "sum"))
    )
    heat_pivot = heat_df.pivot_table(
        index="author", columns="mainLabel", values="hours", fill_value=0
    )
    ordered_authors = [a for a in AUTHORS if a in heat_pivot.index]
    heat_pivot = heat_pivot.reindex(ordered_authors)
    ordered_cats = [v["label"] for v in MAIN_CATEGORIES.values() if v["label"] in heat_pivot.columns]
    heat_pivot = heat_pivot[ordered_cats]

    section_header("🔥", "担当者 × 大分類 ヒートマップ")

    heat_mode = st.radio(
        "表示モード", ["実時間 (h)", "構成比 (%)", "チーム平均との差分"],
        horizontal=True, key="heat_mode"
    )

    if heat_mode == "実時間 (h)":
        fig_heat = px.imshow(
            heat_pivot, text_auto=".1f",
            color_continuous_scale="Plasma",
            labels={"x": "大分類", "y": "担当者", "color": "時間 (h)"},
            title="工数ヒートマップ（実時間）"
        )
    elif heat_mode == "構成比 (%)":
        heat_pct = heat_pivot.div(heat_pivot.sum(axis=1), axis=0) * 100
        fig_heat = px.imshow(
            heat_pct.round(1), text_auto=".1f",
            color_continuous_scale="Viridis",
            labels={"x": "大分類", "y": "担当者", "color": "比率 (%)"},
            title="工数ヒートマップ（構成比）"
        )
    else:
        team_avg = heat_pivot.mean(axis=0)
        heat_diff = heat_pivot.sub(team_avg, axis=1)
        fig_heat = px.imshow(
            heat_diff.round(1), text_auto="+.1f",
            color_continuous_scale="RdBu_r", color_continuous_midpoint=0,
            labels={"x": "大分類", "y": "担当者", "color": "差分 (h)"},
            title="チーム平均との差分（赤=平均以上 / 青=平均以下）"
        )

    fig_heat.update_layout(height=max(350, len(ordered_authors) * 50))
    apply_dark(fig_heat)
    st.plotly_chart(fig_heat, use_container_width=True)

    st.markdown("---")

    # ── レーダーチャート ──
    section_header("🎯", "作業バランス レーダーチャート")

    radar_authors = st.multiselect(
        "比較する担当者（空=全員）",
        options=ordered_authors, default=[],
        key="radar_authors"
    )
    radar_targets = radar_authors if radar_authors else ordered_authors

    heat_pct_radar = heat_pivot.div(heat_pivot.sum(axis=1), axis=0) * 100
    heat_pct_radar = heat_pct_radar.fillna(0)

    fig_radar = go.Figure()
    for i, person in enumerate(radar_targets):
        if person in heat_pct_radar.index:
            vals = heat_pct_radar.loc[person].tolist()
            vals.append(vals[0])
            cats_loop = ordered_cats + [ordered_cats[0]]
            fig_radar.add_trace(go.Scatterpolar(
                r=vals, theta=cats_loop, name=person,
                fill="toself", opacity=0.3,
                line=dict(color=NEON_COLORS[i % len(NEON_COLORS)])
            ))

    team_avg_pct = heat_pct_radar.mean(axis=0).tolist()
    team_avg_pct.append(team_avg_pct[0])
    fig_radar.add_trace(go.Scatterpolar(
        r=team_avg_pct,
        theta=ordered_cats + [ordered_cats[0]],
        name="チーム平均",
        line=dict(dash="dash", color="#E8E8F0", width=2),
        opacity=0.8
    ))

    fig_radar.update_layout(
        polar=dict(
            radialaxis=dict(visible=True, range=[0, 100], gridcolor="rgba(108,99,255,0.15)"),
            angularaxis=dict(gridcolor="rgba(108,99,255,0.15)"),
            bgcolor="rgba(0,0,0,0)",
        ),
        title="作業構成比レーダー（チーム平均は破線）",
        height=500
    )
    apply_dark(fig_radar)
    st.plotly_chart(fig_radar, use_container_width=True)

    st.markdown("---")

    # ── 中分類ヒートマップ ──
    section_header("🔍", "中分類ヒートマップ（詳細）")

    sub_heat = (
        df.groupby(["author", "subLabel"], as_index=False)
        .agg(hours=("hoursTotal", "sum"))
    )
    sub_pivot = sub_heat.pivot_table(
        index="author", columns="subLabel", values="hours", fill_value=0
    )
    sub_pivot = sub_pivot.reindex(ordered_authors)
    col_order = sub_pivot.sum().sort_values(ascending=False).index.tolist()
    sub_pivot = sub_pivot[col_order]

    fig_sub_heat = px.imshow(
        sub_pivot, text_auto=".1f",
        color_continuous_scale="Cividis",
        labels={"x": "中分類", "y": "担当者", "color": "時間 (h)"},
        title="担当者 × 中分類 ヒートマップ"
    )
    fig_sub_heat.update_layout(
        height=max(400, len(ordered_authors) * 50),
        xaxis=dict(tickangle=-45)
    )
    apply_dark(fig_sub_heat)
    st.plotly_chart(fig_sub_heat, use_container_width=True)

    st.markdown("---")

    # ── 作業集中度 ──
    section_header("📡", "作業集中度")
    st.caption("一つの大分類に工数が集中している度合いを分析")

    conc_df = heat_df.copy()
    author_totals = conc_df.groupby("author")["hours"].transform("sum")
    conc_df["比率"] = (conc_df["hours"] / author_totals * 100).round(1)

    hhi_df = conc_df.groupby("author").apply(
        lambda g: (g["比率"] ** 2).sum(), include_groups=False
    ).reset_index(name="HHI")
    hhi_df = sort_by_author(hhi_df)
    hhi_df["集中レベル"] = hhi_df["HHI"].map(
        lambda h: "高い" if h > 5000 else ("やや高い" if h > 3000 else "分散")
    )

    col_hhi_chart, col_hhi_table = st.columns([2, 1])

    with col_hhi_chart:
        fig_hhi = go.Figure()
        fig_hhi.add_trace(go.Bar(
            x=hhi_df["author"], y=hhi_df["HHI"],
            marker_color=[
                "#EF4444" if h > 5000 else "#FBBF24" if h > 3000 else "#4ECDC4"
                for h in hhi_df["HHI"]
            ],
            text=hhi_df["集中レベル"],
            textposition="outside",
            textfont=dict(color="#C8CDDF"),
        ))
        fig_hhi.add_hline(y=3000, line_dash="dash", line_color="#FBBF24",
                          annotation_text="集中度しきい値", annotation_font_color="#FBBF24")
        fig_hhi.update_layout(title="作業集中度 (HHI指数)", yaxis_title="HHI", height=400)
        apply_dark(fig_hhi)
        st.plotly_chart(fig_hhi, use_container_width=True)

    with col_hhi_table:
        st.markdown("**HHI指数の見方**")
        st.markdown("""
        - **~1,250**: 全分類に均等配分
        - **3,000超**: やや集中傾向
        - **5,000超**: 特定作業に偏り
        - **10,000**: 1つの作業のみ
        """)

        high_conc = conc_df[conc_df["比率"] >= 50].sort_values("比率", ascending=False)
        if not high_conc.empty:
            st.markdown("**50%超の集中:**")
            st.dataframe(
                high_conc[["author", "mainLabel", "hours", "比率"]].rename(
                    columns={"author": "担当者", "mainLabel": "大分類",
                             "hours": "時間 (h)", "比率": "比率 (%)"}
                ),
                use_container_width=True, hide_index=True
            )
        else:
            st.markdown('<div class="alert-card success">50%超の集中なし</div>', unsafe_allow_html=True)

    st.markdown("---")

    # ── 担当者間の類似度 ──
    section_header("🔗", "担当者間 作業パターン類似度")
    st.caption("作業構成比のコサイン類似度。1.0に近いほど似た作業パターン")

    heat_pct_sim = heat_pivot.div(heat_pivot.sum(axis=1), axis=0).fillna(0)
    from numpy.linalg import norm
    sim_data = []
    authors_list = heat_pct_sim.index.tolist()
    for i, a1 in enumerate(authors_list):
        row = []
        v1 = heat_pct_sim.loc[a1].values
        for j, a2 in enumerate(authors_list):
            v2 = heat_pct_sim.loc[a2].values
            n1, n2 = norm(v1), norm(v2)
            cos_sim = float(np.dot(v1, v2) / (n1 * n2)) if n1 and n2 else 0
            row.append(round(cos_sim, 2))
        sim_data.append(row)

    sim_df = pd.DataFrame(sim_data, index=authors_list, columns=authors_list)

    fig_sim = px.imshow(
        sim_df, text_auto=".2f",
        color_continuous_scale="Magma",
        zmin=0, zmax=1,
        labels={"x": "担当者", "y": "担当者", "color": "類似度"},
        title="作業パターン類似度マトリクス"
    )
    fig_sim.update_layout(height=max(400, len(authors_list) * 55))
    apply_dark(fig_sim)
    st.plotly_chart(fig_sim, use_container_width=True)

# ────────────────────────────────────
# 月別・前年比
# ────────────────────────────────────
elif page == "📅 月別・前年比":
    st.markdown('<p class="main-header">月別・前年比</p>', unsafe_allow_html=True)
    st.markdown(f'<p class="sub-header">{start_date.strftime("%Y/%m/%d")} — {end_date.strftime("%Y/%m/%d")}</p>', unsafe_allow_html=True)

    year_start_str = date(end_date.year, 1, 1).isoformat()
    with st.spinner("月別データ取得中..."):
        try:
            monthly_df = load_monthly(year_start_str, str(end_date))
        except Exception as e:
            st.error(f"月別データ取得エラー: {e}")
            monthly_df = pd.DataFrame()

    if not monthly_df.empty:
        if selected_authors:
            monthly_df = monthly_df[monthly_df["author"].isin(selected_authors)].copy()

        m_agg = (
            monthly_df.groupby("month", as_index=False)
            .agg(通常=("hoursNormal", "sum"), 残業=("hoursOT", "sum"))
        )
        m_agg["月名"] = m_agg["month"].map(lambda m: f"{m}月")

        fig_month = go.Figure()
        fig_month.add_trace(go.Bar(x=m_agg["月名"], y=m_agg["通常"],
                                   name="通常", marker_color="#60A5FA"))
        fig_month.add_trace(go.Bar(x=m_agg["月名"], y=m_agg["残業"],
                                   name="残業", marker_color="#F87171"))
        fig_month.update_layout(barmode="stack", title=f"{end_date.year}年 月別工数推移",
                                yaxis_title="時間 (h)", height=400)
        apply_dark(fig_month)
        st.plotly_chart(fig_month, use_container_width=True)

        m_ot = (
            monthly_df.groupby(["month", "author"], as_index=False)
            .agg(残業=("hoursOT", "sum"))
        )
        m_ot["月名"] = m_ot["month"].map(lambda m: f"{m}月")

        fig_m_ot = px.line(
            m_ot, x="月名", y="残業", color="author",
            title=f"{end_date.year}年 月別残業推移（担当者別）",
            labels={"月名": "月", "残業": "残業時間 (h)", "author": "担当者"},
            markers=True, color_discrete_sequence=NEON_COLORS
        )
        fig_m_ot.update_layout(height=400)
        apply_dark(fig_m_ot)
        st.plotly_chart(fig_m_ot, use_container_width=True)
    else:
        st.info("月別データがありません。")

    # ── 前年比 ──
    st.markdown("---")
    section_header("📊", "前年比")

    prev_year = end_date.year - 1
    prev_start = date(prev_year, start_date.month, start_date.day)
    prev_end = date(prev_year, end_date.month, min(end_date.day,
                    calendar.monthrange(prev_year, end_date.month)[1]))

    with st.spinner(f"前年（{prev_year}年）データ取得中..."):
        try:
            prev_raw = load_data(prev_start.isoformat(), prev_end.isoformat())
            df_prev = prev_raw[0]
        except Exception:
            df_prev = pd.DataFrame()

    if not df_prev.empty:
        if selected_authors:
            df_prev = df_prev[df_prev["author"].isin(selected_authors)].copy()

    if not df_prev.empty:
        cur_summary = {
            "年": str(end_date.year),
            "通常": df["hoursNormal"].sum(),
            "残業": df["hoursOT"].sum(),
        }
        cur_summary["合計"] = cur_summary["通常"] + cur_summary["残業"]

        prev_summary = {
            "年": str(prev_year),
            "通常": df_prev["hoursNormal"].sum(),
            "残業": df_prev["hoursOT"].sum(),
        }
        prev_summary["合計"] = prev_summary["通常"] + prev_summary["残業"]

        diff_total = cur_summary["合計"] - prev_summary["合計"]
        diff_ot = cur_summary["残業"] - prev_summary["残業"]
        pct_change = ((cur_summary["残業"] / prev_summary["残業"] - 1) * 100
                      if prev_summary["残業"] else 0)

        yoy_c1, yoy_c2, yoy_c3 = st.columns(3)
        with yoy_c1:
            delta_class = "good" if diff_total >= 0 else "warn"
            st.markdown(kpi_card(
                f"総工数 ({start_date.month}月〜{end_date.month}月)",
                f"{cur_summary['合計']:.1f}h",
                f"{diff_total:+.1f}h vs {prev_year}年",
                delta_class, "purple"
            ), unsafe_allow_html=True)
        with yoy_c2:
            delta_class = "bad" if diff_ot > 0 else "good"
            st.markdown(kpi_card(
                "残業時間",
                f"{cur_summary['残業']:.1f}h",
                f"{diff_ot:+.1f}h vs {prev_year}年",
                delta_class, "red"
            ), unsafe_allow_html=True)
        with yoy_c3:
            delta_class = "bad" if pct_change > 0 else "good"
            st.markdown(kpi_card(
                "残業増減率",
                f"{pct_change:+.1f}%",
                f"前年比",
                delta_class, "amber"
            ), unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)

        # 担当者別 前年比棒グラフ
        cur_auth = (
            df.groupby("author", as_index=False)
            .agg(残業_当年=("hoursOT", "sum"), 合計_当年=("hoursTotal", "sum"))
        )
        prev_auth = (
            df_prev.groupby("author", as_index=False)
            .agg(残業_前年=("hoursOT", "sum"), 合計_前年=("hoursTotal", "sum"))
        )
        yoy_auth = cur_auth.merge(prev_auth, on="author", how="outer").fillna(0)
        yoy_auth = sort_by_author(yoy_auth)

        fig_yoy = go.Figure()
        fig_yoy.add_trace(go.Bar(
            x=yoy_auth["author"], y=yoy_auth["残業_前年"],
            name=f"{prev_year}年 残業", marker_color="rgba(96,165,250,0.5)"
        ))
        fig_yoy.add_trace(go.Bar(
            x=yoy_auth["author"], y=yoy_auth["残業_当年"],
            name=f"{end_date.year}年 残業", marker_color="#F87171"
        ))
        fig_yoy.update_layout(
            barmode="group",
            title=f"担当者別 残業時間 前年比（{start_date.month}月〜{end_date.month}月）",
            yaxis_title="残業時間 (h)", height=450
        )
        apply_dark(fig_yoy)
        st.plotly_chart(fig_yoy, use_container_width=True)

        yoy_auth["残業増減"] = yoy_auth["残業_当年"] - yoy_auth["残業_前年"]
        yoy_auth["合計増減"] = yoy_auth["合計_当年"] - yoy_auth["合計_前年"]
        st.dataframe(
            yoy_auth[["author", "合計_前年", "合計_当年", "合計増減",
                       "残業_前年", "残業_当年", "残業増減"]].rename(
                columns={
                    "author": "担当者",
                    "合計_前年": f"合計{prev_year}", "合計_当年": f"合計{end_date.year}",
                    "合計増減": "合計増減",
                    "残業_前年": f"残業{prev_year}", "残業_当年": f"残業{end_date.year}",
                    "残業増減": "残業増減"
                }
            ),
            use_container_width=True, hide_index=True
        )
    else:
        st.info(f"{prev_year}年の同期間データがありません。")

# ────────────────────────────────────
# 個人サマリ
# ────────────────────────────────────
elif page == "🔍 個人サマリ":
    st.markdown('<p class="main-header">個人サマリ</p>', unsafe_allow_html=True)

    target_authors = selected_authors if selected_authors else AUTHORS
    available = [a for a in target_authors if a in df["author"].values]

    if not available:
        st.warning("担当者データがありません。")
    else:
        person = st.selectbox("担当者を選択", options=available)
        pdf = df[df["author"] == person]

        if pdf.empty:
            st.warning(f"{person} のデータがありません。")
        else:
            p_normal = pdf["hoursNormal"].sum()
            p_ot = pdf["hoursOT"].sum()
            p_total = p_normal + p_ot
            p_ot_rate = (p_ot / p_total * 100) if p_total else 0

            # 個人KPI
            pc1, pc2, pc3, pc4 = st.columns(4)
            with pc1:
                st.markdown(kpi_card("総工数", f"{p_total:.1f}h", accent="purple"), unsafe_allow_html=True)
            with pc2:
                st.markdown(kpi_card("通常", f"{p_normal:.1f}h", accent="blue"), unsafe_allow_html=True)
            with pc3:
                st.markdown(kpi_card("残業", f"{p_ot:.1f}h", accent="red"), unsafe_allow_html=True)
            with pc4:
                delta_class = "bad" if p_ot_rate > 20 else "warn" if p_ot_rate > 10 else "good"
                st.markdown(kpi_card("残業率", f"{p_ot_rate:.1f}%", delta_class=delta_class, accent="amber"), unsafe_allow_html=True)

            st.markdown("<br>", unsafe_allow_html=True)

            col_left, col_right = st.columns(2)

            with col_left:
                p_cat = (
                    pdf.groupby("mainLabel", as_index=False)
                    .agg(hours=("hoursTotal", "sum"))
                )
                p_cat = p_cat[p_cat["hours"] > 0]
                fig_p_pie = px.pie(
                    p_cat, values="hours", names="mainLabel",
                    color="mainLabel", color_discrete_map=cat_colors_dark,
                    title=f"{person} 作業構成", hole=0.4
                )
                fig_p_pie.update_traces(textinfo="label+percent+value")
                apply_dark(fig_p_pie)
                st.plotly_chart(fig_p_pie, use_container_width=True)

            with col_right:
                section_header("⚡", f"{person} 36協定進捗")

                st.markdown(f"**月残業: {p_ot:.1f}h / 45h**")
                st.progress(min(p_ot / 45, 1.0))
                if p_ot > 45:
                    st.markdown(f'<div class="alert-card danger">月上限超過 (+{p_ot - 45:.1f}h)</div>', unsafe_allow_html=True)
                elif p_ot > 36:
                    st.markdown(f'<div class="alert-card warning">残り {45 - p_ot:.1f}h</div>', unsafe_allow_html=True)
                else:
                    st.markdown('<div class="alert-card success">正常範囲</div>', unsafe_allow_html=True)

                try:
                    yr_data = load_monthly(
                        date(end_date.year, 1, 1).isoformat(),
                        str(end_date)
                    )
                    if not yr_data.empty:
                        yr_person = yr_data[yr_data["author"] == person]
                        yr_ot = yr_person["hoursOT"].sum()
                    else:
                        yr_ot = p_ot
                except Exception:
                    yr_ot = p_ot

                st.markdown(f"**年間累計残業: {yr_ot:.1f}h / 360h**")
                st.progress(min(yr_ot / 360, 1.0))
                if yr_ot > 360:
                    st.markdown(f'<div class="alert-card danger">年上限超過 (+{yr_ot - 360:.1f}h)</div>', unsafe_allow_html=True)
                elif yr_ot > 300:
                    st.markdown(f'<div class="alert-card warning">残り {360 - yr_ot:.1f}h</div>', unsafe_allow_html=True)
                else:
                    st.markdown('<div class="alert-card success">正常範囲</div>', unsafe_allow_html=True)

            # 中分類別テーブル
            section_header("📋", f"{person} 作業内訳")
            p_sub = (
                pdf.groupby(["cdSub", "subLabel", "mainLabel"], as_index=False)
                .agg(通常=("hoursNormal", "sum"), 残業=("hoursOT", "sum"))
            )
            p_sub["合計"] = p_sub["通常"] + p_sub["残業"]
            p_sub = p_sub.sort_values("合計", ascending=False)
            st.dataframe(
                p_sub.rename(columns={
                    "cdSub": "コード", "subLabel": "作業内容",
                    "mainLabel": "大分類"
                }),
                use_container_width=True, hide_index=True
            )

            # 月別残業推移（個人）
            section_header("📈", f"{person} 月別推移")
            try:
                yr_data_full = load_monthly(
                    date(end_date.year, 1, 1).isoformat(),
                    str(end_date)
                )
                if not yr_data_full.empty:
                    p_monthly = (
                        yr_data_full[yr_data_full["author"] == person]
                        .groupby("month", as_index=False)
                        .agg(残業=("hoursOT", "sum"), 通常=("hoursNormal", "sum"))
                    )
                    p_monthly["月名"] = p_monthly["month"].map(lambda m: f"{m}月")

                    fig_p_month = go.Figure()
                    fig_p_month.add_trace(go.Bar(
                        x=p_monthly["月名"], y=p_monthly["通常"],
                        name="通常", marker_color="#60A5FA"
                    ))
                    fig_p_month.add_trace(go.Bar(
                        x=p_monthly["月名"], y=p_monthly["残業"],
                        name="残業", marker_color="#F87171"
                    ))
                    fig_p_month.add_hline(
                        y=45, line_dash="dash", line_color="#EF4444",
                        annotation_text="月上限 45h", annotation_font_color="#EF4444"
                    )
                    fig_p_month.update_layout(
                        barmode="stack",
                        title=f"{person} {end_date.year}年 月別工数",
                        yaxis_title="時間 (h)", height=400
                    )
                    apply_dark(fig_p_month)
                    st.plotly_chart(fig_p_month, use_container_width=True)
            except Exception:
                st.info("月別データを取得できませんでした。")

# ────────────────────────────────────
# データ一覧
# ────────────────────────────────────
elif page == "📋 データ一覧":
    st.markdown('<p class="main-header">データ一覧</p>', unsafe_allow_html=True)
    st.markdown(f'<p class="sub-header">{start_date.strftime("%Y/%m/%d")} — {end_date.strftime("%Y/%m/%d")}　|　{len(df)} 件</p>', unsafe_allow_html=True)

    display_df = df[["author", "cdSub", "subLabel", "mainLabel",
                     "hoursNormal", "hoursOT", "hoursTotal"]].copy()
    display_df.columns = ["担当者", "コード", "作業内容", "大分類", "通常(h)", "残業(h)", "合計(h)"]
    st.dataframe(display_df, use_container_width=True, hide_index=True)

    csv = display_df.to_csv(index=False).encode("utf-8-sig")
    st.download_button("📥 CSVダウンロード", csv, "npa_export.csv", "text/csv")
