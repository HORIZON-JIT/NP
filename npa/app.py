"""
NPA (日報分析) - Streamlit ダッシュボード

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
st.set_page_config(page_title="日報分析 (NPA)", page_icon="📊", layout="wide")
st.title("📊 日報分析ダッシュボード")

# ── サイドバー：期間・フィルタ ──
st.sidebar.header("分析条件")

today = date.today()
default_start = today.replace(day=1)
default_end = today

col_s, col_e = st.sidebar.columns(2)
start_date = col_s.date_input("開始日", value=default_start)
end_date = col_e.date_input("終了日", value=default_end)

if start_date > end_date:
    st.sidebar.error("開始日が終了日より後です")
    st.stop()

# 担当者フィルタ
selected_authors = st.sidebar.multiselect(
    "担当者（空=全員）", options=AUTHORS, default=[]
)

# ── 共通ユーティリティ ──
author_order = {a: i for i, a in enumerate(AUTHORS)}
cat_colors = {v["label"]: v["color"] for k, v in MAIN_CATEGORIES.items()}


def sort_by_author(frame: pd.DataFrame) -> pd.DataFrame:
    frame["_o"] = frame["author"].map(lambda a: author_order.get(a, 99))
    return frame.sort_values("_o").drop(columns="_o")


def business_days(start: date, end: date) -> int:
    """期間内の営業日数（土日除外）"""
    return int(np.busday_count(start, end + timedelta(days=1)))


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

# ══════════════════════════════════════════
# KPI カード
# ══════════════════════════════════════════
total_normal = df["hoursNormal"].sum()
total_ot = df["hoursOT"].sum()
total_all = total_normal + total_ot
n_authors = df["author"].nunique()
bdays = business_days(start_date, end_date)
avg_per_day = total_all / bdays if bdays else 0

k1, k2, k3, k4, k5 = st.columns(5)
k1.metric("総工数", f"{total_all:.1f}h")
k2.metric("通常時間", f"{total_normal:.1f}h")
k3.metric("残業時間", f"{total_ot:.1f}h",
          delta=f"{total_ot/total_all*100:.0f}%" if total_all else None)
k4.metric("担当者数", f"{n_authors}名")
k5.metric("1日平均/人", f"{avg_per_day/n_authors:.1f}h" if n_authors else "-")

st.divider()

# ══════════════════════════════════════════
# タブ構成
# ══════════════════════════════════════════
tab1, tab2, tab3, tab4, tab5, tab5b, tab6, tab7, tab8 = st.tabs([
    "👤 担当者別", "📂 工程別", "⏰ 残業・36協定",
    "📈 推移分析", "📊 稼働率", "🔥 作業偏り",
    "📅 月別・前年比", "🔍 個人サマリ", "📋 データ一覧"
])

# ══════════════════════════════════════════
# タブ1: 担当者別（既存）
# ══════════════════════════════════════════
with tab1:
    auth_df = (
        df.groupby("author", as_index=False)
        .agg(通常=("hoursNormal", "sum"), 残業=("hoursOT", "sum"))
    )
    auth_df["合計"] = auth_df["通常"] + auth_df["残業"]
    auth_df = sort_by_author(auth_df)

    fig = go.Figure()
    fig.add_trace(go.Bar(x=auth_df["author"], y=auth_df["通常"],
                         name="通常", marker_color="#3b82f6"))
    fig.add_trace(go.Bar(x=auth_df["author"], y=auth_df["残業"],
                         name="残業", marker_color="#ef4444"))
    fig.update_layout(barmode="stack", title="担当者別工数",
                      yaxis_title="時間 (h)", height=450)
    st.plotly_chart(fig, use_container_width=True)

    cross = (
        df.groupby(["author", "mainLabel"], as_index=False)
        .agg(hours=("hoursTotal", "sum"))
    )
    fig2 = px.bar(cross, x="author", y="hours", color="mainLabel",
                  color_discrete_map=cat_colors,
                  title="担当者×大分類 工数内訳",
                  labels={"hours": "時間 (h)", "author": "担当者", "mainLabel": "大分類"})
    fig2.update_layout(barmode="stack", height=450)
    st.plotly_chart(fig2, use_container_width=True)

# ══════════════════════════════════════════
# タブ2: 工程別（既存）
# ══════════════════════════════════════════
with tab2:
    col_a, col_b = st.columns(2)

    cat_df = (
        df.groupby(["mainCd", "mainLabel"], as_index=False)
        .agg(hours=("hoursTotal", "sum"))
    )
    cat_df = cat_df[cat_df["hours"] > 0]

    with col_a:
        fig3 = px.pie(cat_df, values="hours", names="mainLabel",
                      color="mainLabel", color_discrete_map=cat_colors,
                      title="大分類別 工数比率")
        fig3.update_traces(textinfo="label+percent+value")
        st.plotly_chart(fig3, use_container_width=True)

    with col_b:
        cat_order = list(MAIN_CATEGORIES.keys())
        cat_df["_o"] = cat_df["mainCd"].map(lambda c: cat_order.index(c) if c in cat_order else 99)
        cat_df = cat_df.sort_values("_o")
        fig4 = px.bar(cat_df, x="mainLabel", y="hours",
                      color="mainLabel", color_discrete_map=cat_colors,
                      title="大分類別 工数",
                      labels={"hours": "時間 (h)", "mainLabel": "大分類"})
        fig4.update_layout(showlegend=False)
        st.plotly_chart(fig4, use_container_width=True)

    sub_df = (
        df.groupby(["cdSub", "subLabel", "mainLabel"], as_index=False)
        .agg(通常=("hoursNormal", "sum"), 残業=("hoursOT", "sum"))
    )
    sub_df["合計"] = sub_df["通常"] + sub_df["残業"]
    sub_df = sub_df.sort_values("cdSub")
    sub_df = sub_df.rename(columns={"cdSub": "コード", "subLabel": "作業内容", "mainLabel": "大分類"})
    st.subheader("中分類別 工数一覧")
    st.dataframe(sub_df, use_container_width=True, hide_index=True)

# ══════════════════════════════════════════
# タブ3: 残業・36協定
# ══════════════════════════════════════════
with tab3:
    # 担当者別残業棒グラフ
    ot_df = (
        df.groupby("author", as_index=False)
        .agg(残業=("hoursOT", "sum"))
    )
    ot_df = sort_by_author(ot_df)

    fig5 = px.bar(ot_df, x="author", y="残業",
                  title="担当者別 残業時間",
                  labels={"残業": "残業時間 (h)", "author": "担当者"},
                  color_discrete_sequence=["#ef4444"])
    fig5.update_layout(height=400)
    st.plotly_chart(fig5, use_container_width=True)

    # 残業比率テーブル
    auth_full = (
        df.groupby("author", as_index=False)
        .agg(通常=("hoursNormal", "sum"), 残業=("hoursOT", "sum"))
    )
    auth_full["合計"] = auth_full["通常"] + auth_full["残業"]
    auth_full["残業率"] = (auth_full["残業"] / auth_full["合計"] * 100).round(1)
    auth_full = sort_by_author(auth_full)

    st.subheader("担当者別 残業比率")
    st.dataframe(
        auth_full[["author", "通常", "残業", "合計", "残業率"]].rename(
            columns={"author": "担当者", "残業率": "残業率 (%)"}
        ),
        use_container_width=True, hide_index=True
    )

    # ── 36協定アラート ──
    st.subheader("36協定 残業上限チェック")
    st.caption("月45h / 年360h が上限（特別条項: 月100h / 年720h）")

    # 当月の残業（選択期間が1ヶ月以内ならそのまま使用）
    ot_by_author = (
        df.groupby("author", as_index=False)
        .agg(月残業=("hoursOT", "sum"))
    )
    ot_by_author = sort_by_author(ot_by_author)

    # 年間累計（1月〜現在の選択終了月まで）
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
            color_m = "red" if monthly_ot > 45 else ("orange" if monthly_ot > 36 else "green")
            st.markdown(f"月: **{monthly_ot:.1f}h** / 45h")
            st.progress(min(pct_m / 100, 1.0))
            if monthly_ot > 45:
                st.error(f"月上限超過 (+{monthly_ot - 45:.1f}h)")
            elif monthly_ot > 36:
                st.warning(f"月上限まで残り {45 - monthly_ot:.1f}h")
        with col_year:
            pct_y = min(yearly_ot / 360 * 100, 100)
            st.markdown(f"年: **{yearly_ot:.1f}h** / 360h")
            st.progress(min(pct_y / 100, 1.0))
            if yearly_ot > 360:
                st.error(f"年上限超過 (+{yearly_ot - 360:.1f}h)")
            elif yearly_ot > 300:
                st.warning(f"年上限まで残り {360 - yearly_ot:.1f}h")

# ══════════════════════════════════════════
# タブ4: 推移分析
# ══════════════════════════════════════════
with tab4:
    st.subheader("週別推移")

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
            markers=True
        )
        fig_trend_ot.update_layout(height=400)
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
            name="通常", marker_color="#3b82f6"
        ))
        fig_trend_total.add_trace(go.Bar(
            x=wk_total["week_start"], y=wk_total["残業"],
            name="残業", marker_color="#ef4444"
        ))
        fig_trend_total.update_layout(
            barmode="stack", title="週別 全体工数推移",
            xaxis_title="週", yaxis_title="時間 (h)", height=400
        )
        st.plotly_chart(fig_trend_total, use_container_width=True)

        # 週別×大分類 推移
        wk_cat = (
            weekly_df.groupby(["week_start", "mainLabel"], as_index=False)
            .agg(hours=("hoursTotal", "sum"))
        )
        fig_trend_cat = px.bar(
            wk_cat, x="week_start", y="hours", color="mainLabel",
            color_discrete_map=cat_colors,
            title="週別 大分類別工数推移",
            labels={"week_start": "週", "hours": "時間 (h)", "mainLabel": "大分類"}
        )
        fig_trend_cat.update_layout(barmode="stack", height=400)
        st.plotly_chart(fig_trend_cat, use_container_width=True)
    else:
        st.info("週別データがありません。")

# ══════════════════════════════════════════
# タブ5: 稼働率・偏り分析
# ══════════════════════════════════════════
with tab5:
    # ── 稼働率 ──
    st.subheader("稼働率（所定労働時間に対する実績）")
    st.caption(f"所定: 営業日{bdays}日 × {REGULAR_HOURS}h = {bdays * REGULAR_HOURS:.0f}h/人")

    expected_hours = bdays * REGULAR_HOURS

    util_df = (
        df.groupby("author", as_index=False)
        .agg(通常=("hoursNormal", "sum"), 残業=("hoursOT", "sum"))
    )
    util_df["合計"] = util_df["通常"] + util_df["残業"]
    util_df["所定"] = expected_hours
    util_df["稼働率"] = (util_df["合計"] / expected_hours * 100).round(1) if expected_hours else 0
    util_df = sort_by_author(util_df)

    # 稼働率棒グラフ
    fig_util = go.Figure()
    fig_util.add_trace(go.Bar(
        x=util_df["author"], y=util_df["稼働率"],
        marker_color=[
            "#ef4444" if r > 120 else "#eab308" if r > 105 else "#22c55e"
            for r in util_df["稼働率"]
        ],
        text=[f"{r:.0f}%" for r in util_df["稼働率"]],
        textposition="outside"
    ))
    fig_util.add_hline(y=100, line_dash="dash", line_color="gray",
                       annotation_text="所定100%")
    fig_util.update_layout(
        title="担当者別 稼働率",
        yaxis_title="稼働率 (%)", height=400
    )
    st.plotly_chart(fig_util, use_container_width=True)

    # 稼働率テーブル
    st.dataframe(
        util_df[["author", "通常", "残業", "合計", "所定", "稼働率"]].rename(
            columns={"author": "担当者", "稼働率": "稼働率 (%)"}
        ),
        use_container_width=True, hide_index=True
    )

# ══════════════════════════════════════════
# タブ5b: 作業偏り分析（専用タブ）
# ══════════════════════════════════════════
with tab5b:

    # ── 共通データ準備 ──
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

    # ── 1. 工数ヒートマップ（実時間） ──
    st.subheader("担当者 × 大分類 ヒートマップ")

    heat_mode = st.radio(
        "表示モード", ["実時間 (h)", "構成比 (%)", "チーム平均との差分"],
        horizontal=True, key="heat_mode"
    )

    if heat_mode == "実時間 (h)":
        fig_heat = px.imshow(
            heat_pivot, text_auto=".1f",
            color_continuous_scale="YlOrRd",
            labels={"x": "大分類", "y": "担当者", "color": "時間 (h)"},
            title="工数ヒートマップ（実時間）"
        )
    elif heat_mode == "構成比 (%)":
        heat_pct = heat_pivot.div(heat_pivot.sum(axis=1), axis=0) * 100
        fig_heat = px.imshow(
            heat_pct.round(1), text_auto=".1f",
            color_continuous_scale="Blues",
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
    st.plotly_chart(fig_heat, use_container_width=True)

    st.divider()

    # ── 2. レーダーチャート（担当者比較） ──
    st.subheader("作業バランス レーダーチャート")

    radar_authors = st.multiselect(
        "比較する担当者（空=全員）",
        options=ordered_authors, default=[],
        key="radar_authors"
    )
    radar_targets = radar_authors if radar_authors else ordered_authors

    # 構成比に正規化
    heat_pct_radar = heat_pivot.div(heat_pivot.sum(axis=1), axis=0) * 100
    heat_pct_radar = heat_pct_radar.fillna(0)

    fig_radar = go.Figure()
    for person in radar_targets:
        if person in heat_pct_radar.index:
            vals = heat_pct_radar.loc[person].tolist()
            vals.append(vals[0])  # 閉じる
            cats_loop = ordered_cats + [ordered_cats[0]]
            fig_radar.add_trace(go.Scatterpolar(
                r=vals, theta=cats_loop, name=person,
                fill="toself", opacity=0.3
            ))

    # チーム平均をオーバーレイ
    team_avg_pct = heat_pct_radar.mean(axis=0).tolist()
    team_avg_pct.append(team_avg_pct[0])
    fig_radar.add_trace(go.Scatterpolar(
        r=team_avg_pct,
        theta=ordered_cats + [ordered_cats[0]],
        name="チーム平均",
        line=dict(dash="dash", color="black", width=2),
        opacity=0.8
    ))

    fig_radar.update_layout(
        polar=dict(radialaxis=dict(visible=True, range=[0, 100])),
        title="作業構成比レーダー（チーム平均は破線）",
        height=500
    )
    st.plotly_chart(fig_radar, use_container_width=True)

    st.divider()

    # ── 3. 中分類ヒートマップ（ドリルダウン） ──
    st.subheader("中分類ヒートマップ（詳細）")

    sub_heat = (
        df.groupby(["author", "subLabel"], as_index=False)
        .agg(hours=("hoursTotal", "sum"))
    )
    sub_pivot = sub_heat.pivot_table(
        index="author", columns="subLabel", values="hours", fill_value=0
    )
    sub_pivot = sub_pivot.reindex(ordered_authors)
    # 合計が大きい順に列ソート
    col_order = sub_pivot.sum().sort_values(ascending=False).index.tolist()
    sub_pivot = sub_pivot[col_order]

    fig_sub_heat = px.imshow(
        sub_pivot, text_auto=".1f",
        color_continuous_scale="Viridis",
        labels={"x": "中分類", "y": "担当者", "color": "時間 (h)"},
        title="担当者 × 中分類 ヒートマップ"
    )
    fig_sub_heat.update_layout(
        height=max(400, len(ordered_authors) * 50),
        xaxis=dict(tickangle=-45)
    )
    st.plotly_chart(fig_sub_heat, use_container_width=True)

    st.divider()

    # ── 4. 作業集中度分析 ──
    st.subheader("作業集中度")
    st.caption("一つの大分類に工数が集中している度合いを分析")

    conc_df = heat_df.copy()
    author_totals = conc_df.groupby("author")["hours"].transform("sum")
    conc_df["比率"] = (conc_df["hours"] / author_totals * 100).round(1)

    # HHI（ハーフィンダール指数）で集中度を数値化
    hhi_df = conc_df.groupby("author").apply(
        lambda g: (g["比率"] ** 2).sum(), include_groups=False
    ).reset_index(name="HHI")
    hhi_df = sort_by_author(hhi_df)
    # HHIの目安: 10000=完全集中（1つだけ）、~1250=均等（8分類）
    hhi_df["集中レベル"] = hhi_df["HHI"].map(
        lambda h: "⚠️ 高い" if h > 5000 else ("△ やや高い" if h > 3000 else "○ 分散")
    )

    col_hhi_chart, col_hhi_table = st.columns([2, 1])

    with col_hhi_chart:
        fig_hhi = go.Figure()
        fig_hhi.add_trace(go.Bar(
            x=hhi_df["author"], y=hhi_df["HHI"],
            marker_color=[
                "#ef4444" if h > 5000 else "#eab308" if h > 3000 else "#22c55e"
                for h in hhi_df["HHI"]
            ],
            text=hhi_df["集中レベル"],
            textposition="outside"
        ))
        fig_hhi.add_hline(y=3000, line_dash="dash", line_color="orange",
                          annotation_text="集中度しきい値")
        fig_hhi.update_layout(
            title="作業集中度 (HHI指数)",
            yaxis_title="HHI", height=400
        )
        st.plotly_chart(fig_hhi, use_container_width=True)

    with col_hhi_table:
        st.markdown("**HHI指数の見方**")
        st.markdown("""
        - **~1,250**: 全分類に均等配分
        - **3,000超**: やや集中傾向
        - **5,000超**: 特定作業に偏り
        - **10,000**: 1つの作業のみ
        """)

        # 50%超の集中リスト
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
            st.success("50%超の集中なし")

    st.divider()

    # ── 5. 担当者間の類似度 ──
    st.subheader("担当者間 作業パターン類似度")
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
        color_continuous_scale="Greens",
        zmin=0, zmax=1,
        labels={"x": "担当者", "y": "担当者", "color": "類似度"},
        title="作業パターン類似度マトリクス"
    )
    fig_sim.update_layout(height=max(400, len(authors_list) * 55))
    st.plotly_chart(fig_sim, use_container_width=True)

# ══════════════════════════════════════════
# タブ6: 月別・前年比
# ══════════════════════════════════════════
with tab6:
    st.subheader("月別比較")

    # 当年の月別データ
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

        # 月別 通常/残業 推移
        m_agg = (
            monthly_df.groupby("month", as_index=False)
            .agg(通常=("hoursNormal", "sum"), 残業=("hoursOT", "sum"))
        )
        m_agg["月名"] = m_agg["month"].map(lambda m: f"{m}月")

        fig_month = go.Figure()
        fig_month.add_trace(go.Bar(x=m_agg["月名"], y=m_agg["通常"],
                                   name="通常", marker_color="#3b82f6"))
        fig_month.add_trace(go.Bar(x=m_agg["月名"], y=m_agg["残業"],
                                   name="残業", marker_color="#ef4444"))
        fig_month.update_layout(barmode="stack", title=f"{end_date.year}年 月別工数推移",
                                yaxis_title="時間 (h)", height=400)
        st.plotly_chart(fig_month, use_container_width=True)

        # 月別×担当者 残業推移
        m_ot = (
            monthly_df.groupby(["month", "author"], as_index=False)
            .agg(残業=("hoursOT", "sum"))
        )
        m_ot["月名"] = m_ot["month"].map(lambda m: f"{m}月")

        fig_m_ot = px.line(
            m_ot, x="月名", y="残業", color="author",
            title=f"{end_date.year}年 月別残業推移（担当者別）",
            labels={"月名": "月", "残業": "残業時間 (h)", "author": "担当者"},
            markers=True
        )
        fig_m_ot.update_layout(height=400)
        st.plotly_chart(fig_m_ot, use_container_width=True)
    else:
        st.info("月別データがありません。")

    # ── 前年比 ──
    st.divider()
    st.subheader("前年比")

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
        # 当年 vs 前年 サマリ
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

        cmp_df = pd.DataFrame([prev_summary, cur_summary])
        cmp_df["残業率"] = (cmp_df["残業"] / cmp_df["合計"] * 100).round(1)

        yoy_c1, yoy_c2, yoy_c3 = st.columns(3)
        diff_total = cur_summary["合計"] - prev_summary["合計"]
        diff_ot = cur_summary["残業"] - prev_summary["残業"]
        with yoy_c1:
            st.metric(
                f"総工数 ({start_date.month}月〜{end_date.month}月)",
                f"{cur_summary['合計']:.1f}h",
                delta=f"{diff_total:+.1f}h vs {prev_year}年"
            )
        with yoy_c2:
            st.metric(
                "残業時間",
                f"{cur_summary['残業']:.1f}h",
                delta=f"{diff_ot:+.1f}h vs {prev_year}年",
                delta_color="inverse"
            )
        with yoy_c3:
            pct_change = ((cur_summary["残業"] / prev_summary["残業"] - 1) * 100
                          if prev_summary["残業"] else 0)
            st.metric(
                "残業増減率",
                f"{pct_change:+.1f}%",
                delta=f"前年比",
                delta_color="inverse"
            )

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
            name=f"{prev_year}年 残業", marker_color="#93c5fd"
        ))
        fig_yoy.add_trace(go.Bar(
            x=yoy_auth["author"], y=yoy_auth["残業_当年"],
            name=f"{end_date.year}年 残業", marker_color="#ef4444"
        ))
        fig_yoy.update_layout(
            barmode="group",
            title=f"担当者別 残業時間 前年比（{start_date.month}月〜{end_date.month}月）",
            yaxis_title="残業時間 (h)", height=450
        )
        st.plotly_chart(fig_yoy, use_container_width=True)

        # 前年比テーブル
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

# ══════════════════════════════════════════
# タブ7: 個人サマリ
# ══════════════════════════════════════════
with tab7:
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
            pc1.metric("総工数", f"{p_total:.1f}h")
            pc2.metric("通常", f"{p_normal:.1f}h")
            pc3.metric("残業", f"{p_ot:.1f}h")
            pc4.metric("残業率", f"{p_ot_rate:.1f}%")

            col_left, col_right = st.columns(2)

            with col_left:
                # 作業構成（円グラフ）
                p_cat = (
                    pdf.groupby("mainLabel", as_index=False)
                    .agg(hours=("hoursTotal", "sum"))
                )
                p_cat = p_cat[p_cat["hours"] > 0]
                fig_p_pie = px.pie(
                    p_cat, values="hours", names="mainLabel",
                    color="mainLabel", color_discrete_map=cat_colors,
                    title=f"{person} 作業構成"
                )
                fig_p_pie.update_traces(textinfo="label+percent+value")
                st.plotly_chart(fig_p_pie, use_container_width=True)

            with col_right:
                # 36協定進捗
                st.markdown(f"### {person} 36協定進捗")

                # 月残業
                st.markdown(f"**月残業: {p_ot:.1f}h / 45h**")
                st.progress(min(p_ot / 45, 1.0))
                if p_ot > 45:
                    st.error(f"月上限超過 (+{p_ot - 45:.1f}h)")
                elif p_ot > 36:
                    st.warning(f"残り {45 - p_ot:.1f}h")
                else:
                    st.success("正常範囲")

                # 年間累計
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
                    st.error(f"年上限超過 (+{yr_ot - 360:.1f}h)")
                elif yr_ot > 300:
                    st.warning(f"残り {360 - yr_ot:.1f}h")
                else:
                    st.success("正常範囲")

            # 中分類別テーブル
            st.subheader(f"{person} 作業内訳")
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
            st.subheader(f"{person} 月別残業推移")
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
                        name="通常", marker_color="#3b82f6"
                    ))
                    fig_p_month.add_trace(go.Bar(
                        x=p_monthly["月名"], y=p_monthly["残業"],
                        name="残業", marker_color="#ef4444"
                    ))
                    fig_p_month.add_hline(
                        y=45, line_dash="dash", line_color="red",
                        annotation_text="月上限 45h"
                    )
                    fig_p_month.update_layout(
                        barmode="stack",
                        title=f"{person} {end_date.year}年 月別工数",
                        yaxis_title="時間 (h)", height=400
                    )
                    st.plotly_chart(fig_p_month, use_container_width=True)
            except Exception:
                st.info("月別データを取得できませんでした。")

# ══════════════════════════════════════════
# タブ8: データ一覧（既存）
# ══════════════════════════════════════════
with tab8:
    st.subheader("生データ")
    display_df = df[["author", "cdSub", "subLabel", "mainLabel",
                     "hoursNormal", "hoursOT", "hoursTotal"]].copy()
    display_df.columns = ["担当者", "コード", "作業内容", "大分類", "通常(h)", "残業(h)", "合計(h)"]
    st.dataframe(display_df, use_container_width=True, hide_index=True)

    csv = display_df.to_csv(index=False).encode("utf-8-sig")
    st.download_button("📥 CSVダウンロード", csv, "npa_export.csv", "text/csv")
