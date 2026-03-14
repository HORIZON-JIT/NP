"""
NPA (日報分析) - Streamlit ダッシュボード

起動: streamlit run npa/app.py
"""

import streamlit as st
import plotly.express as px
import plotly.graph_objects as go
import pandas as pd
from datetime import date, timedelta

from config import MAIN_CATEGORIES, AUTHORS, SUB_CATEGORIES, get_main_cd, get_main_label
from fetch_data import fetch_date_range, to_dataframe

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

# ── データ取得 ──
@st.cache_data(ttl=300, show_spinner=False)
def load_data(s: str, e: str):
    raw = fetch_date_range(s, e)
    return to_dataframe(raw), raw

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

# ── カラー定義 ──
cat_colors = {v["label"]: v["color"] for k, v in MAIN_CATEGORIES.items()}

# ══════════════════════════════════════════
# KPI カード
# ══════════════════════════════════════════
total_normal = df["hoursNormal"].sum()
total_ot = df["hoursOT"].sum()
total_all = total_normal + total_ot
n_authors = df["author"].nunique()

k1, k2, k3, k4 = st.columns(4)
k1.metric("総工数", f"{total_all:.1f}h")
k2.metric("通常時間", f"{total_normal:.1f}h")
k3.metric("残業時間", f"{total_ot:.1f}h", delta=f"{total_ot/total_all*100:.0f}%" if total_all else None)
k4.metric("担当者数", f"{n_authors}名")

st.divider()

# ══════════════════════════════════════════
# タブ構成
# ══════════════════════════════════════════
tab1, tab2, tab3, tab4 = st.tabs(["👤 担当者別", "📂 工程別", "⏰ 残業分析", "📋 データ一覧"])

# ── タブ1: 担当者別 ──
with tab1:
    # 担当者別 通常/残業 積み上げ棒
    auth_df = (
        df.groupby("author", as_index=False)
        .agg(通常=("hoursNormal", "sum"), 残業=("hoursOT", "sum"))
    )
    auth_df["合計"] = auth_df["通常"] + auth_df["残業"]
    # ソート
    author_order = {a: i for i, a in enumerate(AUTHORS)}
    auth_df["_o"] = auth_df["author"].map(lambda a: author_order.get(a, 99))
    auth_df = auth_df.sort_values("_o").drop(columns="_o")

    fig = go.Figure()
    fig.add_trace(go.Bar(x=auth_df["author"], y=auth_df["通常"],
                         name="通常", marker_color="#3b82f6"))
    fig.add_trace(go.Bar(x=auth_df["author"], y=auth_df["残業"],
                         name="残業", marker_color="#ef4444"))
    fig.update_layout(barmode="stack", title="担当者別工数",
                      yaxis_title="時間 (h)", height=450)
    st.plotly_chart(fig, use_container_width=True)

    # 担当者×大分類 積み上げ棒
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

# ── タブ2: 工程別 ──
with tab2:
    col_a, col_b = st.columns(2)

    # 大分類別 円グラフ
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

    # 大分類別 棒グラフ
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

    # 中分類別テーブル
    sub_df = (
        df.groupby(["cdSub", "subLabel", "mainLabel"], as_index=False)
        .agg(通常=("hoursNormal", "sum"), 残業=("hoursOT", "sum"))
    )
    sub_df["合計"] = sub_df["通常"] + sub_df["残業"]
    sub_df = sub_df.sort_values("cdSub")
    sub_df = sub_df.rename(columns={"cdSub": "コード", "subLabel": "作業内容", "mainLabel": "大分類"})
    st.subheader("中分類別 工数一覧")
    st.dataframe(sub_df, use_container_width=True, hide_index=True)

# ── タブ3: 残業分析 ──
with tab3:
    # 担当者別残業
    ot_df = (
        df.groupby("author", as_index=False)
        .agg(残業=("hoursOT", "sum"))
    )
    ot_df["_o"] = ot_df["author"].map(lambda a: author_order.get(a, 99))
    ot_df = ot_df.sort_values("_o").drop(columns="_o")

    fig5 = px.bar(ot_df, x="author", y="残業",
                  title="担当者別 残業時間",
                  labels={"残業": "残業時間 (h)", "author": "担当者"},
                  color_discrete_sequence=["#ef4444"])
    fig5.update_layout(height=400)
    st.plotly_chart(fig5, use_container_width=True)

    # 残業比率
    auth_full = (
        df.groupby("author", as_index=False)
        .agg(通常=("hoursNormal", "sum"), 残業=("hoursOT", "sum"))
    )
    auth_full["合計"] = auth_full["通常"] + auth_full["残業"]
    auth_full["残業率"] = (auth_full["残業"] / auth_full["合計"] * 100).round(1)
    auth_full["_o"] = auth_full["author"].map(lambda a: author_order.get(a, 99))
    auth_full = auth_full.sort_values("_o").drop(columns="_o")

    st.subheader("担当者別 残業比率")
    st.dataframe(
        auth_full[["author", "通常", "残業", "合計", "残業率"]].rename(
            columns={"author": "担当者", "残業率": "残業率 (%)"}
        ),
        use_container_width=True, hide_index=True
    )

# ── タブ4: データ一覧 ──
with tab4:
    st.subheader("生データ")
    display_df = df[["author", "cdSub", "subLabel", "mainLabel",
                     "hoursNormal", "hoursOT", "hoursTotal"]].copy()
    display_df.columns = ["担当者", "コード", "作業内容", "大分類", "通常(h)", "残業(h)", "合計(h)"]
    st.dataframe(display_df, use_container_width=True, hide_index=True)

    # CSV出力
    csv = display_df.to_csv(index=False).encode("utf-8-sig")
    st.download_button("📥 CSVダウンロード", csv, "npa_export.csv", "text/csv")
