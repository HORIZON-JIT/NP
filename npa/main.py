#!/usr/bin/env python3
"""
NPA (日報分析) - CLIエントリーポイント

使い方:
    python main.py --start 2025-01-01 --end 2025-12-31
    python main.py --start 2025-04-01 --end 2025-04-30 --text-only
"""

import argparse
import sys
from datetime import date

from fetch_data import fetch_date_range, to_dataframe
from analyze import (
    by_author,
    by_main_category,
    by_sub_category,
    by_author_and_category,
    overtime_by_author,
    summary_text,
)
from visualize import (
    bar_by_author,
    pie_by_category,
    bar_by_category,
    bar_overtime,
    stacked_author_category,
)


def main():
    parser = argparse.ArgumentParser(description="日報データ分析ツール (NPA)")
    parser.add_argument("--start", required=True, help="開始日 (YYYY-MM-DD)")
    parser.add_argument("--end", required=True, help="終了日 (YYYY-MM-DD)")
    parser.add_argument("--text-only", action="store_true",
                        help="テキストサマリーのみ出力（グラフ生成なし）")
    args = parser.parse_args()

    print(f"データ取得中: {args.start} 〜 {args.end} ...")
    try:
        raw = fetch_date_range(args.start, args.end)
    except Exception as e:
        print(f"エラー: {e}", file=sys.stderr)
        sys.exit(1)

    df = to_dataframe(raw)
    if df.empty:
        print("データがありません。")
        sys.exit(0)

    print(f"取得完了: {len(df)}行\n")

    # テキストサマリー
    print(summary_text(df, args.start, args.end))

    if args.text_only:
        return

    # グラフ生成
    print("\nグラフ生成中...")

    period = f"{args.start}〜{args.end}"

    auth_df = by_author(df)
    path = bar_by_author(auth_df, f"担当者別工数 ({period})")
    if path:
        print(f"  -> {path}")

    cat_df = by_main_category(df)
    path = pie_by_category(cat_df, f"大分類別工数比率 ({period})")
    if path:
        print(f"  -> {path}")

    path = bar_by_category(cat_df, f"大分類別工数 ({period})")
    if path:
        print(f"  -> {path}")

    ot_df = overtime_by_author(df)
    path = bar_overtime(ot_df, f"担当者別残業時間 ({period})")
    if path:
        print(f"  -> {path}")

    cross_df = by_author_and_category(df)
    path = stacked_author_category(cross_df, f"担当者×大分類 工数内訳 ({period})")
    if path:
        print(f"  -> {path}")

    print("\n完了!")


if __name__ == "__main__":
    main()
