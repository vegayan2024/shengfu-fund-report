#!/usr/bin/env python3
"""
基金净值数据更新脚本
用法: python scripts/update_nav.py

读取 data/ 目录下的 Excel 文件，生成对应的 JSON 数据文件。
自动从 akshare 获取沪深300基准数据并完成日期对齐。
"""
import sys
import os
import json
import glob
from datetime import datetime

import xlrd
import akshare as ak

# ============================================================
# 配置
# ============================================================
PRODUCTS = {
    "zhouqi": {"name": "泽鑫周期", "json_path": "zhouqi/combined-data.json"},
    "jiazhi": {"name": "泽鑫价值", "json_path": "jiazhi/combined-data.json"},
}
DATA_DIR = "data"
BENCHMARK_START = "2024-07-01"  # 沪深300数据起始日期


def find_latest_excel(product_key):
    """查找 data/ 目录下该产品的最新 Excel 文件"""
    pattern = os.path.join(DATA_DIR, f"{product_key}_*.xlsx")
    files = glob.glob(pattern)
    if not files:
        # 尝试 .xls 格式
        pattern = os.path.join(DATA_DIR, f"{product_key}_*.xls")
        files = glob.glob(pattern)
    if not files:
        return None
    # 按文件名排序，取最新的
    files.sort(reverse=True)
    return files[0]


def read_excel_nav(filepath):
    """从 Excel 文件读取净值数据，返回 [(date_str, nav_value), ...]"""
    print(f"  读取文件: {filepath}")
    wb = xlrd.open_workbook(filepath, encoding_override="utf-8")
    sheet = wb.sheet_by_index(0)

    records = []
    for row_idx in range(1, sheet.nrows):  # 跳过表头
        cell_date = sheet.cell_value(row_idx, 0)
        cell_nav = sheet.cell_value(row_idx, 1)

        if not cell_date or not cell_nav:
            continue

        # 处理日期格式
        if isinstance(cell_date, float):
            # Excel 日期序列号
            date_tuple = xlrd.xldate_as_tuple(cell_date, wb.datemode)
            date_str = f"{date_tuple[0]}-{date_tuple[1]:02d}-{date_tuple[2]:02d}"
        else:
            date_str = str(cell_date).strip()[:10]

        # 处理净值
        try:
            nav = float(cell_nav)
        except (ValueError, TypeError):
            continue

        records.append((date_str, nav))

    print(f"  读取到 {len(records)} 条记录")
    return records


def fetch_benchmark(start_date):
    """从 akshare 获取沪深300指数数据"""
    print(f"  获取沪深300数据 (起始: {start_date})...")
    df = ak.stock_zh_index_daily(symbol="sh000300")
    df["date"] = df["date"].astype(str)
    df = df[df["date"] >= start_date].copy()
    df = df.sort_values("date").reset_index(drop=True)
    print(f"  沪深300数据: {len(df)} 条 ({df['date'].iloc[0]} ~ {df['date'].iloc[-1]})")
    return dict(zip(df["date"], df["close"]))


def align_data(fund_records, benchmark_dict):
    """对齐基金净值与基准数据，归一化到起始日为1.0"""
    # 构建基金日期->净值字典
    fund_dict = {d: v for d, v in fund_records}

    # 找到共同日期范围
    fund_dates = sorted(fund_dict.keys())
    start_date = fund_dates[0]

    # 获取基准起始日净值用于归一化
    bench_dates_sorted = sorted(d for d in benchmark_dict.keys() if d >= start_date)
    if not bench_dates_sorted:
        print("  警告: 没有找到匹配的基准数据")
        return None

    bench_start_val = benchmark_dict[bench_dates_sorted[0]]

    # 对齐：只保留基金有数据且基准也有数据的日期
    aligned_dates = []
    aligned_nav = []
    aligned_bench = []

    for d in fund_dates:
        if d in benchmark_dict:
            aligned_dates.append(d)
            aligned_nav.append(fund_dict[d])
            # 基准归一化：起始日为1.0
            aligned_bench.append(round(benchmark_dict[d] / bench_start_val, 4))

    # 基金净值也归一化到起始日为1.0
    nav_start = aligned_nav[0]
    aligned_nav = [round(v / nav_start, 4) for v in aligned_nav]

    print(f"  对齐后: {len(aligned_dates)} 条 ({aligned_dates[0]} ~ {aligned_dates[-1]})")
    return {
        "dates": aligned_dates,
        "nav": aligned_nav,
        "benchmark": aligned_bench,
    }


def update_jiazhi_html(data):
    """jiazhi 现在从 combined-data.json 加载数据，无需更新 HTML"""
    pass


def main():
    # 切换到项目根目录
    script_dir = os.path.dirname(os.path.abspath(__file__))
    project_root = os.path.dirname(script_dir)
    os.chdir(project_root)
    print(f"工作目录: {os.getcwd()}")

    # 获取沪深300基准数据
    benchmark_dict = fetch_benchmark(BENCHMARK_START)

    for key, config in PRODUCTS.items():
        print(f"\n{'='*50}")
        print(f"处理产品: {config['name']} ({key})")
        print(f"{'='*50}")

        excel_path = find_latest_excel(key)
        if not excel_path:
            print(f"  未找到 {key}_*.xlsx 文件，跳过")
            continue

        # 读取净值
        records = read_excel_nav(excel_path)
        if not records:
            print(f"  净值数据为空，跳过")
            continue

        # 对齐数据
        data = align_data(records, benchmark_dict)
        if not data:
            continue

        # 保存 JSON
        json_path = config["json_path"]
        os.makedirs(os.path.dirname(json_path), exist_ok=True)
        with open(json_path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        print(f"  已保存 {json_path}")

        # 更新 jiazhi HTML（如果是 jiazhi 产品）
        if key == "jiazhi":
            update_jiazhi_html(data)

    print(f"\n{'='*50}")
    print("更新完成！")
    print(f"{'='*50}")


if __name__ == "__main__":
    main()
