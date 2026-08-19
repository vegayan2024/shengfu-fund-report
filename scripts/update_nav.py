#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
净值数据更新脚本
读取 Excel 净值文件，对齐沪深300基准，生成 JSON 数据
"""

import os
import json
import xlrd
from datetime import datetime
from pathlib import Path

# 项目根目录
PROJECT_ROOT = Path(__file__).parent.parent
DATA_DIR = PROJECT_ROOT / 'data'

# 产品配置
PRODUCTS = {
    'zhouqi': {
        'name': '泽鑫周期',
        'pattern': 'zhouqi_*.xlsx',
        'json_path': PROJECT_ROOT / 'zhouqi' / 'combined-data.json',
    },
    'jiazhi': {
        'name': '泽鑫价值',
        'pattern': 'jiazhi_*.xlsx',
        'json_path': PROJECT_ROOT / 'jiazhi' / 'combined-data.json',
    }
}


def get_latest_excel(pattern):
    """获取最新的 Excel 文件"""
    files = list(DATA_DIR.glob(pattern))
    if not files:
        return None
    return max(files, key=lambda f: f.stat().st_mtime)


def read_excel_nav(filepath):
    """读取 Excel 净值数据"""
    wb = xlrd.open_workbook(str(filepath))
    sh = wb.sheet_by_index(0)

    records = []
    for row in range(sh.nrows):
        cell_date = sh.cell_value(row, 0)
        cell_nav = sh.cell_value(row, 1)

        # 处理日期
        if isinstance(cell_date, float):
            date = xlrd.xldate_as_datetime(cell_date, wb.datemode)
        else:
            date = datetime.strptime(str(cell_date).strip(), '%Y-%m-%d')

        # 处理净值
        nav = float(cell_nav)

        records.append((date.strftime('%Y-%m-%d'), nav))

    return records


def get_benchmark_data(start_date):
    """从 akshare 获取沪深300数据"""
    import akshare as ak

    print(f"  获取沪深300数据 (起始: {start_date})...")

    df = ak.stock_zh_index_daily(symbol="sh000300")
    df['date'] = df['date'].astype(str)
    df = df[df['date'] >= start_date]
    df = df.sort_values('date').reset_index(drop=True)

    benchmark = dict(zip(df['date'], df['close']))

    print(f"  沪深300数据: {len(benchmark)} 条 ({df['date'].iloc[0]} ~ {df['date'].iloc[-1]})")
    return benchmark


def align_data(records, benchmark_dict):
    """对齐基金数据和基准数据，归一化到起始日=1.0"""
    # 找到基金最早日期
    fund_start = records[0][0]

    # 过滤基准数据：从基金起始日开始
    benchmark_filtered = {d: v for d, v in benchmark_dict.items() if d >= fund_start}

    if not benchmark_filtered:
        print("  基准数据为空")
        return None

    # 获取基准起始日的值
    bench_dates_sorted = sorted(benchmark_filtered.keys())
    bench_start_val = benchmark_filtered[bench_dates_sorted[0]]

    # 对齐：只保留基金和基准都有数据的日期
    aligned_dates = []
    aligned_nav = []
    aligned_benchmark = []

    for date, nav in records:
        if date in benchmark_filtered:
            aligned_dates.append(date)
            aligned_nav.append(nav)
            # 基准归一化
            aligned_benchmark.append(benchmark_filtered[date] / bench_start_val)

    if not aligned_dates:
        print("  对齐后无数据")
        return None

    # 基金净值归一化
    nav_start = aligned_nav[0]
    aligned_nav = [v / nav_start for v in aligned_nav]

    print(f"  对齐数据: {len(aligned_dates)} 条 ({aligned_dates[0]} ~ {aligned_dates[-1]})")

    return {
        'dates': aligned_dates,
        'nav': aligned_nav,
        'benchmark': aligned_benchmark
    }


def update_jiazhi_html(data):
    """更新 jiazhi HTML 中的嵌入数据（保留兼容）"""
    pass


def process_clients():
    """处理客户信息表，生成 clients.json"""
    # 查找最新的客户信息表
    client_files = list(DATA_DIR.glob('客户信息*.xls')) + list(DATA_DIR.glob('客户信息*.xlsx'))

    if not client_files:
        print("  未找到客户信息表，跳过")
        return False

    # 使用最新的文件
    client_file = sorted(client_files)[-1]
    print(f"  处理客户信息表: {client_file.name}")

    try:
        wb = xlrd.open_workbook(str(client_file))
        sh = wb.sheet_by_index(0)

        clients = {}
        for r in range(1, sh.nrows):
            phone = str(int(sh.cell_value(r, 2)))
            last4 = phone[-4:]
            clients[last4] = {
                'name': sh.cell_value(r, 1),
                'product': sh.cell_value(r, 3),
                'date': sh.cell_value(r, 4)
            }

        json_path = PROJECT_ROOT / 'clients.json'
        with open(str(json_path), 'w', encoding='utf-8') as f:
            json.dump(clients, f, ensure_ascii=False, indent=2)

        print(f"  客户信息已更新: {json_path}")
        print(f"  客户数量: {len(clients)}")
        return True

    except Exception as e:
        print(f"  处理客户信息失败: {e}")
        return False


def main():
    """主函数：处理所有产品和客户信息"""
    print(f"工作目录: {PROJECT_ROOT}")

    # 获取基准数据
    benchmark_dict = get_benchmark_data('2024-07-01')

    for key, config in PRODUCTS.items():
        print(f"
{'='*50}")
        print(f"处理产品: {config['name']} ({key})")
        print(f"{'='*50}")

        # 获取最新 Excel
        excel_path = get_latest_excel(config['pattern'])
        if not excel_path:
            print(f"  未找到 {config['pattern']} 文件，跳过")
            continue

        print(f"  读取: {excel_path.name}")

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

    # 处理客户信息
    process_clients()

    print(f"
{'='*50}")
    print("更新完成！")
    print(f"{'='*50}")


if __name__ == "__main__":
    main()
