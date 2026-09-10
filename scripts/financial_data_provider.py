#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
晟孚投资统一金融数据调度层 (Financial Data Provider)
严格执行降级链条优先级规范 (Rule 5):
同花顺 (THS: 扶摇 REST / iFinD) > TuShare Pro > AKShare > 本地离线备份
"""

import os
import sys
import json
import logging
from datetime import datetime

logging.basicConfig(level=logging.INFO, format='%(asctime)s [%(levelname)s] %(message)s')
logger = logging.getLogger(__name__)


class FinancialDataProvider:
    def __init__(self):
        self.ths_available = False
        self.tushare_available = False
        self.akshare_available = False
        self._init_sources()

    def _init_sources(self):
        # 1. 尝试同花顺 iFinD / 扶摇
        try:
            import iFinDPy
            logger.info("同花顺 iFinD 模块已检测")
            self.ths_available = True
        except ImportError:
            self.ths_available = False

        # 2. 尝试 TuShare Pro
        try:
            import tushare as ts
            token = os.environ.get('TUSHARE_TOKEN', '')
            if token:
                ts.set_token(token)
                self.ts_pro = ts.pro_api()
                self.tushare_available = True
                logger.info("TuShare Pro 接口已就绪")
            else:
                self.tushare_available = False
        except ImportError:
            self.tushare_available = False

        # 3. 尝试 AKShare
        try:
            import akshare as ak
            self.ak = ak
            self.akshare_available = True
            logger.info("AKShare 接口已就绪")
        except ImportError:
            self.akshare_available = False

    def get_hs300_daily(self, start_date='2024-07-01'):
        """
        获取沪深300指数日度收盘行情
        返回: dict { 'YYYY-MM-DD': float(close) }
        优先级: THS > TuShare Pro > AKShare
        """
        # --- Level 1: 同花顺 iFinD ---
        if self.ths_available:
            try:
                import iFinDPy
                logger.info("【Priority 1】正在通过同花顺 (THS/iFinD) 拉取沪深300指数行情...")
                ths_login = iFinDPy.THS_iFinDLogin()
                if ths_login == 0:
                    today_str = datetime.now().strftime('%Y-%m-%d')
                    data = iFinDPy.THS_HistoryQuotes('000300.SH', 'close', '', start_date, today_str)
                    if hasattr(data, 'data'):
                        res = {}
                        for _, row in data.data.iterrows():
                            res[str(row['time'])[:10]] = float(row['close'])
                        logger.info(f"同花顺拉取成功: 共 {len(res)} 条记录")
                        return res
            except Exception as e:
                logger.warning(f"同花顺获取失败，降级至下一级: {e}")

        # --- Level 2: TuShare Pro ---
        if self.tushare_available:
            try:
                logger.info("【Priority 2】正在通过 TuShare Pro 拉取沪深300指数行情...")
                start_ts = start_date.replace('-', '')
                df = self.ts_pro.index_daily(ts_code='000300.SH', start_date=start_ts)
                if not df.empty:
                    df['date'] = df['trade_date'].apply(lambda x: f"{x[:4]}-{x[4:6]}-{x[6:]}")
                    df = df.sort_values('date')
                    res = dict(zip(df['date'], df['close'].astype(float)))
                    logger.info(f"TuShare Pro 拉取成功: 共 {len(res)} 条记录")
                    return res
            except Exception as e:
                logger.warning(f"TuShare Pro 获取失败，降级至下一级: {e}")

        # --- Level 3: AKShare ---
        if self.akshare_available:
            try:
                logger.info("【Priority 3】正在通过 AKShare 拉取沪深300 (sh000300) 指数行情...")
                df = self.ak.stock_zh_index_daily(symbol="sh000300")
                df['date'] = df['date'].astype(str)
                df = df[df['date'] >= start_date]
                df = df.sort_values('date').reset_index(drop=True)
                res = dict(zip(df['date'], df['close'].astype(float)))
                logger.info(f"AKShare 拉取成功: 共 {len(res)} 条记录 ({df['date'].iloc[0]} ~ {df['date'].iloc[-1]})")
                return res
            except Exception as e:
                logger.error(f"AKShare 获取失败: {e}")

        raise RuntimeError("所有金融数据源均无法获取沪深300数据，请检查网络或授权配置。")


provider = FinancialDataProvider()

if __name__ == '__main__':
    sys.stdout.reconfigure(encoding='utf-8')
    data = provider.get_hs300_daily('2026-05-26')
    dates = sorted(data.keys())
    print(f"获取样本: 起始 {dates[0]} = {data[dates[0]]}, 截止 {dates[-1]} = {data[dates[-1]]}")
