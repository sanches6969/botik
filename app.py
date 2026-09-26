#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Breakout Bot v9.6.2 - LIMIT ENTRY + HF WEB SERVER
- Добавлен HTTP-сервер на порту 7860 (требование HF Spaces)
- Бот работает в отдельном потоке
- Веб-сервер отвечает статусом (для мониторинга)
"""

import ccxt
import pandas as pd
import numpy as np
from datetime import datetime
import time
import os
import json
import math
import warnings
import openpyxl
import requests
import threading
from http.server import HTTPServer, BaseHTTPRequestHandler
from colorama import init, Fore
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from collections import defaultdict
from typing import Dict, List, Optional

warnings.filterwarnings('ignore')
init(autoreset=True)


# =========================================================
# 🌐 HF WEB SERVER (обязателен для Hugging Face Spaces)
# =========================================================
BOT_INSTANCE = {'bot': None}


class HealthCheckHandler(BaseHTTPRequestHandler):
    def do_GET(self):
        try:
            bot = BOT_INSTANCE.get('bot')
            if bot is None:
                html = "<h1>🤖 Bot starting...</h1>"
            else:
                stats = bot.calculate_statistics()
                filled_margin = sum(
                    p.get('required_margin', 0) for p in bot.positions.values() if p.get('entry_filled')
                )
                real_balance = bot.balance + filled_margin

                positions_html = ""
                for sym, pos in bot.positions.items():
                    pnl = pos.get('pnl', 0)
                    color = "green" if pnl >= 0 else "red"
                    sign = "+" if pnl >= 0 else ""
                    filled = "✅" if pos.get('entry_filled') else "⏳"
                    tp1 = "🎯" if pos.get('tp1_hit') else ""
                    tfs = "+".join(pos.get('tfs', []))
                    positions_html += (
                        f"<li>{filled}{tp1} <b>{sym}</b> {pos['type']} [{tfs}] "
                        f"<span style='color:{color}'>{sign}${pnl:.2f} ({sign}{pos.get('pnl_percent', 0):.2f}%)</span></li>"
                    )

                if not positions_html:
                    positions_html = "<li>Нет открытых позиций</li>"

                unrealized = sum(p.get('pnl', 0) for p in bot.positions.values() if p.get('entry_filled'))

                html = f"""
                <html><head><meta charset="utf-8"><title>Trading Bot</title>
                <style>body{{font-family:monospace;padding:20px;background:#1a1a1a;color:#eee}}
                h1{{color:#0ff}}h2{{color:#ff0;margin-top:20px}}
                .green{{color:#0f0}}.red{{color:#f00}}</style></head>
                <body>
                <h1>🤖 BREAKOUT BOT v9.6.2</h1>
                <h2>💰 Баланс</h2>
                <p>Реальный капитал: <b>${real_balance:,.2f}</b></p>
                <p>Свободно: ${bot.balance:,.2f} | В позициях: ${filled_margin:,.2f}</p>
                <p>Реализ. PnL: <b>{'+' if stats['total_pnl'] >= 0 else ''}${stats['total_pnl']:.2f}</b></p>
                <p>Комиссии: ${bot.total_fees:.2f}</p>
                <h2>📊 Статистика</h2>
                <p>Сделок: {stats['total_trades']} | Выиграно: {stats['win_trades']} ({stats['win_rate']:.1f}%)</p>
                <p>PF: {stats['profit_factor']:.2f}</p>
                <h2>📂 Открытые позиции ({len(bot.positions)})</h2>
                <ul>{positions_html}</ul>
                <p>Нереализ. PnL: <b>{'+' if unrealized >= 0 else ''}${unrealized:.2f}</b></p>
                <p style="color:#888;margin-top:30px">
                    Обновлено: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
                </p>
                </body></html>
                """
            self.send_response(200)
            self.send_header('Content-Type', 'text/html; charset=utf-8')
            self.end_headers()
            self.wfile.write(html.encode('utf-8'))
        except Exception as e:
            self.send_response(500)
            self.end_headers()
            self.wfile.write(f"Error: {e}".encode())

    def log_message(self, format, *args):
        pass  # отключаем спам-логи HTTP


def start_web_server(port=7860):
    """Запускает HTTP-сервер в фоновом режиме"""
    server = HTTPServer(('0.0.0.0', port), HealthCheckHandler)
    print(f"🌐 Web server: http://0.0.0.0:{port}")
    server.serve_forever()


# =========================================================
# TELEGRAM
# =========================================================
TELEGRAM_TOKEN = "8250935517:AAGrJ8H9lyuTwzBoimWZUrVHViFNpopYZTM"
TELEGRAM_CHAT_ID = "511975317"


class TelegramNotifier:
    def __init__(self, token: str, chat_id: str):
        self.token = token
        self.chat_id = chat_id
        self.enabled = bool(token and chat_id)

    def send(self, text: str):
        if not self.enabled:
            return
        try:
            url = f"https://api.telegram.org/bot{self.token}/sendMessage"
            requests.post(url, json={
                'chat_id': self.chat_id,
                'text': text,
                'parse_mode': 'HTML'
            }, timeout=10)
        except Exception as e:
            print(f"⚠️ Telegram: {e}")

    def notify_open(self, trade: Dict, margin: float, positions: Dict):
        tfs = "+".join(trade.get('tfs', []))
        weight = trade.get('weight', 0)
        self.send(
            f"📈 <b>ЛИМИТНЫЙ ОРДЕР</b>\n"
            f"Монета: <b>{trade['symbol']}</b> | {trade['type']}\n"
            f"🎯 {tfs} | Вес: {weight:.1f}\n"
            f"Вход: {trade['entry_price']:.6f}\n"
            f"Стоп: {trade['sl']:.6f}\n"
            f"Тейк: {trade['tp']:.6f}\n"
            f"Резерв: ${margin:.2f}"
        )

    def notify_tp1(self, symbol: str, pnl: float):
        self.send(f"✅ <b>TP1 (50%)</b> {symbol}\nPnL: <b>+${pnl:.2f}</b>\n🛡 Стоп в БУ")

    def notify_close(self, trade: Dict, balance: float, total_pnl: float,
                     total_trades: int, wins: int, losses: int, pf: float,
                     total_fees: float = 0.0):
        emoji = "🟢" if trade['pnl'] >= 0 else "🔴"
        sign = "+" if trade['pnl'] >= 0 else ""
        winrate = (wins / total_trades * 100) if total_trades else 0
        self.send(
            f"{emoji} <b>ЗАКРЫТА</b> {trade['symbol']}\n"
            f"PnL: <b>{sign}${trade['pnl']:.2f}</b> ({sign}{trade['pnl_percent']:.2f}%)\n"
            f"Причина: {trade['reason']}\n"
            f"━━━━━━━━━━━━━━━━\n"
            f"💰 Баланс: ${balance:,.2f}\n"
            f"📈 PnL: {sign}${total_pnl:.2f}\n"
            f"💸 Комиссии: ${total_fees:.2f}\n"
            f"📋 Сделок: {total_trades} | ✅ {wins} ({winrate:.1f}%)\n"
            f"⚖️ PF: {pf:.2f}"
        )

    def send_hourly(self, balance, total_pnl, total_trades, wins, losses, pf,
                    active, initial, total_fees: float = 0.0):
        winrate = (wins / total_trades * 100) if total_trades else 0
        pnl_pct = (total_pnl / initial * 100) if initial else 0
        self.send(
            f"⏰ <b>ОТЧЁТ</b>\n"
            f"💰 Баланс: ${balance:,.2f}\n"
            f"📈 PnL: {'+' if total_pnl >= 0 else ''}${total_pnl:.2f} ({pnl_pct:.2f}%)\n"
            f"💸 Комиссии: ${total_fees:.2f}\n"
            f"📋 Сделок: {total_trades} | ✅ {wins} ({winrate:.1f}%)\n"
            f"⚖️ PF: {pf:.2f}\n"
            f"🔄 Активных: {active}"
        )


# =========================================================
# EXCEL
# =========================================================
class ExcelLogger:
    def __init__(self, initial_balance: float):
        self.file = 'trading_report.xlsx'
        self.initial_balance = initial_balance

    def log_trade(self, trade: Dict, all_trades: List[Dict], active_positions: int, total_fees: float = 0.0):
        try:
            if os.path.exists(self.file):
                wb = openpyxl.load_workbook(self.file)
            else:
                wb = openpyxl.Workbook()
                wb.remove(wb.active)

            headers = ['#', 'Дата', 'Монета', 'Тип', 'Вход', 'Выход',
                       'PnL ($)', 'PnL (%)', 'Комиссия', 'Причина', 'Сектор', 'Таймфреймы', 'Вес']

            if 'Trades' in wb.sheetnames:
                ws = wb['Trades']
            else:
                ws = wb.create_sheet('Trades', 0)
                ws.append(headers)
                for col in range(1, len(headers) + 1):
                    c = ws.cell(row=1, column=col)
                    c.font = Font(bold=True, color='FFFFFF')
                    c.fill = PatternFill('solid', fgColor='2F5496')

            num = ws.max_row
            tfs = "+".join(trade.get('tfs', []))
            ws.append([
                num,
                datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                trade['symbol'], trade['type'],
                float(trade['entry']), float(trade['exit']),
                float(trade['pnl']), float(trade['pnl_percent']),
                float(trade.get('fee', 0)),
                trade['reason'], trade.get('sector', 'UNKNOWN'),
                tfs, float(trade.get('weight', 0))
            ])
            pnl_cell = ws.cell(row=ws.max_row, column=7)
            if float(trade['pnl']) >= 0:
                pnl_cell.fill = PatternFill('solid', fgColor='C6EFCE')
            else:
                pnl_cell.fill = PatternFill('solid', fgColor='FFC7CE')

            if 'Summary' in wb.sheetnames:
                del wb['Summary']
            s = wb.create_sheet('Summary', 0)
            s['A1'] = '📊 ИТОГОВАЯ СТАТИСТИКА'

            total = len(all_trades)
            wins = [t for t in all_trades if t['pnl'] > 0]
            losses = [t for t in all_trades if t['pnl'] <= 0]
            total_pnl = sum(t['pnl'] for t in all_trades)
            prof = sum(t['pnl'] for t in wins) if wins else 0
            los = sum(t['pnl'] for t in losses) if losses else 0

            rows = [
                ['Начальный капитал', f'${self.initial_balance:.2f}'],
                ['Реальный баланс', f'${self.initial_balance + total_pnl - total_fees:.2f}'],
                ['Чистый PnL', f'${total_pnl - total_fees:.2f}'],
                ['Комиссии', f'${total_fees:.2f}'],
                ['', ''],
                ['Всего сделок', total],
                ['✅ Прибыльных', f'{len(wins)} ({(len(wins)/total*100 if total else 0):.1f}%)'],
                ['❌ Убыточных', f'{len(losses)}'],
                ['Средний выигрыш', f'${(prof/len(wins) if wins else 0):.2f}'],
                ['Средний проигрыш', f'${(los/len(losses) if losses else 0):.2f}'],
                ['Profit Factor', f'{(abs(prof/los) if los else 0):.2f}'],
                ['Активных позиций', active_positions],
            ]
            for i, (k, v) in enumerate(rows, start=3):
                s.cell(row=i, column=1, value=k).font = Font(bold=True)
                s.cell(row=i, column=2, value=v)

            wb.save(self.file)
        except Exception as e:
            print(f"{Fore.RED}❌ Excel: {e}")


# =========================================================
# CONFIG
# =========================================================
class TradingConfig:
    INITIAL_BALANCE = 1000
    MAX_POSITIONS = 5
    SCAN_INTERVAL = 60
    CHECK_INTERVAL = 1
    LEVERAGE = 10

    FIXED_MARGIN = 200

    STOP_LOSS = 0.003
    MIN_SL_PERCENT = 0.003
    MAX_SL_PERCENT = 0.005
    ATR_SL_MULTIPLIER = 0.4

    BASE_RISK_PERCENT = 0.01
    MIN_RISK_PERCENT = 0.002
    MAX_RISK_PERCENT = 0.02

    MIN_VOLUME_USD = 500_000
    MIN_VOLUME_RATIO = 1.3
    TOUCH_COUNT = 2
    ZONE_WIDTH = 0.005
    MAX_VOLATILITY = 0.08

    PROFIT_TARGET = 0.015
    TRAILING_STOP = 0.01

    MAX_DRAWDOWN = 0.15
    MAX_EQUITY_DRAWDOWN = 0.20

    REENTRY_COOLDOWN = 2 * 3600
    COOLDOWN_STREAK_MULTIPLIER = 3
    LOSS_STREAK_FOR_LONG_COOLDOWN = 2

    ORDER_TIMEOUT_SEC = 600

    MAKER_FEE = 0.0
    TAKER_FEE = 0.001

    MAX_OTHER_SECTOR_POSITIONS = 3
    MAX_TOKEN_QUANTITY = 10_000_000

    EXCLUDED_SECTORS = ['MEME']

    TIMEFRAMES = [
        {'tf': '1d', 'limit': 100, 'weight': 3.0, 'label': 'D1'},
        {'tf': '4h', 'limit': 100, 'weight': 2.0, 'label': '4H'},
        {'tf': '1h', 'limit': 100, 'weight': 1.0, 'label': '1H'},
    ]
    LEVEL_CLUSTER_DISTANCE = 0.01
    MIN_LEVEL_WEIGHT = 3.0
    ROUND_NUMBER_BONUS = 0.5

    SECTORS = {
        'DEFI': ['UNI', 'AAVE', 'MKR', 'COMP', 'CRV', 'SUSHI', 'BAL'],
        'LAYER1': ['BTC', 'ETH', 'SOL', 'AVAX', 'NEAR', 'DOT', 'ADA'],
        'LAYER2': ['ARB', 'OP', 'MATIC', 'MNT', 'STRK', 'METIS'],
        'GAMING': ['SAND', 'MANA', 'AXS', 'GALA', 'ENJ', 'ILV'],
        'AI': ['FET', 'AGIX', 'OCEAN', 'RLC', 'NMR'],
        'MEME': ['DOGE', 'SHIB', 'PEPE', 'BONK', 'WIF', 'FLOKI', 'TRUMP',
                 'BOME', 'BRETT', 'POPCAT', 'MOG', 'TURBO'],
        'INFRA': ['LINK', 'GRT', 'THETA', 'LPT'],
        'DEX': ['CAKE', '1INCH', 'DYDX']
    }

    BLACKLIST = [
        'USDC', 'USDT', 'DAI', 'BUSD', 'TUSD', 'USDP', 'FDUSD',
        'RLUSD', 'GUSD', 'PAX', 'HUSD', 'USDN', 'USTC',
        'LUNC', 'UST', 'MIM', 'FRAX', 'EURT', 'XAUT',
        'TUT', 'BOM', 'SHI', 'XEC',
        'DGAI', 'GRVT', 'AKE', 'POWER', 'MODA',
        'DOGE', 'SHIB', 'PEPE', 'BONK', 'WIF', 'FLOKI', 'TRUMP',
        'BOME', 'BRETT', 'POPCAT', 'MOG', 'TURBO',
        'BULLA', 'ZAMA', 'MEY', 'UAI', 'CC', 'STONK', 'AERO',
        'PUMP', 'WXT', 'DOS', 'MARSCOIN', 'ULTIMA', 'ZIG',
        'USELESS', 'CASHCAT', 'HYPE', 'ZRO', 'VTHO', 'REZ',
        'CRCLON', 'NPC', 'ZEN', 'BTW', 'CNPY', 'BSB', 'MUBARAK',
    ]


# =========================================================
# RISK MANAGER
# =========================================================
class RiskManager:
    def __init__(self, initial_balance: float, config: TradingConfig):
        self.initial_balance = initial_balance
        self.config = config
        self.trade_history = []

    def update_kelly(self, trades):
        if len(trades) < 10:
            return None
        w = [t for t in trades if t.get('pnl', 0) > 0]
        l = [t for t in trades if t.get('pnl', 0) <= 0]
        if not w or not l:
            return None
        wr = len(w) / len(trades)
        aw = sum(t['pnl'] for t in w) / len(w)
        al = abs(sum(t['pnl'] for t in l) / len(l))
        if al == 0:
            return None
        kelly = (wr * aw - (1 - wr) * al) / (aw * wr)
        return min(max(kelly * 0.5, self.config.MIN_RISK_PERCENT), self.config.MAX_RISK_PERCENT)

    def calculate_position_size(self, entry_price, sl_percent, available_balance):
        margin = self.config.FIXED_MARGIN
        if available_balance < margin:
            margin = max(available_balance * 0.95, 0)
        if margin <= 0:
            return 0, 0
        quantity = (margin * self.config.LEVERAGE) / entry_price
        return max(quantity, 0), margin


# =========================================================
# TECH ANALYZER
# =========================================================
class TechnicalAnalyzer:
    def __init__(self, exchange, config):
        self.exchange = exchange
        self.config = config

    def get_klines(self, symbol, timeframe='1h', limit=100):
        try:
            ohlcv = self.exchange.fetch_ohlcv(symbol, timeframe, limit=limit)
            df = pd.DataFrame(ohlcv, columns=['timestamp', 'open', 'high', 'low', 'close', 'volume'])
            df['timestamp'] = pd.to_datetime(df['timestamp'], unit='ms')
            return df
        except Exception:
            return None

    def calculate_rsi(self, prices, period=14):
        if len(prices) < period + 1:
            return 50
        deltas = np.diff(prices)
        gains = np.where(deltas > 0, deltas, 0)
        losses = np.where(deltas < 0, -deltas, 0)
        avg_gain = pd.Series(gains).rolling(period).mean().iloc[-1]
        avg_loss = pd.Series(losses).rolling(period).mean().iloc[-1]
        if avg_loss == 0:
            return 100
        rs = avg_gain / avg_loss
        return 100 - (100 / (1 + rs))

    def calculate_atr(self, df, period=14):
        if df is None or len(df) < period + 1:
            return 0
        high = df['high'].values
        low = df['low'].values
        close = df['close'].values
        tr = np.maximum(high[1:] - low[1:],
                        np.maximum(abs(high[1:] - close[:-1]), abs(low[1:] - close[:-1])))
        return float(np.mean(tr[-period:]))

    def _is_round_number(self, price):
        if price <= 0:
            return False
        try:
            magnitude = 10 ** math.floor(math.log10(price))
            for factor in [0.5, 1.0, 2.0, 5.0, 10.0]:
                if abs(price - magnitude * factor) / price < 0.005:
                    return True
        except:
            pass
        return False

    def _find_levels_single_tf(self, df, touch_count, tf_weight, tf_label):
        if df is None or len(df) < 30:
            return []
        highs = df['high'].values
        lows = df['low'].values
        closes = df['close'].values
        volumes = df['volume'].values
        last_price = closes[-1]
        levels = []
        window = 5
        avg_vol = np.mean(volumes[-20:]) if len(volumes) >= 20 else np.mean(volumes)

        for i in range(window, len(df) - window):
            if highs[i] == max(highs[i-window:i+window+1]):
                level = highs[i]
                touches = sum(1 for j in range(len(highs)) if abs(highs[j] - level) / level < 0.008)
                if touches >= touch_count:
                    distance = abs(last_price - level) / last_price
                    if distance < 0.15:
                        vol_at = sum(volumes[j] for j in range(len(highs)) if abs(highs[j] - level) / level < 0.008)
                        vol_strength = min(vol_at / (avg_vol * touches), 3.0) if avg_vol > 0 else 1.0
                        round_bonus = self.config.ROUND_NUMBER_BONUS if self._is_round_number(level) else 0
                        weight = tf_weight * (1 + touches * 0.3) * vol_strength + round_bonus
                        levels.append({'price': level, 'touches': touches, 'distance': distance,
                                       'weight': weight, 'tf': tf_label, 'type': 'resistance', 'vol_strength': vol_strength})
            if lows[i] == min(lows[i-window:i+window+1]):
                level = lows[i]
                touches = sum(1 for j in range(len(lows)) if abs(lows[j] - level) / level < 0.008)
                if touches >= touch_count:
                    distance = abs(last_price - level) / last_price
                    if distance < 0.15:
                        vol_at = sum(volumes[j] for j in range(len(lows)) if abs(lows[j] - level) / level < 0.008)
                        vol_strength = min(vol_at / (avg_vol * touches), 3.0) if avg_vol > 0 else 1.0
                        round_bonus = self.config.ROUND_NUMBER_BONUS if self._is_round_number(level) else 0
                        weight = tf_weight * (1 + touches * 0.3) * vol_strength + round_bonus
                        levels.append({'price': level, 'touches': touches, 'distance': distance,
                                       'weight': weight, 'tf': tf_label, 'type': 'support', 'vol_strength': vol_strength})
        return levels

    def _cluster_levels(self, levels, cluster_distance):
        if not levels:
            return []
        sorted_levels = sorted(levels, key=lambda x: x['price'])
        clusters = []
        current = [sorted_levels[0]]
        for level in sorted_levels[1:]:
            if abs(level['price'] - current[-1]['price']) / current[-1]['price'] < cluster_distance:
                current.append(level)
            else:
                clusters.append(current)
                current = [level]
        clusters.append(current)

        result = []
        for cluster in clusters:
            tw = sum(l['weight'] for l in cluster)
            if tw == 0:
                continue
            avg = sum(l['price'] * l['weight'] for l in cluster) / tw
            tfs = list(set(l['tf'] for l in cluster))
            types = [l['type'] for l in cluster]
            result.append({
                'price': avg,
                'touches': sum(l['touches'] for l in cluster),
                'distance': min(l['distance'] for l in cluster),
                'weight': tw + len(tfs) * 1.5,
                'tfs': sorted(tfs),
                'type': max(set(types), key=types.count),
                'vol_strength': max(l['vol_strength'] for l in cluster)
            })
        return result

    def find_levels_multitimeframe(self, symbol):
        all_levels = []
        for tf_config in self.config.TIMEFRAMES:
            df = self.get_klines(symbol, tf_config['tf'], tf_config['limit'])
            if df is None or len(df) < 30:
                continue
            all_levels.extend(self._find_levels_single_tf(
                df, self.config.TOUCH_COUNT, tf_config['weight'], tf_config['label']))
        if not all_levels:
            return {'support': [], 'resistance': []}
        clustered = self._cluster_levels(all_levels, self.config.LEVEL_CLUSTER_DISTANCE)
        levels = {'support': [], 'resistance': []}
        for c in clustered:
            if c['type'] == 'support':
                levels['support'].append(c)
            else:
                levels['resistance'].append(c)
        for k in levels:
            levels[k] = sorted(levels[k], key=lambda x: x['weight'], reverse=True)
        return levels

    def check_breakout_quality(self, df, level_price, direction, config):
        if len(df) < 30:
            return False
        vol_ma = df['volume'].rolling(20).mean().iloc[-1]
        if df['volume'].iloc[-1] < vol_ma * config.MIN_VOLUME_RATIO:
            return False
        last = df.iloc[-1]
        if direction == 'LONG' and last['close'] <= level_price * (1 + config.ZONE_WIDTH * 0.5):
            return False
        if direction == 'SHORT' and last['close'] >= level_price * (1 - config.ZONE_WIDTH * 0.5):
            return False
        rsi = self.calculate_rsi(df['close'].values)
        if direction == 'LONG' and rsi > 80:
            return False
        if direction == 'SHORT' and rsi < 20:
            return False
        ma20 = df['close'].iloc[-20:].mean()
        ma50 = df['close'].iloc[-50:].mean()
        if direction == 'LONG' and ma20 < ma50 * 0.98:
            return False
        if direction == 'SHORT' and ma20 > ma50 * 1.02:
            return False
        atr = self.calculate_atr(df)
        if atr > 0 and abs(last['close'] - last['open']) < atr * 0.3:
            return False
        return True


# =========================================================
# BOT
# =========================================================
class BreakoutBotV7:
    def __init__(self, initial_balance: float = 1000):
        self.config = TradingConfig()
        self.initial_balance = initial_balance
        self.balance = initial_balance
        self.positions = {}
        self.closed_trades = []
        self.total_trades = 0
        self.win_trades = 0
        self.total_fees = 0.0
        self.blacklist_runtime = set()

        self.risk_manager = RiskManager(initial_balance, self.config)

        mexc = ccxt.mexc({'enableRateLimit': True, 'options': {'defaultType': 'spot'}})
        self.analyzer = TechnicalAnalyzer(mexc, self.config)
        self.exchange = mexc

        self.tg = TelegramNotifier(TELEGRAM_TOKEN, TELEGRAM_CHAT_ID)
        self.logger = ExcelLogger(initial_balance)
        self.last_hourly_report = time.time()

        self.reentry_cooldown_until = {}
        self.loss_streak = defaultdict(int)
        self.sector_positions = defaultdict(int)
        self.emergency_stop_activated = False
        self.is_running = False

        self.log_file = 'trading_log_v9.csv'
        self.state_file = 'bot_state.json'

        self._init_log_files()
        self._load_state()
        self._print_welcome()

        if self.tg.enabled:
            self.tg.send(f"🚀 <b>BOT v9.6.2 запущен!</b>\n"
                         f"💰 Баланс: ${self.balance:.2f}\n"
                         f"📋 Сделок: {len(self.closed_trades)}\n"
                         f"📂 Открытых: {len(self.positions)}")

    def _init_log_files(self):
        if not os.path.exists(self.log_file):
            with open(self.log_file, 'w', encoding='utf-8') as f:
                f.write("timestamp,symbol,side,entry,exit,quantity,leverage,pnl,pnl_percent,fee,reason,sector,tfs,weight\n")

    def _load_state(self):
        if not os.path.exists(self.state_file):
            print(f"{Fore.CYAN}🆕 Первый запуск — ${self.initial_balance:.2f}")
            return
        try:
            with open(self.state_file, 'r', encoding='utf-8') as f:
                state = json.load(f)
            self.balance = state.get('balance', self.initial_balance)
            self.total_fees = state.get('total_fees', 0.0)
            self.win_trades = state.get('win_trades', 0)
            self.closed_trades = state.get('closed_trades', [])
            for sym, pos in state.get('positions', {}).items():
                if all(k in pos for k in ['symbol', 'type', 'entry_price', 'quantity']):
                    pos.setdefault('trailing_distance', self.config.TRAILING_STOP)
                    pos.setdefault('tfs', [])
                    pos.setdefault('weight', 0)
                    self.positions[sym] = pos
                    self.sector_positions[pos.get('sector', 'OTHER')] += 1
            self.reentry_cooldown_until = state.get('reentry_cooldown_until', {})
            self.loss_streak = defaultdict(int, state.get('loss_streak', {}))
            print(f"{Fore.GREEN}✅ Загружено: ${self.balance:.2f} | сделок {len(self.closed_trades)} | позиций {len(self.positions)}")
        except Exception as e:
            print(f"{Fore.RED}⚠️ {e}")

    def _save_state(self):
        try:
            state = {
                'version': '9.6.2',
                'timestamp': datetime.now().isoformat(),
                'balance': self.balance,
                'total_fees': self.total_fees,
                'win_trades': self.win_trades,
                'closed_trades': self.closed_trades,
                'positions': {s: dict(p) for s, p in self.positions.items()},
                'reentry_cooldown_until': self.reentry_cooldown_until,
                'loss_streak': dict(self.loss_streak),
                'initial_balance': self.initial_balance,
            }
            tmp = self.state_file + '.tmp'
            with open(tmp, 'w', encoding='utf-8') as f:
                json.dump(state, f, indent=2, default=str)
            os.replace(tmp, self.state_file)
        except Exception as e:
            print(f"{Fore.RED}⚠️ Save: {e}")

    def _print_welcome(self):
        print(f"\n{Fore.GREEN}{'='*60}")
        print(f"{Fore.CYAN}🚀 BREAKOUT BOT v9.6.2 - LIMIT ENTRY + HF")
        print(f"{Fore.GREEN}{'='*60}")
        print(f"{Fore.YELLOW}💰 Баланс: ${self.balance:,.2f}")
        print(f"{Fore.MAGENTA}🔱 Плечо: {self.config.LEVERAGE}x | Маржа: ${self.config.FIXED_MARGIN}")
        print(f"{Fore.CYAN}🎯 Анализ: D1+4H+1H | Min вес: {self.config.MIN_LEVEL_WEIGHT}")
        print(f"{Fore.BLUE}🎯 TP1: {self.config.PROFIT_TARGET*100:.1f}% | Trailing: {self.config.TRAILING_STOP*100:.1f}%")
        print(f"{Fore.GREEN}💰 Maker 0% (вход+TP1) | Taker 0.1% (стоп)")
        print(f"{Fore.GREEN}{'='*60}\n")

    def get_all_symbols(self):
        try:
            markets = self.exchange.load_markets()
            return [s for s in markets if s.endswith('/USDT') 
                    and not any(b in s.replace('/USDT','').upper() for b in self.config.BLACKLIST)
                    and s not in self.blacklist_runtime
                    and markets[s].get('spot', False)]
        except Exception as e:
            print(f"{Fore.RED}❌ {e}")
            return []

    def get_sector(self, symbol):
        base = symbol.replace('/USDT', '')
        for sector, coins in self.config.SECTORS.items():
            if base in coins:
                return sector
        return 'OTHER'

    def check_diversification(self, symbol):
        sector = self.get_sector(symbol)
        if sector in self.config.EXCLUDED_SECTORS:
            return False
        if sector == 'OTHER':
            return self.sector_positions.get('OTHER', 0) < self.config.MAX_OTHER_SECTOR_POSITIONS
        return self.sector_positions.get(sector, 0) < 2

    def is_reentry_blocked(self, symbol):
        until = self.reentry_cooldown_until.get(symbol)
        return until is not None and time.time() < until

    def find_signals(self):
        symbols = self.get_all_symbols()
        if not symbols:
            return []
        tickers = {}
        try:
            all_t = self.exchange.fetch_tickers()
            for s in symbols:
                if s in all_t and all_t[s].get('quoteVolume', 0) > self.config.MIN_VOLUME_USD:
                    tickers[s] = all_t[s]
        except Exception as e:
            print(f"{Fore.RED}❌ {e}")
            return []

        sorted_syms = sorted(tickers.keys(), key=lambda x: tickers[x].get('quoteVolume', 0), reverse=True)[:80]
        print(f"{Fore.CYAN}🔍 MTF-скан {len(sorted_syms)} монет...")
        signals = []
        for symbol in sorted_syms:
            if not self.check_diversification(symbol): continue
            if self.is_reentry_blocked(symbol): continue
            if symbol in self.positions: continue
            ticker = tickers[symbol]
            price = ticker.get('last')
            if not price or price <= 0: continue
            levels = self.analyzer.find_levels_multitimeframe(symbol)
            if not levels['support'] and not levels['resistance']: continue
            df = self.analyzer.get_klines(symbol, '1h', 100)
            if df is None or len(df) < 50: continue
            atr = self.analyzer.calculate_atr(df)
            if price > 0 and atr / price > self.config.MAX_VOLATILITY: continue
            td = {'price': price, 'volume_24h': ticker.get('quoteVolume', 0)}
            signals.extend(self._check_breakouts_mtf(df, levels, td, symbol, atr))
        signals.sort(key=lambda x: (x.get('weight', 0) * 0.5 + x.get('volume_ratio', 0) * 0.3 + (1 - x.get('distance', 1)) * 0.2), reverse=True)
        return signals

    def _adaptive_sl(self, atr, price):
        if atr and price:
            return min(max((atr / price) * self.config.ATR_SL_MULTIPLIER, self.config.MIN_SL_PERCENT), self.config.MAX_SL_PERCENT)
        return self.config.STOP_LOSS

    def _check_breakouts_mtf(self, df, levels, td, symbol, atr):
        price = td['price']
        if td['volume_24h'] < self.config.MIN_VOLUME_USD: return []
        sl_p = self._adaptive_sl(atr, price)
        signals = []
        for level in levels['resistance']:
            if level['weight'] < self.config.MIN_LEVEL_WEIGHT: continue
            p = level['price']
            if p * (1 - self.config.ZONE_WIDTH) <= price <= p * (1 + self.config.ZONE_WIDTH):
                if not self.analyzer.check_breakout_quality(df, p, 'LONG', self.config): continue
                signals.append({
                    'type': 'LONG', 'symbol': symbol,
                    'entry_price': p * 1.001, 'level_price': p,
                    'sl': p * (1 - sl_p), 'sl_percent': sl_p,
                    'tp': p * (1 + self.config.PROFIT_TARGET), 'tp_percent': self.config.PROFIT_TARGET,
                    'volume_ratio': df['volume'].iloc[-1] / df['volume'].iloc[-20:].mean(),
                    'touches': level['touches'], 'distance': level['distance'],
                    'weight': level['weight'], 'tfs': level['tfs'],
                    'strength': min(level['weight'] / 10.0, 1.0)
                })
        for level in levels['support']:
            if level['weight'] < self.config.MIN_LEVEL_WEIGHT: continue
            p = level['price']
            if p * (1 - self.config.ZONE_WIDTH) <= price <= p * (1 + self.config.ZONE_WIDTH):
                if not self.analyzer.check_breakout_quality(df, p, 'SHORT', self.config): continue
                signals.append({
                    'type': 'SHORT', 'symbol': symbol,
                    'entry_price': p * 0.999, 'level_price': p,
                    'sl': p * (1 + sl_p), 'sl_percent': sl_p,
                    'tp': p * (1 - self.config.PROFIT_TARGET), 'tp_percent': self.config.PROFIT_TARGET,
                    'volume_ratio': df['volume'].iloc[-1] / df['volume'].iloc[-20:].mean(),
                    'touches': level['touches'], 'distance': level['distance'],
                    'weight': level['weight'], 'tfs': level['tfs'],
                    'strength': min(level['weight'] / 10.0, 1.0)
                })
        return signals

    def get_reserved_margin(self):
        return sum(p.get('required_margin', 0) for p in self.positions.values() if not p.get('entry_filled'))

    def open_position(self, signal):
        if len(self.positions) >= self.config.MAX_POSITIONS: return False
        symbol = signal['symbol']
        if self.emergency_stop_activated or self.is_reentry_blocked(symbol): return False
        if not self.check_diversification(symbol): return False
        entry = signal['entry_price']
        available = self.balance - self.get_reserved_margin()
        qty, _ = self.risk_manager.calculate_position_size(entry, signal['sl_percent'], available)
        if qty <= 0 or qty > self.config.MAX_TOKEN_QUANTITY: return False
        req_margin = qty * entry / self.config.LEVERAGE
        pos = {
            'symbol': symbol, 'type': signal['type'],
            'entry_price': entry, 'quantity': qty,
            'sl': signal['sl'], 'tp': signal['tp'],
            'sl_percent': signal['sl_percent'],
            'leverage': self.config.LEVERAGE,
            'level_price': signal['level_price'],
            'tfs': signal.get('tfs', []), 'weight': signal.get('weight', 0),
            'entry_time': datetime.now().isoformat(), 'created_at': time.time(),
            'required_margin': req_margin,
            'current_price': entry, 'pnl': 0, 'pnl_percent': 0,
            'entry_filled': False,
            'sector': self.get_sector(symbol),
            'tp1_hit': False,
            'trailing_high': entry, 'trailing_low': entry,
            'trailing_distance': self.config.TRAILING_STOP,
            'atr_entry': self.analyzer.calculate_atr(self.analyzer.get_klines(symbol, '5m', 20)),
            'max_profit_reached': 0
        }
        self.positions[symbol] = pos
        self.sector_positions[pos['sector']] += 1
        print(f"\n{Fore.GREEN}📈 ЛИМИТ: {symbol} {signal['type']} | {'+'.join(signal.get('tfs',[]))} | вес {signal.get('weight',0):.1f}")
        print(f"{Fore.CYAN}   Вход ${self._fp(entry)} | Залог ${req_margin:.2f} | maker 0%")
        self.tg.notify_open(pos, req_margin, self.positions)
        self._save_state()
        return True

    def _fp(self, p):
        if p is None: return "0"
        if p < 0.00001: return f"{p:.8f}"
        if p < 0.001: return f"{p:.6f}"
        if p < 1: return f"{p:.4f}"
        if p < 10: return f"{p:.3f}"
        return f"{p:.2f}"

    def check_positions(self):
        for symbol in list(self.positions.keys()):
            try:
                pos = self.positions.get(symbol)
                if pos is None: continue
                if not pos['entry_filled'] and time.time() - pos['created_at'] > self.config.ORDER_TIMEOUT_SEC:
                    print(f"{Fore.YELLOW}⏱️ Отмена {symbol}")
                    self._release_slot(symbol)
                    del self.positions[symbol]
                    self._save_state()
                    continue
                t = self.exchange.fetch_ticker(symbol)
                if t is None or t.get('last') is None: continue
                cp = t['last']
                if not pos['entry_filled']:
                    self._check_fill(symbol, cp); continue
                self._update_pnl(symbol, cp)
                self._check_exit(symbol, cp)
            except Exception as e:
                err = str(e)
                if "headers" in err or "not associated" in err:
                    self.blacklist_runtime.add(symbol)
                    try:
                        self._release_slot(symbol); del self.positions[symbol]; self._save_state()
                    except: pass
                else:
                    print(f"{Fore.RED}❌ {symbol}: {err[:80]}")

    def _release_slot(self, symbol):
        p = self.positions.get(symbol)
        if p and p.get('sector') in self.sector_positions:
            self.sector_positions[p['sector']] = max(0, self.sector_positions[p['sector']] - 1)

    def _check_fill(self, symbol, cp):
        pos = self.positions[symbol]
        filled = (pos['type'] == 'LONG' and cp >= pos['entry_price']) or \
                 (pos['type'] == 'SHORT' and cp <= pos['entry_price'])
        if not filled: return
        pos['entry_filled'] = True
        fee = pos['quantity'] * pos['entry_price'] * self.config.MAKER_FEE
        self.balance -= pos['required_margin'] + fee
        self.total_fees += fee
        print(f"{Fore.GREEN}✅ {symbol} (maker 0%)")
        self._save_state()

    def _update_pnl(self, symbol, cp):
        pos = self.positions[symbol]
        e = pos['entry_price']
        if pos['type'] == 'LONG':
            pos['pnl_percent'] = (cp - e) / e * 100
            pos['pnl'] = (cp - e) * pos['quantity']
        else:
            pos['pnl_percent'] = (e - cp) / e * 100
            pos['pnl'] = (e - cp) * pos['quantity']
        pos['current_price'] = cp
        if pos['pnl_percent'] > pos.get('max_profit_reached', 0):
            pos['max_profit_reached'] = pos['pnl_percent']

    def _check_exit(self, symbol, cp):
        pos = self.positions[symbol]
        if not pos['entry_filled']: return
        if not pos['tp1_hit'] and pos['pnl_percent'] >= self.config.PROFIT_TARGET * 100:
            self._tp1(symbol, cp); return
        if pos['tp1_hit'] and pos['quantity'] > 0:
            if self._trail(symbol, cp): return
        if pos['quantity'] > 0:
            if self._sl(symbol, cp): return

    def _tp1(self, symbol, cp):
        pos = self.positions[symbol]
        pos['tp1_hit'] = True
        q = pos['quantity'] * 0.5
        pnl = self._calc(pos, cp, q)
        fee = q * cp * self.config.MAKER_FEE
        self.balance += pnl + (q * cp / pos['leverage']) - fee
        self.total_fees += fee
        pos['quantity'] -= q
        pos['required_margin'] = pos['quantity'] * pos['entry_price'] / pos['leverage']
        pos['sl'] = pos['entry_price']
        rec = {'symbol': symbol + " (TP1)", 'type': pos['type'], 'entry': pos['entry_price'],
               'exit': cp, 'pnl': pnl - fee, 'pnl_percent': self.config.PROFIT_TARGET * 100,
               'leverage': pos['leverage'], 'reason': 'TP1', 'sector': pos.get('sector'),
               'entry_time': pos['entry_time'], 'quantity': q, 'fee': fee,
               'tfs': pos.get('tfs', []), 'weight': pos.get('weight', 0)}
        self.closed_trades.append(rec)
        if rec['pnl'] > 0: self.win_trades += 1
        self._log(rec); self.logger.log_trade(rec, self.closed_trades, len(self.positions), self.total_fees)
        pos['trailing_high'] = cp; pos['trailing_low'] = cp
        if pos.get('atr_entry', 0) > 0:
            pos['trailing_distance'] = max(self.config.TRAILING_STOP, (pos['atr_entry'] / cp) * 0.5)
        print(f"{Fore.GREEN}✅ TP1: {symbol} +${rec['pnl']:.2f} (maker 0%) | 🛡 стоп в БУ")
        self.tg.notify_tp1(symbol, rec['pnl'])
        self._save_state()

    def _trail(self, symbol, cp):
        pos = self.positions[symbol]
        if pos['type'] == 'LONG':
            if cp > pos['trailing_high']: pos['trailing_high'] = cp
            if cp <= pos['trailing_high'] * (1 - pos['trailing_distance']):
                self._close(symbol, cp, "Trailing Stop"); return True
        else:
            if cp < pos['trailing_low']: pos['trailing_low'] = cp
            if cp >= pos['trailing_low'] * (1 + pos['trailing_distance']):
                self._close(symbol, cp, "Trailing Stop"); return True
        return False

    def _sl(self, symbol, cp):
        pos = self.positions[symbol]
        if pos['type'] == 'LONG' and cp <= pos['sl']:
            self._close(symbol, cp, "Stop Loss"); return True
        if pos['type'] == 'SHORT' and cp >= pos['sl']:
            self._close(symbol, cp, "Stop Loss"); return True
        return False

    def _calc(self, pos, cp, q):
        e = pos['entry_price']
        return (cp - e) * q if pos['type'] == 'LONG' else (e - cp) * q

    def _update_reentry(self, symbol, pnl):
        if pnl > 0: self.loss_streak[symbol] = 0
        else: self.loss_streak[symbol] += 1
        cd = self.config.REENTRY_COOLDOWN
        if self.loss_streak[symbol] >= self.config.LOSS_STREAK_FOR_LONG_COOLDOWN:
            cd *= self.config.COOLDOWN_STREAK_MULTIPLIER
        self.reentry_cooldown_until[symbol] = time.time() + cd

    def _close(self, symbol, exit_price, reason):
        if symbol not in self.positions: return
        pos = self.positions[symbol]
        pnl = pos['pnl']
        q = pos['quantity']
        fee = q * exit_price * self.config.TAKER_FEE
        net = pnl - fee
        self.balance += pnl + (q * exit_price / pos['leverage']) - fee
        self.total_fees += fee
        rec = {'symbol': symbol, 'type': pos['type'], 'entry': pos['entry_price'],
               'exit': exit_price, 'pnl': net, 'pnl_percent': pos['pnl_percent'],
               'leverage': pos['leverage'], 'reason': reason, 'sector': pos.get('sector'),
               'entry_time': pos['entry_time'], 'quantity': q, 'fee': fee,
               'tfs': pos.get('tfs', []), 'weight': pos.get('weight', 0)}
        self.closed_trades.append(rec)
        if net > 0: self.win_trades += 1
        if len(self.closed_trades) >= 10:
            self.risk_manager.trade_history = self.closed_trades.copy()
        self._release_slot(symbol)
        self._update_reentry(symbol, net)
        self._log(rec); self.logger.log_trade(rec, self.closed_trades, len(self.positions) - 1, self.total_fees)
        c = Fore.GREEN if net >= 0 else Fore.RED
        s = "+" if net >= 0 else ""
        print(f"{c}🔚 {symbol} | {s}${net:.2f} ({pos['pnl_percent']:.2f}%) | {reason}")
        del self.positions[symbol]
        wins = [t for t in self.closed_trades if t['pnl'] > 0]
        losses = [t for t in self.closed_trades if t['pnl'] <= 0]
        prof = sum(t['pnl'] for t in wins) if wins else 0
        los = sum(t['pnl'] for t in losses) if losses else 0
        pf = abs(prof / los) if los else 0
        total = sum(t['pnl'] for t in self.closed_trades)
        fm = sum(p['required_margin'] for p in self.positions.values() if p['entry_filled'])
        rb = self.balance + fm
        self.tg.notify_close(rec, rb, total, len(self.closed_trades), len(wins), len(losses), pf, self.total_fees)
        self._save_state()

    def _log(self, t):
        with open(self.log_file, 'a', encoding='utf-8') as f:
            tfs = "+".join(t.get('tfs', []))
            f.write(f"{datetime.now()},{t['symbol']},{t['type']},{t['entry']:.8f},{t['exit']:.8f},"
                    f"{t.get('quantity',0):.2f},{t['leverage']},{t['pnl']:.2f},{t['pnl_percent']:.2f},"
                    f"{t.get('fee',0):.2f},{t['reason']},{t.get('sector','?')},{tfs},{t.get('weight',0):.2f}\n")

    def maybe_hourly(self):
        if not self.tg.enabled or time.time() - self.last_hourly_report < 3600: return
        self.last_hourly_report = time.time()
        w = [x for x in self.closed_trades if x['pnl'] > 0]
        l = [x for x in self.closed_trades if x['pnl'] <= 0]
        prof = sum(x['pnl'] for x in w) if w else 0
        los = sum(x['pnl'] for x in l) if l else 0
        pf = abs(prof / los) if los else 0
        total = sum(x['pnl'] for x in self.closed_trades)
        fm = sum(p['required_margin'] for p in self.positions.values() if p['entry_filled'])
        self.tg.send_hourly(self.balance + fm, total, len(self.closed_trades), len(w), len(l), pf, len(self.positions), self.initial_balance, self.total_fees)

    def check_emergency(self):
        if len(self.closed_trades) >= 5:
            total = sum(t['pnl'] for t in self.closed_trades)
            if total / self.initial_balance < -self.config.MAX_DRAWDOWN:
                self._emergency(f"Просадка: {total/self.initial_balance*100:.1f}%"); return True
        realized = sum(t['pnl'] for t in self.closed_trades)
        unrealized = sum(p.get('pnl', 0) for p in self.positions.values() if p['entry_filled'])
        if (realized + unrealized) / self.initial_balance < -self.config.MAX_EQUITY_DRAWDOWN:
            self._emergency(f"Просадка капитала"); return True
        return False

    def _emergency(self, reason):
        print(f"\n{Fore.RED}🚨 ЭКСТРЕННЫЙ СТОП! {reason}")
        for s in list(self.positions.keys()):
            try:
                p = self.positions[s]
                if not p['entry_filled']:
                    self._release_slot(s); del self.positions[s]; continue
                t = self.exchange.fetch_ticker(s)
                self._close(s, t['last'], "Emergency Stop")
            except: pass
        self.emergency_stop_activated = True
        self._save_state()
        if self.tg.enabled: self.tg.send(f"🚨 <b>ЭКСТРЕННЫЙ СТОП</b>\n{reason}")

    def stats(self):
        s = {'total_trades': len(self.closed_trades), 'win_trades': self.win_trades, 'win_rate': 0,
             'total_pnl': 0, 'avg_win': 0, 'avg_loss': 0, 'profit_factor': 0, 'total_fees': self.total_fees}
        if not self.closed_trades: return s
        df = pd.DataFrame(self.closed_trades)
        s['total_pnl'] = df['pnl'].sum()
        s['win_rate'] = (s['win_trades'] / s['total_trades']) * 100
        w = df[df['pnl'] > 0]; l = df[df['pnl'] <= 0]
        if not w.empty: s['avg_win'] = w['pnl'].mean()
        if not l.empty: s['avg_loss'] = l['pnl'].mean()
        if not w.empty and not l.empty:
            s['profit_factor'] = abs(w['pnl'].sum() / l['pnl'].sum())
        return s

    def print_summary(self):
        s = self.stats()
        fm = sum(p.get('required_margin', 0) for p in self.positions.values() if p['entry_filled'])
        ro = sum(p.get('required_margin', 0) for p in self.positions.values() if not p['entry_filled'])
        ur = sum(p.get('pnl', 0) for p in self.positions.values() if p['entry_filled'])
        rb = self.balance + fm
        print(f"\n{Fore.YELLOW}{'='*60}")
        print(f"📊 Свободно ${self.balance:,.2f} | В позициях ${fm:,.2f} | Резерв ${ro:,.2f}")
        print(f"💰 Реальный капитал: ${rb:,.2f}")
        print(f"📈 PnL: ${s['total_pnl']:+,.2f} ({s['total_pnl']/self.initial_balance*100:+.1f}%)")
        print(f"💸 Комиссии: ${self.total_fees:,.2f}")
        print(f"📋 Сделок: {s['total_trades']} | Выиграно: {s['win_trades']} ({s['win_rate']:.1f}%)")
        print(f"Активных: {len(self.positions)}/{self.config.MAX_POSITIONS}")
        for sym, p in self.positions.items():
            pnl = p.get('pnl', 0); sign = "+" if pnl >= 0 else ""
            color = Fore.GREEN if pnl >= 0 else Fore.RED
            f = "✅" if p.get('entry_filled') else "⏳"
            tp1 = "🎯" if p.get('tp1_hit') else ""
            tfs = "+".join(p.get('tfs', []))
            print(f"   {f}{tp1} {color}{sym} {p['type']} [{tfs}] {sign}${pnl:.2f}{Fore.RESET}")
        print(f"Нереализ.: {ur:+.2f}")
        if s['total_trades'] > 0:
            print(f"Ср. выигрыш ${s['avg_win']:.2f} | Ср. проигрыш ${s['avg_loss']:.2f} | PF: {s['profit_factor']:.2f}")
        print(f"{'='*60}")

    def run(self):
        print(f"{Fore.GREEN}🚀 БОТ ЗАПУЩЕН")
        self.is_running = True
        last_scan = 0
        last_summary = 0
        last_save = 0
        try:
            while self.is_running:
                now = time.time()
                if self.check_emergency(): break
                if self.positions: self.check_positions()
                if now - last_scan >= self.config.SCAN_INTERVAL:
                    last_scan = now
                    print(f"\n{Fore.CYAN}🔍 СКАН... {datetime.now().strftime('%H:%M:%S')}")
                    signals = self.find_signals()
                    if signals:
                        print(f"📋 Найдено: {len(signals)}")
                        for s in signals[:5]:
                            print(f"   {s['symbol']} {s['type']} | {'+'.join(s.get('tfs',[]))} | вес {s.get('weight',0):.1f}")
                        opened = 0
                        for sig in signals:
                            if len(self.positions) >= self.config.MAX_POSITIONS: break
                            if sig['symbol'] not in self.positions and self.open_position(sig):
                                opened += 1
                        if opened: print(f"{Fore.GREEN}✅ Открыто: {opened}")
                    else:
                        print("ℹ️ Нет сигналов")
                if now - last_summary >= 30:
                    last_summary = now
                    self.print_summary()
                if now - last_save >= 300:
                    last_save = now
                    self._save_state()
                self.maybe_hourly()
                time.sleep(self.config.CHECK_INTERVAL)
        except KeyboardInterrupt:
            print(f"\n{Fore.RED}🛑 Остановка...")
        finally:
            self.is_running = False
            self.print_summary()
            self._save_state()
            if self.tg.enabled: self.tg.send("🛑 <b>Бот остановлен</b>")


def main():
    # 🌐 Запускаем web-сервер в отдельном потоке (обязательно для HF)
    web_thread = threading.Thread(target=start_web_server, args=(7860,), daemon=True)
    web_thread.start()
    time.sleep(1)  # даём серверу подняться

    # 🤖 Запускаем бота в главном потоке
    try:
        bot = BreakoutBotV7(initial_balance=1000)
        BOT_INSTANCE['bot'] = bot
        bot.run()
    except Exception as e:
        print(f"{Fore.RED}❌ Критическая: {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    main()