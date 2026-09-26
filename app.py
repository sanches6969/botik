#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Breakout Bot v9.7 - LIMIT ENTRY (Server Version)
- Чистая версия без Gradio/веб-сервера
- Готова для запуска на ClawCloud / VPS / любом Linux
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
from colorama import init, Fore
from openpyxl.styles import Font, PatternFill
from collections import defaultdict
from typing import Dict, List

warnings.filterwarnings('ignore')
init(autoreset=True)


# =========================================================
# TELEGRAM
# =========================================================
TELEGRAM_TOKEN = os.environ.get("TELEGRAM_TOKEN", "8250935517:AAGrJ8H9lyuTwzBoimWZUrVHViFNpopYZTM")
TELEGRAM_CHAT_ID = os.environ.get("TELEGRAM_CHAT_ID", "511975317")


class TelegramNotifier:
    def __init__(self, token, chat_id):
        self.token = token
        self.chat_id = chat_id
        self.enabled = bool(token and chat_id)

    def send(self, text):
        if not self.enabled:
            return
        try:
            requests.post(
                f"https://api.telegram.org/bot{self.token}/sendMessage",
                json={'chat_id': self.chat_id, 'text': text, 'parse_mode': 'HTML'},
                timeout=10
            )
        except Exception as e:
            print(f"⚠️ TG: {e}")

    def notify_open(self, t, margin):
        tfs = "+".join(t.get('tfs', []))
        self.send(
            f"📈 <b>ЛИМИТ</b> {t['symbol']} {t['type']}\n"
            f"🎯 {tfs} | вес {t.get('weight', 0):.1f}\n"
            f"Вход: {t['entry_price']:.6f}\n"
            f"Стоп: {t['sl']:.6f}\n"
            f"Тейк: {t['tp']:.6f}"
        )

    def notify_tp1(self, s, pnl):
        self.send(f"✅ <b>TP1 (50%)</b> {s}\nPnL: <b>+${pnl:.2f}</b>\n🛡 Стоп в БУ")

    def notify_close(self, t, bal, total, trades, w, l, pf, fees):
        e = "🟢" if t['pnl'] >= 0 else "🔴"
        sg = "+" if t['pnl'] >= 0 else ""
        self.send(
            f"{e} <b>ЗАКРЫТА</b> {t['symbol']}\n"
            f"PnL: {sg}${t['pnl']:.2f} ({sg}{t['pnl_percent']:.2f}%)\n"
            f"Причина: {t['reason']}\n"
            f"━━━━━━━━━━━━━━━━\n"
            f"💰 Баланс: ${bal:,.2f}\n"
            f"📈 PnL: {sg}${total:.2f}\n"
            f"💸 Комиссии: ${fees:.2f}\n"
            f"📋 Сделок: {trades} | ✅ {w}\n"
            f"⚖️ PF: {pf:.2f}"
        )

    def send_hourly(self, bal, total, trades, w, l, pf, active, ib, fees):
        self.send(
            f"⏰ <b>ОТЧЁТ</b>\n"
            f"💰 Баланс: ${bal:,.2f}\n"
            f"📈 PnL: {'+' if total >= 0 else ''}${total:.2f} ({total/ib*100:.2f}%)\n"
            f"💸 Комиссии: ${fees:.2f}\n"
            f"📋 Сделок: {trades} | ✅ {w}\n"
            f"⚖️ PF: {pf:.2f}\n"
            f"🔄 Активных: {active}"
        )


# =========================================================
# EXCEL
# =========================================================
class ExcelLogger:
    def __init__(self, initial_balance):
        self.file = 'trading_report.xlsx'
        self.ib = initial_balance

    def log(self, trade, all_trades, active_positions):
        try:
            if os.path.exists(self.file):
                wb = openpyxl.load_workbook(self.file)
            else:
                wb = openpyxl.Workbook()
                wb.remove(wb.active)

            if 'Trades' in wb.sheetnames:
                ws = wb['Trades']
            else:
                ws = wb.create_sheet('Trades', 0)
                headers = ['#', 'Дата', 'Монета', 'Тип', 'Вход', 'Выход',
                           'PnL ($)', 'PnL (%)', 'Fee', 'Причина', 'Сектор', 'TFs', 'Вес']
                ws.append(headers)
                for c in range(1, len(headers) + 1):
                    cc = ws.cell(row=1, column=c)
                    cc.font = Font(bold=True, color='FFFFFF')
                    cc.fill = PatternFill('solid', fgColor='2F5496')

            ws.append([
                ws.max_row,
                datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                trade['symbol'], trade['type'],
                float(trade['entry']), float(trade['exit']),
                float(trade['pnl']), float(trade['pnl_percent']),
                float(trade.get('fee', 0)),
                trade['reason'], trade.get('sector', '?'),
                "+".join(trade.get('tfs', [])),
                float(trade.get('weight', 0))
            ])

            pnl_cell = ws.cell(row=ws.max_row, column=7)
            if float(trade['pnl']) >= 0:
                pnl_cell.fill = PatternFill('solid', fgColor='C6EFCE')
                pnl_cell.font = Font(color='006100')
            else:
                pnl_cell.fill = PatternFill('solid', fgColor='FFC7CE')
                pnl_cell.font = Font(color='9C0006')

            if 'Summary' in wb.sheetnames:
                del wb['Summary']
            s = wb.create_sheet('Summary', 0)
            s['A1'] = '📊 ИТОГОВАЯ СТАТИСТИКА'
            s['A1'].font = Font(bold=True, size=14)

            total = len(all_trades)
            wins = [t for t in all_trades if t['pnl'] > 0]
            losses = [t for t in all_trades if t['pnl'] <= 0]
            total_pnl = sum(t['pnl'] for t in all_trades)
            prof = sum(t['pnl'] for t in wins) if wins else 0
            los = sum(t['pnl'] for t in losses) if losses else 0

            rows = [
                ['Начальный капитал', f'${self.ib:.2f}'],
                ['Реальный баланс', f'${self.ib + total_pnl:.2f}'],
                ['Общий PnL', f'${total_pnl:.2f}'],
                ['PnL %', f'{(total_pnl / self.ib * 100):.2f}%'],
                ['', ''],
                ['Всего сделок', total],
                ['✅ Прибыльных', f'{len(wins)} ({(len(wins) / total * 100 if total else 0):.1f}%)'],
                ['❌ Убыточных', f'{len(losses)}'],
                ['Средний выигрыш', f'${(prof / len(wins) if wins else 0):.2f}'],
                ['Средний проигрыш', f'${(los / len(losses) if losses else 0):.2f}'],
                ['Profit Factor', f'{(abs(prof / los) if los else 0):.2f}'],
                ['Активных позиций', active_positions],
            ]
            for i, (k, v) in enumerate(rows, start=3):
                s.cell(row=i, column=1, value=k).font = Font(bold=True)
                s.cell(row=i, column=2, value=v)
            s.column_dimensions['A'].width = 25
            s.column_dimensions['B'].width = 25

            wb.save(self.file)
        except Exception as e:
            print(f"❌ Excel: {e}")


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
        'MEME': ['DOGE', 'SHIB', 'PEPE', 'BONK', 'WIF', 'FLOKI', 'TRUMP'],
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
    def __init__(self, ib, cfg):
        self.ib = ib
        self.cfg = cfg

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
        return min(max(kelly * 0.5, self.cfg.MIN_RISK_PERCENT), self.cfg.MAX_RISK_PERCENT)

    def calc_size(self, price, sl_pct, avail):
        margin = self.cfg.FIXED_MARGIN
        if avail < margin:
            margin = max(avail * 0.95, 0)
        if margin <= 0:
            return 0, 0
        qty = (margin * self.cfg.LEVERAGE) / price
        return max(qty, 0), margin


# =========================================================
# TECH ANALYZER
# =========================================================
class TechAnalyzer:
    def __init__(self, ex, cfg):
        self.ex = ex
        self.cfg = cfg

    def klines(self, s, tf='1h', limit=100):
        try:
            o = self.ex.fetch_ohlcv(s, tf, limit=limit)
            df = pd.DataFrame(o, columns=['t', 'o', 'h', 'l', 'c', 'v'])
            df['t'] = pd.to_datetime(df['t'], unit='ms')
            return df
        except Exception:
            return None

    def rsi(self, prices, period=14):
        if len(prices) < period + 1:
            return 50
        d = np.diff(prices)
        g = np.where(d > 0, d, 0)
        l = np.where(d < 0, -d, 0)
        ag = pd.Series(g).rolling(period).mean().iloc[-1]
        al = pd.Series(l).rolling(period).mean().iloc[-1]
        if al == 0:
            return 100
        return 100 - (100 / (1 + ag / al))

    def atr(self, df, period=14):
        if df is None or len(df) < period + 1:
            return 0
        h, l, c = df['h'].values, df['l'].values, df['c'].values
        tr = np.maximum(h[1:] - l[1:], np.maximum(abs(h[1:] - c[:-1]), abs(l[1:] - c[:-1])))
        return float(np.mean(tr[-period:]))

    def _round(self, p):
        if p <= 0:
            return False
        try:
            m = 10 ** math.floor(math.log10(p))
            return any(abs(p - m * f) / p < 0.005 for f in [0.5, 1, 2, 5, 10])
        except Exception:
            return False

    def _levels_tf(self, df, tc, w, lbl):
        if df is None or len(df) < 30:
            return []
        h, l, c, v = df['h'].values, df['l'].values, df['c'].values, df['v'].values
        last = c[-1]
        out = []
        win = 5
        av = np.mean(v[-20:]) if len(v) >= 20 else np.mean(v)

        for i in range(win, len(df) - win):
            if h[i] == max(h[i - win:i + win + 1]):
                lv = h[i]
                t = sum(1 for j in range(len(h)) if abs(h[j] - lv) / lv < 0.008)
                if t >= tc:
                    dist = abs(last - lv) / last
                    if dist < 0.15:
                        vat = sum(v[j] for j in range(len(h)) if abs(h[j] - lv) / lv < 0.008)
                        vs = min(vat / (av * t), 3.0) if av > 0 else 1.0
                        rb = self.cfg.ROUND_NUMBER_BONUS if self._round(lv) else 0
                        out.append({
                            'price': lv, 'touches': t, 'distance': dist,
                            'weight': w * (1 + t * 0.3) * vs + rb,
                            'tf': lbl, 'type': 'resistance', 'vs': vs
                        })

            if l[i] == min(l[i - win:i + win + 1]):
                lv = l[i]
                t = sum(1 for j in range(len(l)) if abs(l[j] - lv) / lv < 0.008)
                if t >= tc:
                    dist = abs(last - lv) / last
                    if dist < 0.15:
                        vat = sum(v[j] for j in range(len(l)) if abs(l[j] - lv) / lv < 0.008)
                        vs = min(vat / (av * t), 3.0) if av > 0 else 1.0
                        rb = self.cfg.ROUND_NUMBER_BONUS if self._round(lv) else 0
                        out.append({
                            'price': lv, 'touches': t, 'distance': dist,
                            'weight': w * (1 + t * 0.3) * vs + rb,
                            'tf': lbl, 'type': 'support', 'vs': vs
                        })
        return out

    def _cluster(self, levels, cd):
        if not levels:
            return []
        s = sorted(levels, key=lambda x: x['price'])
        cl = []
        cur = [s[0]]
        for l in s[1:]:
            if abs(l['price'] - cur[-1]['price']) / cur[-1]['price'] < cd:
                cur.append(l)
            else:
                cl.append(cur)
                cur = [l]
        cl.append(cur)

        out = []
        for c in cl:
            tw = sum(l['weight'] for l in c)
            if tw == 0:
                continue
            tfs = list(set(l['tf'] for l in c))
            ty = [l['type'] for l in c]
            out.append({
                'price': sum(l['price'] * l['weight'] for l in c) / tw,
                'touches': sum(l['touches'] for l in c),
                'distance': min(l['distance'] for l in c),
                'weight': tw + len(tfs) * 1.5,
                'tfs': sorted(tfs),
                'type': max(set(ty), key=ty.count),
                'vs': max(l['vs'] for l in c)
            })
        return out

    def levels_mtf(self, s):
        all_l = []
        for c in self.cfg.TIMEFRAMES:
            df = self.klines(s, c['tf'], c['limit'])
            if df is None or len(df) < 30:
                continue
            all_l.extend(self._levels_tf(df, self.cfg.TOUCH_COUNT, c['weight'], c['label']))

        if not all_l:
            return {'support': [], 'resistance': []}

        cl = self._cluster(all_l, self.cfg.LEVEL_CLUSTER_DISTANCE)
        out = {'support': [], 'resistance': []}
        for c in cl:
            out[c['type']].append(c)

        for k in out:
            out[k] = sorted(out[k], key=lambda x: x['weight'], reverse=True)
        return out

    def quality(self, df, lv, d, c):
        if len(df) < 30:
            return False
        vm = df['v'].rolling(20).mean().iloc[-1]
        if df['v'].iloc[-1] < vm * c.MIN_VOLUME_RATIO:
            return False
        lc = df.iloc[-1]
        if d == 'LONG' and lc['c'] <= lv * (1 + c.ZONE_WIDTH * 0.5):
            return False
        if d == 'SHORT' and lc['c'] >= lv * (1 - c.ZONE_WIDTH * 0.5):
            return False
        r = self.rsi(df['c'].values)
        if d == 'LONG' and r > 80:
            return False
        if d == 'SHORT' and r < 20:
            return False
        m20, m50 = df['c'].iloc[-20:].mean(), df['c'].iloc[-50:].mean()
        if d == 'LONG' and m20 < m50 * 0.98:
            return False
        if d == 'SHORT' and m20 > m50 * 1.02:
            return False
        a = self.atr(df)
        if a > 0 and abs(lc['c'] - lc['o']) < a * 0.3:
            return False
        return True


# =========================================================
# BOT
# =========================================================
class BreakoutBot:
    def __init__(self, ib=1000):
        self.cfg = TradingConfig()
        self.ib = ib
        self.balance = ib
        self.positions = {}
        self.closed = []
        self.wins = 0
        self.fees = 0.0
        self.runtime_bl = set()

        self.rm = RiskManager(ib, self.cfg)
        mexc = ccxt.mexc({'enableRateLimit': True, 'options': {'defaultType': 'spot'}})
        self.ta = TechAnalyzer(mexc, self.cfg)
        self.ex = mexc

        self.tg = TelegramNotifier(TELEGRAM_TOKEN, TELEGRAM_CHAT_ID)
        self.xl = ExcelLogger(ib)
        self.last_hour = time.time()

        self.cooldown = {}
        self.streak = defaultdict(int)
        self.sector_pos = defaultdict(int)
        self.emerg = False

        self.state_file = 'bot_state.json'
        self.log_file = 'trading_log_v9.csv'

        if not os.path.exists(self.log_file):
            with open(self.log_file, 'w') as f:
                f.write("timestamp,symbol,side,entry,exit,quantity,leverage,pnl,pnl_percent,fee,reason,sector,tfs,weight\n")

        self._load()
        print(f"🚀 BOT v9.7 | ${self.balance:.2f} | сделок: {len(self.closed)} | откр: {len(self.positions)}")

    def _load(self):
        if not os.path.exists(self.state_file):
            print(f"🆕 Первый запуск — старт ${self.ib:.2f}")
            return
        try:
            with open(self.state_file, 'r', encoding='utf-8') as f:
                s = json.load(f)
            self.balance = s.get('balance', self.ib)
            self.fees = s.get('total_fees', 0)
            self.wins = s.get('win_trades', 0)
            self.closed = s.get('closed_trades', [])
            for sym, p in s.get('positions', {}).items():
                if all(k in p for k in ['symbol', 'type', 'entry_price', 'quantity']):
                    p.setdefault('trailing_distance', self.cfg.TRAILING_STOP)
                    p.setdefault('tfs', [])
                    p.setdefault('weight', 0)
                    self.positions[sym] = p
                    self.sector_pos[p.get('sector', 'OTHER')] += 1
            self.cooldown = s.get('reentry_cooldown_until', {})
            self.streak = defaultdict(int, s.get('loss_streak', {}))
            print(f"✅ Загружено: ${self.balance:.2f} | сделок {len(self.closed)} | позиций {len(self.positions)}")
        except Exception as e:
            print(f"⚠️ Ошибка загрузки: {e}")

    def _save(self):
        try:
            s = {
                'version': '9.7',
                'timestamp': datetime.now().isoformat(),
                'balance': self.balance,
                'total_fees': self.fees,
                'win_trades': self.wins,
                'closed_trades': self.closed,
                'positions': {k: dict(v) for k, v in self.positions.items()},
                'reentry_cooldown_until': self.cooldown,
                'loss_streak': dict(self.streak),
                'initial_balance': self.ib
            }
            with open(self.state_file + '.tmp', 'w', encoding='utf-8') as f:
                json.dump(s, f, indent=2, default=str)
            os.replace(self.state_file + '.tmp', self.state_file)
        except Exception as e:
            print(f"⚠️ Save: {e}")

    def symbols(self):
        try:
            m = self.ex.load_markets()
            return [
                s for s in m
                if s.endswith('/USDT')
                and not any(b in s.replace('/USDT', '').upper() for b in self.cfg.BLACKLIST)
                and s not in self.runtime_bl
                and m[s].get('spot')
            ]
        except Exception:
            return []

    def sector(self, s):
        b = s.replace('/USDT', '')
        for sec, c in self.cfg.SECTORS.items():
            if b in c:
                return sec
        return 'OTHER'

    def diversify(self, s):
        sec = self.sector(s)
        if sec in self.cfg.EXCLUDED_SECTORS:
            return False
        if sec == 'OTHER':
            return self.sector_pos.get('OTHER', 0) < self.cfg.MAX_OTHER_SECTOR_POSITIONS
        return self.sector_pos.get(sec, 0) < 2

    def _cool(self, s):
        return s in self.cooldown and time.time() < self.cooldown[s]

    def signals(self):
        syms = self.symbols()
        if not syms:
            return []
        tick = {}
        try:
            at = self.ex.fetch_tickers()
            for s in syms:
                if s in at and at[s].get('quoteVolume', 0) > self.cfg.MIN_VOLUME_USD:
                    tick[s] = at[s]
        except Exception:
            return []

        sorted_s = sorted(tick.keys(), key=lambda x: tick[x].get('quoteVolume', 0), reverse=True)[:80]
        print(f"🔍 MTF-скан {len(sorted_s)} монет...")
        out = []

        for s in sorted_s:
            if not self.diversify(s) or self._cool(s) or s in self.positions:
                continue
            t = tick[s]
            price = t.get('last')
            if not price or price <= 0:
                continue
            lv = self.ta.levels_mtf(s)
            if not lv['support'] and not lv['resistance']:
                continue
            df = self.ta.klines(s, '1h', 100)
            if df is None or len(df) < 50:
                continue
            a = self.ta.atr(df)
            if price > 0 and a / price > self.cfg.MAX_VOLATILITY:
                continue
            out.extend(self._check(df, lv, s, price, a, t.get('quoteVolume', 0)))

        out.sort(key=lambda x: x['weight'] * 0.5 + x['volume_ratio'] * 0.3 + (1 - x['distance']) * 0.2, reverse=True)
        return out

    def _check(self, df, lv, s, price, a, vol):
        if vol < self.cfg.MIN_VOLUME_USD:
            return []
        sl_p = min(max((a / price) * self.cfg.ATR_SL_MULTIPLIER, self.cfg.MIN_SL_PERCENT), self.cfg.MAX_SL_PERCENT) if a and price else self.cfg.STOP_LOSS
        sig = []

        for l in lv['resistance']:
            if l['weight'] < self.cfg.MIN_LEVEL_WEIGHT:
                continue
            p = l['price']
            if p * (1 - self.cfg.ZONE_WIDTH) <= price <= p * (1 + self.cfg.ZONE_WIDTH):
                if not self.ta.quality(df, p, 'LONG', self.cfg):
                    continue
                sig.append({
                    'type': 'LONG', 'symbol': s,
                    'entry_price': p * 1.001, 'level_price': p,
                    'sl': p * (1 - sl_p), 'sl_percent': sl_p,
                    'tp': p * (1 + self.cfg.PROFIT_TARGET),
                    'volume_ratio': df['v'].iloc[-1] / df['v'].iloc[-20:].mean(),
                    'touches': l['touches'], 'distance': l['distance'],
                    'weight': l['weight'], 'tfs': l['tfs'],
                    'strength': min(l['weight'] / 10, 1)
                })

        for l in lv['support']:
            if l['weight'] < self.cfg.MIN_LEVEL_WEIGHT:
                continue
            p = l['price']
            if p * (1 - self.cfg.ZONE_WIDTH) <= price <= p * (1 + self.cfg.ZONE_WIDTH):
                if not self.ta.quality(df, p, 'SHORT', self.cfg):
                    continue
                sig.append({
                    'type': 'SHORT', 'symbol': s,
                    'entry_price': p * 0.999, 'level_price': p,
                    'sl': p * (1 + sl_p), 'sl_percent': sl_p,
                    'tp': p * (1 - self.cfg.PROFIT_TARGET),
                    'volume_ratio': df['v'].iloc[-1] / df['v'].iloc[-20:].mean(),
                    'touches': l['touches'], 'distance': l['distance'],
                    'weight': l['weight'], 'tfs': l['tfs'],
                    'strength': min(l['weight'] / 10, 1)
                })
        return sig

    def reserved(self):
        return sum(p.get('required_margin', 0) for p in self.positions.values() if not p.get('entry_filled'))

    def open_pos(self, sig):
        if len(self.positions) >= self.cfg.MAX_POSITIONS or self.emerg:
            return False
        s = sig['symbol']
        if self._cool(s) or not self.diversify(s):
            return False
        e = sig['entry_price']
        q, _ = self.rm.calc_size(e, sig['sl_percent'], self.balance - self.reserved())
        if q <= 0 or q > self.cfg.MAX_TOKEN_QUANTITY:
            return False
        m = q * e / self.cfg.LEVERAGE

        pos = {
            'symbol': s, 'type': sig['type'],
            'entry_price': e, 'quantity': q,
            'sl': sig['sl'], 'tp': sig['tp'],
            'sl_percent': sig['sl_percent'],
            'leverage': self.cfg.LEVERAGE,
            'level_price': sig['level_price'],
            'tfs': sig.get('tfs', []),
            'weight': sig.get('weight', 0),
            'entry_time': datetime.now().isoformat(),
            'created_at': time.time(),
            'required_margin': m,
            'current_price': e, 'pnl': 0, 'pnl_percent': 0,
            'entry_filled': False,
            'sector': self.sector(s),
            'tp1_hit': False,
            'trailing_high': e, 'trailing_low': e,
            'trailing_distance': self.cfg.TRAILING_STOP,
            'atr_entry': self.ta.atr(self.ta.klines(s, '5m', 20)),
            'max_profit_reached': 0
        }
        self.positions[s] = pos
        self.sector_pos[pos['sector']] += 1
        print(f"📈 ЛИМИТ {s} {sig['type']} | {'+'.join(sig.get('tfs', []))} | вес {sig.get('weight', 0):.1f} | залог ${m:.2f}")
        self.tg.notify_open(pos, m)
        self._save()
        return True

    def check(self):
        for s in list(self.positions.keys()):
            try:
                p = self.positions.get(s)
                if p is None:
                    continue

                if not p['entry_filled'] and time.time() - p['created_at'] > self.cfg.ORDER_TIMEOUT_SEC:
                    print(f"⏱️ Отмена {s}")
                    if p.get('sector') in self.sector_pos:
                        self.sector_pos[p['sector']] = max(0, self.sector_pos[p['sector']] - 1)
                    del self.positions[s]
                    self._save()
                    continue

                t = self.ex.fetch_ticker(s)
                if not t or t.get('last') is None:
                    continue
                cp = t['last']

                if not p['entry_filled']:
                    filled = (
                        (p['type'] == 'LONG' and cp >= p['entry_price']) or
                        (p['type'] == 'SHORT' and cp <= p['entry_price'])
                    )
                    if filled:
                        p['entry_filled'] = True
                        f = p['quantity'] * p['entry_price'] * self.cfg.MAKER_FEE
                        self.balance -= p['required_margin'] + f
                        self.fees += f
                        print(f"✅ ИСПОЛНЕН {s} (maker 0%)")
                        self._save()
                    continue

                self._upd(s, cp)
                self._exit(s, cp)
            except Exception as e:
                err = str(e)
                if "headers" in err or "not associated" in err:
                    self.runtime_bl.add(s)
                    try:
                        if s in self.positions:
                            p = self.positions[s]
                            self.sector_pos[p['sector']] = max(0, self.sector_pos[p['sector']] - 1)
                            del self.positions[s]
                            self._save()
                    except Exception:
                        pass
                else:
                    print(f"❌ {s}: {err[:80]}")

    def _upd(self, s, cp):
        p = self.positions[s]
        e = p['entry_price']
        if p['type'] == 'LONG':
            p['pnl_percent'] = (cp - e) / e * 100
            p['pnl'] = (cp - e) * p['quantity']
        else:
            p['pnl_percent'] = (e - cp) / e * 100
            p['pnl'] = (e - cp) * p['quantity']
        p['current_price'] = cp
        if p['pnl_percent'] > p.get('max_profit_reached', 0):
            p['max_profit_reached'] = p['pnl_percent']

    def _exit(self, s, cp):
        p = self.positions[s]
        if not p['entry_filled']:
            return

        if not p['tp1_hit'] and p['pnl_percent'] >= self.cfg.PROFIT_TARGET * 100:
            self._tp1(s, cp)
            return

        if p['tp1_hit'] and p['quantity'] > 0:
            if p['type'] == 'LONG':
                if cp > p['trailing_high']:
                    p['trailing_high'] = cp
                if cp <= p['trailing_high'] * (1 - p['trailing_distance']):
                    self._close(s, cp, "Trailing Stop")
                    return
            else:
                if cp < p['trailing_low']:
                    p['trailing_low'] = cp
                if cp >= p['trailing_low'] * (1 + p['trailing_distance']):
                    self._close(s, cp, "Trailing Stop")
                    return

        if p['quantity'] > 0:
            if (p['type'] == 'LONG' and cp <= p['sl']) or (p['type'] == 'SHORT' and cp >= p['sl']):
                self._close(s, cp, "Stop Loss")
                return

    def _tp1(self, s, cp):
        p = self.positions[s]
        p['tp1_hit'] = True
        q = p['quantity'] * 0.5

        if p['type'] == 'LONG':
            pnl = (cp - p['entry_price']) * q
        else:
            pnl = (p['entry_price'] - cp) * q

        f = q * cp * self.cfg.MAKER_FEE
        self.balance += pnl + (q * cp / p['leverage']) - f
        self.fees += f
        p['quantity'] -= q
        p['required_margin'] = p['quantity'] * p['entry_price'] / p['leverage']
        p['sl'] = p['entry_price']

        rec = {
            'symbol': s + " (TP1)", 'type': p['type'],
            'entry': p['entry_price'], 'exit': cp,
            'pnl': pnl - f, 'pnl_percent': self.cfg.PROFIT_TARGET * 100,
            'leverage': p['leverage'], 'reason': 'TP1',
            'sector': p.get('sector'),
            'entry_time': p['entry_time'],
            'quantity': q, 'fee': f,
            'tfs': p.get('tfs', []),
            'weight': p.get('weight', 0)
        }
        self.closed.append(rec)
        if rec['pnl'] > 0:
            self.wins += 1
        self._log(rec)
        self.xl.log(rec, self.closed, len(self.positions))

        p['trailing_high'] = cp
        p['trailing_low'] = cp
        if p.get('atr_entry', 0) > 0:
            p['trailing_distance'] = max(self.cfg.TRAILING_STOP, (p['atr_entry'] / cp) * 0.5)

        print(f"✅ TP1 {s} +${rec['pnl']:.2f} (maker 0%) | 🛡 стоп в БУ")
        self.tg.notify_tp1(s, rec['pnl'])
        self._save()

    def _close(self, s, cp, reason):
        if s not in self.positions:
            return
        p = self.positions[s]
        q = p['quantity']
        f = q * cp * self.cfg.TAKER_FEE
        net = p['pnl'] - f
        self.balance += p['pnl'] + (q * cp / p['leverage']) - f
        self.fees += f

        rec = {
            'symbol': s, 'type': p['type'],
            'entry': p['entry_price'], 'exit': cp,
            'pnl': net, 'pnl_percent': p['pnl_percent'],
            'leverage': p['leverage'], 'reason': reason,
            'sector': p.get('sector'),
            'entry_time': p['entry_time'],
            'quantity': q, 'fee': f,
            'tfs': p.get('tfs', []),
            'weight': p.get('weight', 0)
        }
        self.closed.append(rec)
        if net > 0:
            self.wins += 1

        if p.get('sector') in self.sector_pos:
            self.sector_pos[p['sector']] = max(0, self.sector_pos[p['sector']] - 1)

        if net > 0:
            self.streak[s] = 0
        else:
            self.streak[s] += 1
        cd = self.cfg.REENTRY_COOLDOWN
        if self.streak[s] >= self.cfg.LOSS_STREAK_FOR_LONG_COOLDOWN:
            cd *= self.cfg.COOLDOWN_STREAK_MULTIPLIER
        self.cooldown[s] = time.time() + cd

        self._log(rec)
        self.xl.log(rec, self.closed, len(self.positions) - 1)

        c = Fore.GREEN if net >= 0 else Fore.RED
        sg = "+" if net >= 0 else ""
        print(f"{c}🔚 {s} | {sg}${net:.2f} ({p['pnl_percent']:.2f}%) | {reason}")

        del self.positions[s]

        wins = [t for t in self.closed if t['pnl'] > 0]
        los = [t for t in self.closed if t['pnl'] <= 0]
        prof = sum(t['pnl'] for t in wins) if wins else 0
        l = sum(t['pnl'] for t in los) if los else 0
        pf = abs(prof / l) if l else 0
        total = sum(t['pnl'] for t in self.closed)
        fm = sum(p['required_margin'] for p in self.positions.values() if p['entry_filled'])
        self.tg.notify_close(rec, self.balance + fm, total, len(self.closed), len(wins), len(los), pf, self.fees)
        self._save()

    def _log(self, t):
        try:
            with open(self.log_file, 'a', encoding='utf-8') as f:
                tfs = "+".join(t.get('tfs', []))
                f.write(
                    f"{datetime.now()},{t['symbol']},{t['type']},"
                    f"{t['entry']:.8f},{t['exit']:.8f},"
                    f"{t.get('quantity', 0):.2f},{t['leverage']},"
                    f"{t['pnl']:.2f},{t['pnl_percent']:.2f},"
                    f"{t.get('fee', 0):.2f},"
                    f"{t['reason']},{t.get('sector', '?')},"
                    f"{tfs},{t.get('weight', 0):.2f}\n"
                )
        except Exception:
            pass

    def emergency(self):
        if len(self.closed) >= 5:
            total = sum(t['pnl'] for t in self.closed)
            if total / self.ib < -self.cfg.MAX_DRAWDOWN:
                return self._do_emergency(f"Просадка {total / self.ib * 100:.1f}%")

        r = sum(t['pnl'] for t in self.closed)
        u = sum(p.get('pnl', 0) for p in self.positions.values() if p['entry_filled'])
        if (r + u) / self.ib < -self.cfg.MAX_EQUITY_DRAWDOWN:
            return self._do_emergency(f"Просадка капитала")
        return False

    def _do_emergency(self, reason):
        print(f"🚨 ЭКСТРЕННЫЙ СТОП! {reason}")
        for s in list(self.positions.keys()):
            try:
                p = self.positions[s]
                if not p['entry_filled']:
                    self.sector_pos[p['sector']] = max(0, self.sector_pos[p['sector']] - 1)
                    del self.positions[s]
                    continue
                t = self.ex.fetch_ticker(s)
                self._close(s, t['last'], "Emergency")
            except Exception:
                pass
        self.emerg = True
        self._save()
        if self.tg.enabled:
            self.tg.send(f"🚨 <b>ЭКСТРЕННЫЙ СТОП</b>\n{reason}")
        return True

    def print_summary(self):
        fm = sum(p.get('required_margin', 0) for p in self.positions.values() if p['entry_filled'])
        ur = sum(p.get('pnl', 0) for p in self.positions.values() if p['entry_filled'])
        total = sum(t['pnl'] for t in self.closed)

        print(f"\n{'='*60}")
        print(f"📊 Свободно ${self.balance:,.2f} | В позициях ${fm:,.2f}")
        print(f"💰 Реальный капитал: ${self.balance + fm:,.2f}")
        print(f"📈 PnL: ${total:+,.2f} ({total / self.ib * 100:+.1f}%)")
        print(f"💸 Комиссии: ${self.fees:,.2f}")
        print(f"📋 Сделок: {len(self.closed)} | Выиграно: {self.wins}")
        print(f"Активных: {len(self.positions)}/{self.cfg.MAX_POSITIONS}")
        for s, p in self.positions.items():
            pnl = p.get('pnl', 0)
            sg = "+" if pnl >= 0 else ""
            f = "✅" if p.get('entry_filled') else "⏳"
            tp1 = "🎯" if p.get('tp1_hit') else ""
            tfs = "+".join(p.get('tfs', []))
            print(f"   {f}{tp1} {s} {p['type']} [{tfs}] {sg}${pnl:.2f}")
        print(f"Нереализ.: {ur:+.2f}")
        print(f"{'='*60}")

    def run(self):
        print("🚀 Старт цикла")
        last_scan = 0
        last_sum = 0
        last_save = 0

        while True:
            try:
                now = time.time()
                if self.emergency():
                    break
                if self.positions:
                    self.check()

                if now - last_scan >= self.cfg.SCAN_INTERVAL:
                    last_scan = now
                    print(f"\n🔍 СКАН {datetime.now().strftime('%H:%M:%S')}")
                    sigs = self.signals()
                    if sigs:
                        print(f"📋 Найдено {len(sigs)}:")
                        for sg in sigs[:5]:
                            print(f"   {sg['symbol']} {sg['type']} | {'+'.join(sg.get('tfs', []))} | вес {sg.get('weight', 0):.1f}")
                        op = 0
                        for sg in sigs:
                            if len(self.positions) >= self.cfg.MAX_POSITIONS:
                                break
                            if sg['symbol'] not in self.positions and self.open_pos(sg):
                                op += 1
                        if op:
                            print(f"✅ Открыто: {op}")
                    else:
                        print("ℹ️ Нет сигналов")

                if now - last_sum >= 30:
                    last_sum = now
                    self.print_summary()

                if now - last_save >= 300:
                    last_save = now
                    self._save()

                if self.tg.enabled and now - self.last_hour >= 3600:
                    self.last_hour = now
                    w = [x for x in self.closed if x['pnl'] > 0]
                    ls = [x for x in self.closed if x['pnl'] <= 0]
                    prof = sum(x['pnl'] for x in w) if w else 0
                    losv = sum(x['pnl'] for x in ls) if ls else 0
                    pf = abs(prof / losv) if losv else 0
                    total = sum(t['pnl'] for t in self.closed)
                    fm = sum(p['required_margin'] for p in self.positions.values() if p['entry_filled'])
                    self.tg.send_hourly(self.balance + fm, total, len(self.closed), len(w), len(ls), pf, len(self.positions), self.ib, self.fees)

                time.sleep(self.cfg.CHECK_INTERVAL)
            except KeyboardInterrupt:
                print("\n🛑 Остановка")
                break
            except Exception as e:
                print(f"❌ Цикл: {e}")
                time.sleep(30)

        self.print_summary()
        self._save()
        if self.tg.enabled:
            self.tg.send("🛑 <b>Бот остановлен</b>")


def main():
    try:
        bot = BreakoutBot(initial_balance=1000)
        bot.run()
    except Exception as e:
        print(f"❌ Критическая: {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    main()
