import pandas as pd
import numpy as np
import yfinance as yf
from datetime import datetime, timedelta
import math
import sys
import logging
from dataclasses import dataclass
from enum import Enum

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] %(message)s',
    datefmt='%H:%M:%S')
log = logging.getLogger(__name__)




# ============================================================
# short_screener_v5.py — Config パッチ (2026/03/25適用)
# 変更箇所のみ。元ファイルの該当部分をそのまま置き換えること。
#
# 【変更理由サマリー】
# ・原油: 高水準だが停戦期待で下落方向 → oil_highは「高コスト苦しいセクター」から
#         「コスト改善で買われてしまうセクター」に切り替え。
#         化学・陸運は原油下落でコスト改善=株が上がる→ショートに不利→除外
#         空運も同様。油安で急騰済み→ショートとしての優位性消滅
# ・円安: 158円台でまだ継続中 → 円安苦しいセクター（小売・食料品）は維持
# ・日銀: 4月利上げ不透明 → RATE_HIKE_MANUAL = False
#         不動産ショートの根拠は「金利高止まり」で残るが自動フラグは切る
# ・us_selloff: SPY直近5日+3%超の可能性→フラグがOFFになるはず（自動判定に任せる）
# ============================================================
 
class Config:
    # ================触らない↓
    MA_WINDOWS        = [20, 50, 200]
    RSI_PERIOD        = 14
    ATR_PERIOD        = 14
    ROLLING_HIGH_DAYS = 252
    # ================触らない↑
 
    VOL_LIQUIDITY_DAYS       = 20
    DRAWDOWN_EXCLUDE_THRESH  = -0.3
    MIN_DAILY_LIQUIDITY      = 450,000,000
    DEAD_CROSS_LOOKBACK_DAYS = 50
    MA200_FRESH_BREAK_DAYS   = 45
    RSI_PEAK_LOOKBACK        = 30
    RSI_PEAK_THRESH          = 55
    RSI_DECLINE_BARS         = 5
    PER_RATIO_THRESH         = 1.1
    EPS_GROWTH_THRESH        = 0.0
    MARGIN_LOW_THRESH        = 0.05
    MARKET_SCORE_MIN         = 2
    MARKET_DD_THRESH         = -0.10
    ATR_VOLATILITY_MULT      = 1.2
    DEFAULT_CAPITAL          = 3_500_000
    RISK_PER_TRADE_PCT       = 0.02
    ATR_STOP_MULT            = 2.0
    DATA_PERIOD              = "1y"
    MIN_BARS                 = 200
 
    TAKE_PROFIT_ATR_MULT = 3.0
    TRAILING_STOP_ATR    = 1.5
    MAX_HOLD_DAYS        = 50
 
    SHINYO_SQUEEZE_RISK_MAX  = 1.0
    SHINYO_BOOST_THRESH_LOW  = 3.0
    SHINYO_BOOST_THRESH_HIGH = 6.0
 
    SECTOR_ETF_GROWTH = "2516.T"
    SECTOR_ETF_VALUE  = "1490.T"
 
# ============================================================
# short_screener_v5.py — Config パッチ (2026/04/25適用)
#
# 【変更理由サマリー】
# ・原油: $117→$96に急落（停戦期待）
#         → 石油・石炭・卸売のショート根拠は維持
#         → 鉱業も資源安で収益直撃→追加
#         → 空運はv5でREMOVEしたが、下位業種に居残り
#           （ホルムズ不確実性でルート混乱リスク残存）→低ブーストで復活
#         → 水産・農林業も下位入り→輸入原材料高+需要低迷でショート根拠あり
# ・AI相場: 非鉄金属・電気機器・情報通信が上位3業種
#           → us_selloffの電気機器・情報通信はショートに極めて不利→0に
# ・円安: 159円台継続 → yen_weakは維持
# ・日銀: RATE_HIKE_MANUAL = False 継続（4月利上げ見送り濃厚）
# ============================================================

    MACRO_SECTOR_BOOST = {

        # 原油下落方向が継続（$117→$96）
        # ショート狙いは「原油安で収益が直撃されるセクター」
        "oil_high": {
            "石油・石炭": 2,   # 維持: 原油下落で収益直撃、下位業種確認済み
            "卸売":       2,   # 維持→強化: 資源商社はコモディティ安で収益悪化、下位確認済み
            "鉱業":       2,   # NEW追加: 資源価格下落で直撃、下位業種入り確認
            "水産・農林": 1,   # NEW追加: 輸入原材料高+需要停滞、下位業種入り
            "空運":       1,   # 復活(低): ホルムズ不確実性でルートリスク残存、下位業種に居残り
                               #            原油安の恩恵が価格に出ていない→ショート余地あり
        },

        # 円安159円台継続 → 変更なし
        "yen_weak": {
            "小売":       2,   # 維持
            "食料品":     2,   # 維持
            "紙・パルプ": 1,   # 維持
        },

        # 日銀据え置き継続（RATE_HIKE_MANUAL=Falseなので発動しない）
        # 不動産・建設の構造的弱さはテクニカルで拾う
        "rate_hike": {
            "不動産": 2,   # 維持
            "建設":   1,   # 維持
        },

        # AI・半導体一極集中相場
        # 電気機器・情報通信・非鉄金属が上位3業種
        # → これらをショートするのは相場の逆張り、危険
        "us_selloff": {
            "電気機器":   0,   # REDUCED→0: 上位業種No.3、踏み上げリスク極大
            "情報・通信": 0,   # REDUCED→0: 上位業種No.2、AI相場の主役
            "精密機器":   1,   # 据え置き: AI直接銘柄でなければ影響限定的
        },
    }

    RATE_HIKE_MANUAL = False   # 継続: 4月利上げ見送り濃厚
 
    OIL_CHANGE_THRESH  = 0.04
    USDJPY_WEAK_THRESH = 0.02
    US_SELLOFF_THRESH  = -0.05
    # ----------------------------------------------------------
    # ▲▲▲ 変更箇所ここまで ▲▲▲
    # ----------------------------------------------------------
 
 
# ============================================================
# 【変更サマリー】
#
# フラグ        | 変更前               | 変更後
# -------------|----------------------|---------------------
# oil_high     | 空運+2,化学+1,陸運+1 | 石油石炭+2,卸売+1
# rate_hike    | 不動産+2,建設+1,金融+1| 不動産+2,建設+1(発動しない)
# us_selloff   | 情報通信+1           | 情報通信+0
# RATE_HIKE_MANUAL | True             | False
#
# 【期待される効果】
# ・原油下落局面で「コスト改善=買われる」セクターへの誤爆ショートを防ぐ
# ・石油・石炭・商社が新たなショート候補として浮上
# ・日銀据え置きで銀行ショートの根拠を自動的にキャンセル
# ============================================================
    RATE_HIKE_MANUAL   = True


# ==============================================================
# 1. データ構造
# ==============================================================
class SignalDirection(Enum):
    SHORT = "short"
    SKIP  = "skip"


@dataclass
class ShortCandidate:
    ticker:       str
    score:        int
    price:        float
    atr:          float
    ma200_dist:   float
    margin_ratio: float
    rec_shares:   int            = 0
    stop_loss:    float          = 0.0
    direction:    SignalDirection = SignalDirection.SKIP
    score_reason: str            = ""
    take_profit:     float = 0.0
    trailing_stop:   float = 0.0
    exit_date_limit: str   = ""
    macro_boost:     int   = 0
    macro_reason:    str   = ""
    shinyo_bairitu:  float = 0.0
    shinyo_boost:    int   = 0
    shinyo_reason:   str   = ""

    @property
    def risk_per_share(self) -> float:
        return self.stop_loss - self.price


@dataclass
class MarketContext:
    score:        int
    close_series: pd.Series
    is_favorable: bool = False

    def __post_init__(self):
        self.is_favorable = self.score >= Config.MARKET_SCORE_MIN


# ==============================================================
# 2. マクロ状態取得 [v5: 加速度チェック追加]
# ==============================================================
def _is_accelerating(series: pd.Series, period: int = 5) -> bool:
    """直近period日の変化率が、その前のperiod日より大きいか（加速中か）"""
    if len(series) < period * 2 + 1:
        return False
    chg_recent = (series.iloc[-1]      - series.iloc[-period])   / series.iloc[-period]
    chg_prev   = (series.iloc[-period] - series.iloc[-period*2]) / series.iloc[-period*2]
    return float(chg_recent) > float(chg_prev)


def get_macro_state() -> dict:
    state = {
        "oil_high":   False,
        "yen_weak":   False,
        "us_selloff": False,
        "rate_hike":  Config.RATE_HIKE_MANUAL,
    }

    try:
        oil = yf.download("CL=F", period="2mo", progress=False, auto_adjust=True)
        if not oil.empty:
            if isinstance(oil.columns, pd.MultiIndex):
                oil.columns = oil.columns.get_level_values(0)
            c = oil['Close'].dropna()
            if len(c) >= 22:
                chg_20       = float((c.iloc[-1] - c.iloc[-22]) / c.iloc[-22])
                accelerating = _is_accelerating(c)
                state["oil_high"] = chg_20 > Config.OIL_CHANGE_THRESH and accelerating
                log.info(f"原油20日変化率: {chg_20:.2%} 加速中:{accelerating} → oil_high={state['oil_high']}")
    except Exception as e:
        log.warning(f"原油データ取得失敗: {e}")

    try:
        fx = yf.download("JPY=X", period="2mo", progress=False, auto_adjust=True)
        if not fx.empty:
            if isinstance(fx.columns, pd.MultiIndex):
                fx.columns = fx.columns.get_level_values(0)
            c = fx['Close'].dropna()
            if len(c) >= 22:
                chg_20       = float((c.iloc[-1] - c.iloc[-22]) / c.iloc[-22])
                accelerating = _is_accelerating(c)
                state["yen_weak"] = chg_20 > Config.USDJPY_WEAK_THRESH and accelerating
                log.info(f"ドル円20日変化率: {chg_20:.2%} 加速中:{accelerating} → yen_weak={state['yen_weak']}")
    except Exception as e:
        log.warning(f"為替データ取得失敗: {e}")

    try:
        spy = yf.download("SPY", period="2mo", progress=False, auto_adjust=True)
        if not spy.empty:
            if isinstance(spy.columns, pd.MultiIndex):
                spy.columns = spy.columns.get_level_values(0)
            c = spy['Close'].dropna()
            if len(c) >= 22:
                chg_20       = float((c.iloc[-1] - c.iloc[-22]) / c.iloc[-22])
                accelerating = _is_accelerating(c)
                state["us_selloff"] = chg_20 < Config.US_SELLOFF_THRESH and accelerating
                log.info(f"SPY20日変化率: {chg_20:.2%} 加速中:{accelerating} → us_selloff={state['us_selloff']}")
    except Exception as e:
        log.warning(f"米株データ取得失敗: {e}")

    active = [k for k, v in state.items() if v]
    log.info(f"マクロ状態: {active if active else '特になし'}")
    return state


def calc_macro_boost(industry: str, macro_state: dict) -> tuple[int, str]:
    boost, reasons = 0, []
    for flag, sector_map in Config.MACRO_SECTOR_BOOST.items():
        if not macro_state.get(flag):
            continue
        if industry in sector_map:
            b = sector_map[industry]
            boost += b
            reasons.append(f"{flag}×{industry}(+{b})")
    return boost, " | ".join(reasons)


# ==============================================================
# 3. 信用倍率スコアリング
# ==============================================================
def calc_shinyo_score(shinyo: float) -> tuple[int, str]:
    """
    < 1.0      → ハード除外（呼び出し元でチェック済み）
    1.0〜3.0   → 中立（0）
    3.0〜6.0   → ショート有利（買い残多い＝崩れたとき下落加速）→ +1
    6.0超      → ショートに非常に有利 → +2
    """
    if pd.isna(shinyo) or shinyo <= 0:
        return 0, "信用倍率データなし"
    if shinyo >= Config.SHINYO_BOOST_THRESH_HIGH:
        return 2, f"信用倍率{shinyo:.1f}倍（超過熱・売り圧力予備軍）(+2)"
    elif shinyo >= Config.SHINYO_BOOST_THRESH_LOW:
        return 1, f"信用倍率{shinyo:.1f}倍（買い残多め・要注視）(+1)"
    else:
        return 0, f"信用倍率{shinyo:.1f}倍（中立）"


# ==============================================================
# 4. ファンダメンタルの動的取得 [v5: pd.notna修正済み]
# ==============================================================
def fetch_dynamic_fundamentals(ticker_t: str) -> dict:
    result = {}
    try:
        info = yf.Ticker(ticker_t).info
        per = info.get('trailingPE') or info.get('forwardPE')
        if per:
            result['per_dynamic'] = float(per)
        margin = info.get('operatingMargins')
        if margin:
            result['margin_dynamic'] = float(margin)
        eps_curr = info.get('trailingEps')
        eps_fwd  = info.get('forwardEps')
        if pd.notna(eps_curr) and pd.notna(eps_fwd) and eps_curr != 0:
            result['eps_growth_dynamic'] = (eps_fwd - eps_curr) / abs(eps_curr)
    except Exception:
        pass
    return result


# ==============================================================
# 5. テクニカル指標
# ==============================================================
def _flatten_columns(df: pd.DataFrame) -> pd.DataFrame:
    if isinstance(df.columns, pd.MultiIndex):
        df.columns = df.columns.get_level_values(0)
    return df


def compute_indicators(df: pd.DataFrame) -> pd.DataFrame:
    out = _flatten_columns(df.copy())
    for m in Config.MA_WINDOWS:
        out[f'ma{m}'] = out['Close'].rolling(window=m).mean()

    delta    = out['Close'].diff()
    gain     = delta.clip(lower=0)
    loss     = (-delta).clip(lower=0)
    avg_gain = gain.ewm(alpha=1/Config.RSI_PERIOD, adjust=False).mean()
    avg_loss = loss.ewm(alpha=1/Config.RSI_PERIOD, adjust=False).mean()
    rs       = avg_gain / avg_loss.replace(0, np.nan)
    out['rsi14'] = 100 - (100 / (1 + rs))

    hl  = out['High'] - out['Low']
    hc  = (out['High'] - out['Close'].shift()).abs()
    lc  = (out['Low']  - out['Close'].shift()).abs()
    tr  = pd.concat([hl, hc, lc], axis=1).max(axis=1)
    out['atr14'] = tr.ewm(alpha=1/Config.ATR_PERIOD, adjust=False).mean()

    out['rolling_high_252'] = out['Close'].rolling(Config.ROLLING_HIGH_DAYS).max()
    out['drawdown']         = (out['Close'] - out['rolling_high_252']) / out['rolling_high_252']
    out['vol_ma20']         = out['Volume'].rolling(Config.VOL_LIQUIDITY_DAYS).mean()
    return out


# ==============================================================
# 6. シグナル検出ヘルパー
# ==============================================================
def detect_dead_cross(df):
    lb = Config.DEAD_CROSS_LOOKBACK_DAYS
    if len(df) < lb + 1:
        return False
    r = df.tail(lb)
    return bool(((r['ma50'].shift(1) >= r['ma200'].shift(1)) & (r['ma50'] < r['ma200'])).any())


def detect_fresh_ma200_break(df):
    lb = Config.MA200_FRESH_BREAK_DAYS
    if len(df) < lb + 1:
        return False
    r = df.tail(lb)
    return bool(((r['Close'].shift(1) >= r['ma200'].shift(1)) & (r['Close'] < r['ma200'])).any())


def detect_rsi_peak_and_decline(df):
    lb = Config.RSI_PEAK_LOOKBACK
    nd = Config.RSI_DECLINE_BARS
    if len(df) < lb + nd:
        return False
    rsi = df['rsi14']
    if rsi.iloc[-lb:].max() <= Config.RSI_PEAK_THRESH:
        return False
    recent = rsi.iloc[-nd:].dropna()
    if len(recent) < nd:
        return False
    return np.polyfit(range(len(recent)), recent.values, 1)[0] < 0


def check_liquidity(df):
    l = df.iloc[-1]
    return float(l['Close']) * float(l['vol_ma20']) >= Config.MIN_DAILY_LIQUIDITY


# ==============================================================
# 7. 地合いフィルター [v5: セクターETF強弱追加]
# ==============================================================
def get_market_context(symbol='^N225') -> MarketContext:
    raw    = yf.download(symbol, period=Config.DATA_PERIOD, progress=False, auto_adjust=True)
    m_tech = compute_indicators(raw)
    latest = m_tech.iloc[-1]
    atr_ma20 = m_tech['atr14'].rolling(20).mean().iloc[-1]

    # 基本スコア（N225）
    base_score = sum([
        latest['Close'] < latest['ma20'],
        latest['Close'] < latest['ma50'],
        latest['Close'] < latest['ma200'],
        latest['drawdown'] < Config.MARKET_DD_THRESH,
        latest['atr14'] > atr_ma20 * Config.ATR_VOLATILITY_MULT,
    ])

    # セクター内部強弱（グロース vs バリュー）
    sector_score = 0
    try:
        g = yf.download(Config.SECTOR_ETF_GROWTH, period="3mo", progress=False, auto_adjust=True)
        v = yf.download(Config.SECTOR_ETF_VALUE,  period="3mo", progress=False, auto_adjust=True)
        if not g.empty and not v.empty:
            if isinstance(g.columns, pd.MultiIndex): g.columns = g.columns.get_level_values(0)
            if isinstance(v.columns, pd.MultiIndex): v.columns = v.columns.get_level_values(0)
            gc = g['Close'].dropna()
            vc = v['Close'].dropna()
            if len(gc) >= 22 and len(vc) >= 22:
                chg_growth  = float((gc.iloc[-1] - gc.iloc[-22]) / gc.iloc[-22])
                chg_value   = float((vc.iloc[-1] - vc.iloc[-22]) / vc.iloc[-22])
                growth_weak = chg_growth < chg_value
                if growth_weak:
                    sector_score += 1
                log.info(f"グロース20日:{chg_growth:.2%} / バリュー20日:{chg_value:.2%} → growth_weak={growth_weak}(+{sector_score})")
    except Exception as e:
        log.warning(f"セクターETF取得失敗: {e}")

    score = base_score + sector_score
    ctx   = MarketContext(score=score, close_series=m_tech['Close'])
    status = "ショート有利" if ctx.is_favorable else "ショート非推奨"
    log.info(f"地合いスコア: {score}/6（基本{base_score}+セクター{sector_score}）  {status}")
    return ctx


# ==============================================================
# 8. 出口シグナル計算
# ==============================================================
def calc_exit_signals(candidate: ShortCandidate) -> ShortCandidate:
    if candidate.price <= 0 or candidate.atr <= 0:
        return candidate
    candidate.take_profit   = round(candidate.price - candidate.atr * Config.TAKE_PROFIT_ATR_MULT, 1)
    candidate.trailing_stop = round(candidate.price + candidate.atr * Config.TRAILING_STOP_ATR, 1)
    exit_date = datetime.now() + timedelta(days=Config.MAX_HOLD_DAYS)
    while exit_date.weekday() >= 5:
        exit_date += timedelta(days=1)
    candidate.exit_date_limit = exit_date.strftime('%Y-%m-%d')
    return candidate


# ==============================================================
# 9. スクリーニング [v5: ANDゲート設計]
# ==============================================================
def screen_short_candidates(fund_df, price_data, market_ctx, macro_state) -> pd.DataFrame:
    fund_df = fund_df.copy()
    dupes = fund_df[fund_df['ticker'].duplicated(keep=False)]['ticker'].unique()
    if len(dupes):
        log.warning(f"重複tickerを除去: {list(dupes)}")
    fund_df = fund_df.drop_duplicates(subset='ticker', keep='first').reset_index(drop=True)
    fund_df['industry_avg_per'] = fund_df.groupby('industry')['per'].transform('mean')
    fund_index = fund_df.set_index('ticker')

    results   = []
    available = set(price_data.columns.get_level_values(0).unique())

    for ticker in fund_df['ticker']:
        if ticker not in available:
            continue
        try:
            df       = compute_indicators(price_data[ticker].dropna())
            if len(df) < Config.MIN_BARS:
                continue
            latest   = df.iloc[-1]
            fund_row = fund_index.loc[ticker]

            # ── ハードフィルター ──────────────────────────────
            if latest['drawdown'] < Config.DRAWDOWN_EXCLUDE_THRESH:
                log.debug(f"Skip {ticker}: 崩壊済み")
                continue
            if not check_liquidity(df):
                log.debug(f"Skip {ticker}: 流動性不足")
                continue

            shinyo     = fund_row.get('shinyo_bairitu')
            shinyo_val = float(shinyo) if pd.notna(shinyo) and shinyo != 0 else None
            if shinyo_val is not None and shinyo_val < Config.SHINYO_SQUEEZE_RISK_MAX:
                log.debug(f"Skip {ticker}: 信用倍率{shinyo_val:.1f}倍 → 踏み上げリスク大")
                continue

            # ── ファンダ動的補完 ──────────────────────────────
            ticker_t = ticker if str(ticker).endswith('.T') else f"{ticker}.T"
            dyn = fetch_dynamic_fundamentals(ticker_t)

            per_val = fund_row.get('per')
            if (pd.isna(per_val) or per_val == 0) and 'per_dynamic' in dyn:
                per_val = dyn['per_dynamic']
            margin_val = fund_row.get('margin_ratio')
            if pd.isna(margin_val) and 'margin_dynamic' in dyn:
                margin_val = dyn['margin_dynamic']
            eps_growth_val = fund_row.get('eps_growth')
            if pd.isna(eps_growth_val) and 'eps_growth_dynamic' in dyn:
                eps_growth_val = dyn['eps_growth_dynamic']

            # ── スコアリング（ANDゲート設計）────────────────────
            score, reasons = 0, []

            # (A) ファンダ評価
            industry_avg = fund_row.get('industry_avg_per', np.nan)
            per_ratio    = (per_val / industry_avg) if (
                pd.notna(per_val) and pd.notna(industry_avg) and industry_avg != 0
            ) else np.nan

            funda_per_bad = pd.notna(per_ratio) and per_ratio > Config.PER_RATIO_THRESH
            funda_eps_bad = pd.notna(eps_growth_val) and eps_growth_val < Config.EPS_GROWTH_THRESH
            funda_rev_bad = pd.notna(fund_row.get('revenue_growth')) and fund_row['revenue_growth'] < 0
            funda_mar_bad = pd.notna(margin_val) and margin_val < Config.MARGIN_LOW_THRESH

            # ファンダが「悪い」：PER高+減益 OR 減収 OR 低利益率
            funda_bad = (funda_per_bad and funda_eps_bad) or funda_rev_bad or funda_mar_bad

            if funda_per_bad and funda_eps_bad:
                score += 2
                reasons.append(f"高PER×減益(per_ratio={per_ratio:.2f})")
            if funda_rev_bad:
                score += 1
                reasons.append(f"減収(rev_g={fund_row['revenue_growth']:.2%})")
            if funda_mar_bad:
                score += 1
                reasons.append(f"低利益率({margin_val:.2%})")

            # (B) テクニカル：ファンダが悪い場合のみ加点（ANDゲート）
            if funda_bad:
                if detect_dead_cross(df):
                    score += 2
                    reasons.append("ファンダ悪化×デッドクロス")
                if detect_fresh_ma200_break(df):
                    score += 1
                    reasons.append("ファンダ悪化×MA200割れ初動")
                if detect_rsi_peak_and_decline(df):
                    score += 1
                    reasons.append("ファンダ悪化×RSI反落中")
            else:
                tech_signals = []
                if detect_dead_cross(df):           tech_signals.append("デッドクロス")
                if detect_fresh_ma200_break(df):    tech_signals.append("MA200割れ")
                if detect_rsi_peak_and_decline(df): tech_signals.append("RSI反落")
                if tech_signals:
                    log.debug(f"{ticker}: テクニカル({','.join(tech_signals)})あり・ファンダ普通のため非加点")

            # (C) 相対強度（ゲートなし）
            m_close = market_ctx.close_series.reindex(df.index).ffill()
            rs      = df['Close'] / m_close.replace(0, np.nan)
            rs_ma50 = rs.rolling(50).mean()
            if not rs_ma50.isna().iloc[-1] and rs.iloc[-1] < rs_ma50.iloc[-1]:
                score += 1
                reasons.append("市場比相対弱")

            # (D) マクロブースト（ゲートなし）
            industry    = str(fund_row.get('industry', ''))
            macro_boost, macro_reason = calc_macro_boost(industry, macro_state)
            score += macro_boost

            # (E) 信用倍率ブースト（ゲートなし）
            shinyo_boost, shinyo_reason = calc_shinyo_score(shinyo_val if shinyo_val else np.nan)
            score += shinyo_boost
            if shinyo_boost > 0:
                reasons.append(shinyo_reason)

            results.append(ShortCandidate(
                ticker         = ticker,
                score          = score,
                price          = float(latest['Close']),
                atr            = float(latest['atr14']),
                ma200_dist     = float((latest['Close'] - latest['ma200']) / latest['ma200']),
                margin_ratio   = float(margin_val) if pd.notna(margin_val) else 0.0,
                direction      = SignalDirection.SHORT if score >= 4 else SignalDirection.SKIP,
                score_reason   = " | ".join(reasons) if reasons else "シグナルなし",
                macro_boost    = macro_boost,
                macro_reason   = macro_reason,
                shinyo_bairitu = shinyo_val if shinyo_val else 0.0,
                shinyo_boost   = shinyo_boost,
                shinyo_reason  = shinyo_reason,
            ))

        except Exception as e:
            log.warning(f"Skip {ticker}: {e}")

    if not results:
        log.warning("候補銘柄が見つかりませんでした。")
        return pd.DataFrame()

    df_result = pd.DataFrame([vars(r) for r in results])
    df_result['direction'] = df_result['direction'].apply(
        lambda x: x.value if isinstance(x, SignalDirection) else x
    )
    return df_result.sort_values('score', ascending=False).reset_index(drop=True)


# ==============================================================
# 10. ポジションサイジング
# ==============================================================
def get_position_recommendation(candidate, capital=Config.DEFAULT_CAPITAL, risk_pct=Config.RISK_PER_TRADE_PCT):
    stop_loss      = candidate.price + candidate.atr * Config.ATR_STOP_MULT
    risk_per_share = stop_loss - candidate.price
    if risk_per_share <= 0:
        return candidate
    candidate.rec_shares = math.floor(capital * risk_pct / risk_per_share)
    candidate.stop_loss  = round(stop_loss, 1)
    candidate = calc_exit_signals(candidate)
    return candidate


# ==============================================================
# 11. Main
# ==============================================================
def main():
    try:
        fund_data = pd.read_excel('fundamentals.xlsx')
        col_map = {}
        for col in fund_data.columns:
            if col.startswith('revenue_g') and col != 'revenue_growth':
                col_map[col] = 'revenue_growth'
            elif col.startswith('eps_g') and col != 'eps_growth':
                col_map[col] = 'eps_growth'
            elif col.startswith('margin_r') and col != 'margin_ratio':
                col_map[col] = 'margin_ratio'
        if col_map:
            fund_data = fund_data.rename(columns=col_map)
        required = {'ticker', 'industry', 'per', 'revenue_growth', 'eps_growth', 'margin_ratio'}
        missing  = required - set(fund_data.columns)
        if missing:
            raise ValueError(f"fundamentals.xlsx に必要な列がありません: {missing}")
    except FileNotFoundError:
        log.error("fundamentals.xlsx が見つかりません。")
        sys.exit(1)
    except ValueError as e:
        log.error(e)
        sys.exit(1)

    log.info("=== マクロ状態を取得中 ===")
    macro_state = get_macro_state()

    market_ctx = get_market_context('^N225')
    if not market_ctx.is_favorable:
        log.error(f"ショート非推奨地合いのため終了します。地合いスコアが{Config.MARKET_SCORE_MIN}/6以上になってから再実行してください。")
        sys.exit(0)

    tickers   = [t if str(t).endswith('.T') else f"{t}.T" for t in fund_data['ticker']]
    log.info(f"{len(tickers)} 銘柄のデータを取得中...")
    price_raw = yf.download(tickers, period=Config.DATA_PERIOD, group_by='ticker',
                            auto_adjust=True, threads=True, progress=False)

    candidates_df = screen_short_candidates(fund_data, price_raw, market_ctx, macro_state)
    if candidates_df.empty:
        log.info("候補なし。終了します。")
        sys.exit(0)

    top20, sized_rows = candidates_df.head(20), []
    for _, row in top20.iterrows():
        c = ShortCandidate(
            ticker         = row['ticker'],
            score          = int(row['score']),
            price          = float(row['price']),
            atr            = float(row['atr']),
            ma200_dist     = float(row['ma200_dist']),
            margin_ratio   = float(row['margin_ratio']),
            direction      = SignalDirection(row['direction']),
            score_reason   = str(row['score_reason']),
            macro_boost    = int(row['macro_boost']),
            macro_reason   = str(row['macro_reason']),
            shinyo_bairitu = float(row['shinyo_bairitu']),
            shinyo_boost   = int(row['shinyo_boost']),
            shinyo_reason  = str(row['shinyo_reason']),
        )
        c = get_position_recommendation(c)
        sized_rows.append(vars(c))

    result_df = pd.DataFrame(sized_rows)
    result_df['direction'] = result_df['direction'].apply(
        lambda x: x.value if isinstance(x, SignalDirection) else x
    )

    display_cols = [
        'ticker', 'score', 'macro_boost', 'shinyo_boost', 'direction',
        'price', 'stop_loss', 'take_profit', 'trailing_stop', 'exit_date_limit',
        'rec_shares', 'atr', 'shinyo_bairitu', 'ma200_dist', 'margin_ratio',
        'score_reason', 'macro_reason', 'shinyo_reason',
    ]
    print("\n===== Top Short Candidates (v5) =====")
    print(result_df[display_cols].to_string(index=False))

    print("\n===== 現在のマクロ状態 =====")
    for k, v in macro_state.items():
        print(f"  {k:15s}: {'ON' if v else 'off'}")

    print("\n===== 信用倍率判定基準 =====")
    print(f"  踏み上げリスク除外 : {Config.SHINYO_SQUEEZE_RISK_MAX}倍未満")
    print(f"  ショート有利(+1)  : {Config.SHINYO_BOOST_THRESH_LOW}〜{Config.SHINYO_BOOST_THRESH_HIGH}倍")
    print(f"  ショート強力(+2)  : {Config.SHINYO_BOOST_THRESH_HIGH}倍超")

    out_path = f"short_output_v5_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
    result_df.to_excel(out_path, index=False, sheet_name='ShortCandidates')
    log.info(f"結果を保存: {out_path}")


if __name__ == "__main__":
    main()