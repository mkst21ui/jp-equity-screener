Japanese Equity Dual Screener (TSE Long/Short)
A Python-based quantitative stock screening system for Tokyo Stock Exchange (TSE) equities, combining macro sector-rotation logic with walk-forward Expected Value (EV) filtering. Built and operated with real capital since October 2025.

Live Track Record
PeriodRealized P&LProfit FactorTradesWin RateNov 2025 – Jun 2026 (8 months)¥1,192,6082.9519849.5%

Verified via Rakuten Securities realized P&L report. Profit factor of 2.95 means gross profits are 2.95× gross losses — achieved with a near-50% win rate, demonstrating risk-reward management rather than raw directional accuracy.


System Overview
TSE Universe (~3,800 equities)
        │
        ▼
[Stage 1] Macro Sector Filter
   └─ Active flags: oil_high / yen_weak / rate_hike / us_recovery / ai_semiconductor
   └─ TSE 33-sector classification → sector boost scores
        │
        ▼
[Stage 2] Walk-Forward EV/RR Filter
   └─ Train period → EV threshold (≥0.3%) + RR ratio (≥1.0×)
   └─ Test period → entry signal validation
   └─ MA convergence filter (spread < 10%)
        │
        ▼
[Output] Ranked candidates with exit signals (ATR-based TP/SL/time stop)

Repository Structure
├── long_ma20_v5_daiken.py   # Long screener (main)
├── short_screener_v5.py     # Short screener (AND-gate logic)
├── longv2_builder.py        # Ticker universe builder (yfinance → Excel)
├── long.xlsx                # Input: TSE ticker list
├── fundamentals.xlsx        # Input: fundamental data for short screener
└── README.md

Key Features
Long Screener (long_ma20_v5_daiken.py)
Macro dial system — manually adjustable integer weights (0 to +2) per TSE sector, keyed to macro flags:
pythonACTIVE_FLAGS_OVERRIDE = {
    "oil_high":         True,   # WTI above threshold + accelerating
    "yen_weak":         True,   # USD/JPY above threshold + accelerating
    "rate_hike":        None,   # auto-detected via BOJ signals
    "us_recovery":      True,   # SPY 5-day return above threshold
    "ai_semiconductor": True,   # manual flag
}
Example sector boost configuration:
python"ai_semiconductor": {
    "電気機器":     2,   # Electronics/Semiconductors  +2
    "情報・通信業": 2,   # IT/Communications           +2
    "精密機器":     1,   # Precision instruments        +1
    "化学":         1,   # Chemicals (materials)        +1
}
Walk-forward validation — training period EV is computed separately from test period to avoid lookahead bias.
Trading value (売買代金) integration — 5-day average trading value computed from Close × Volume as a liquidity gate (≥¥60M/day). Displayed as tier label (High/Mid/Low) but deliberately excluded from EV scoring to prevent overfitting.
ATR-based exit signals — take profit at ATR×3, trailing stop at ATR×1.5, time stop at 50 days.

Short Screener (short_screener_v5.py)
AND-gate logic — technical signals (dead cross, MA200 break, RSI peak-and-decline) are only scored when fundamental deterioration is confirmed first. This prevents false positives from technical noise alone.
Fundamental gate (PER > sector avg × 1.1 AND EPS declining)
        AND
Technical signals (dead cross / MA200 break / RSI decline)
        → Score ≥ 4 → SHORT signal
Margin ratio (信用倍率) scoring — high margin ratio (買い残多め) is treated as a squeeze risk filter: stocks below 1.0× are excluded outright; 3–6× adds +1, 6×+ adds +2 to short score.

Macro Framework
The system uses a causal chain approach rather than pure technical screening:
Macro event (e.g. Hormuz Strait closure risk, WTI >$100)
        │
        ▼
Sector impact assessment (cost pass-through, FX sensitivity, demand shift)
        │
        ▼
Dial adjustment (±0 to ±2 per TSE 33-sector classification)
        │
        ▼
Screener output re-weighted accordingly
Sectors actively rotated in response to 2025–2026 macro environment:

Mining / Wholesale: boosted to +2 under oil_high (commodity revenue exposure)
Air Transport / Chemicals: reduced under oil cost pressure scenarios
Electronics / IT: boosted to +2 under ai_semiconductor flag


Requirements
Python 3.10+
pandas
numpy
yfinance
openpyxl
bashpip install pandas numpy yfinance openpyxl
Usage
bash# Build ticker universe from TSE list
python longv2_builder.py

# Run long screener
python long_ma20_v5_daiken.py longv2.xlsx

# Run short screener (requires fundamentals.xlsx)
python short_screener_v5.py

Output Sample
The long screener outputs a multi-sheet Excel workbook:

Summary sheet — all candidates ranked by EV × macro boost, with trading value tier
Per-ticker sheets — walk-forward statistics, drawdown analysis, beta decomposition, robustness across MA windows (MA15/20/30), ATR-based exit levels


Design Principles

No overfitting via trading value — liquidity metrics are display-only; they gate entry but do not inflate EV scores
Walk-forward separation — train/test split enforced; no lookahead in signal generation
AND-gate for shorts — fundamental confirmation required before technical signals score
Manual macro dials — human-in-the-loop for macro judgments; automation handles signal detection only
Documented live performance — all P&L figures are from actual brokerage records, not backtests


Author
Kaisei Machida — Finance & Business Analytics, Monash University Malaysia
Japanese equity swing trader (3–50 day holds, long/short), Oct 2025–present
