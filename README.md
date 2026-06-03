# Japanese Equity Dual Screener (TSE Long/Short)

A Python-based quantitative stock screening system for Tokyo Stock Exchange (TSE) equities, combining macro sector-rotation logic with walk-forward Expected Value (EV) filtering. Built and operated with real capital since October 2025.

---

# Live Track Record

| Period | Realized P&L | Profit Factor | Trades | Win Rate |
|----------|----------|----------|----------|----------|
| Nov 2025 – Jun 2026 (8 months) | ¥1,192,608 | 2.95 | 198 | 49.5% |

Verified via Rakuten Securities realized P&L report.

Profit factor of 2.95 means gross profits are 2.95× gross losses — achieved with a near-50% win rate, demonstrating risk-reward management rather than raw directional accuracy.

---

# System Overview

```text
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
```

---

# Repository Structure

```text
├── long_ma20_v5_daiken.py   # Long screener (main)
├── short_screener_v5.py     # Short screener (AND-gate logic)
├── longv2_builder.py        # Ticker universe builder (yfinance → Excel)
├── long.xlsx                # Input: TSE ticker list
├── fundamentals.xlsx        # Input: fundamental data for short screener
└── README.md
```

---

# Key Features

## Long Screener (`long_ma20_v5_daiken.py`)

### Macro Dial System

Manually adjustable integer weights (0 to +2) per TSE sector, keyed to macro flags:

```python
ACTIVE_FLAGS_OVERRIDE = {
    "oil_high":         True,   # WTI above threshold + accelerating
    "yen_weak":         True,   # USD/JPY above threshold + accelerating
    "rate_hike":        None,   # auto-detected via BOJ signals
    "us_recovery":      True,   # SPY 5-day return above threshold
    "ai_semiconductor": True,   # manual flag
}
```

Example sector boost configuration:

```python
"ai_semiconductor": {
    "電気機器":     2,   # Electronics/Semiconductors +2
    "情報・通信業": 2,   # IT/Communications +2
    "精密機器":     1,   # Precision instruments +1
    "化学":         1,   # Chemicals (materials) +1
}
```

### Walk-Forward Validation

Training period EV is computed separately from test period to avoid lookahead bias.

### Trading Value Integration

5-day average trading value computed from:

```text
Close × Volume
```

Used as a liquidity gate (≥¥60M/day).

Displayed as tier labels:

- High
- Mid
- Low

Deliberately excluded from EV scoring to prevent overfitting.

### ATR-Based Exit Signals

- Take profit: ATR × 3
- Trailing stop: ATR × 1.5
- Time stop: 50 days

---

## Short Screener (`short_screener_v5.py`)

### AND-Gate Logic

Technical signals are only scored when fundamental deterioration is confirmed first.

```text
Fundamental gate
(PER > sector avg × 1.1 AND EPS declining)

        AND

Technical signals
(dead cross / MA200 break / RSI decline)

        ↓

Score ≥ 4 → SHORT signal
```

This prevents false positives from technical noise alone.

### Margin Ratio (信用倍率) Scoring

High margin ratio (買い残多め) is treated as a squeeze-risk filter.

- Below 1.0× → excluded outright
- 3–6× → +1 score
- Above 6× → +2 score

---

# Macro Framework

The system uses a causal-chain approach rather than pure technical screening.

```text
Macro event
(e.g. Hormuz Strait closure risk, WTI > $100)

        │
        ▼

Sector impact assessment
(cost pass-through, FX sensitivity, demand shift)

        │
        ▼

Dial adjustment
(±0 to ±2 per TSE 33-sector classification)

        │
        ▼

Screener output re-weighted accordingly
```

### Active Sector Rotations During 2025–2026

#### Under `oil_high`

- Mining: +2
- Wholesale: +2

Rationale: commodity revenue exposure.

#### Under Oil Cost Pressure

- Air Transport: reduced
- Chemicals: reduced

#### Under `ai_semiconductor`

- Electronics: +2
- IT / Communications: +2

---

# Requirements

- Python 3.10+
- pandas
- numpy
- yfinance
- openpyxl

Install dependencies:

```bash
pip install pandas numpy yfinance openpyxl
```

---

# Usage

### Build TSE Ticker Universe

```bash
python longv2_builder.py
```

### Run Long Screener

```bash
python long_ma20_v5_daiken.py longv2.xlsx
```

### Run Short Screener

```bash
python short_screener_v5.py
```

Requires:

```text
fundamentals.xlsx
```

---

# Output Sample

The long screener outputs a multi-sheet Excel workbook.

### Summary Sheet

- All candidates ranked by EV × macro boost
- Trading value tier labels

### Per-Ticker Sheets

- Walk-forward statistics
- Drawdown analysis
- Beta decomposition
- Robustness across MA windows (MA15 / MA20 / MA30)
- ATR-based exit levels

---

# Design Principles

### No Overfitting via Trading Value

Liquidity metrics are display-only.

They gate entry but do not inflate EV scores.

### Walk-Forward Separation

Train/test split enforced.

No lookahead in signal generation.

### AND-Gate for Shorts

Fundamental confirmation is required before technical signals contribute to the score.

### Manual Macro Dials

Human-in-the-loop for macro judgments.

Automation handles signal detection only.

### Documented Live Performance

All P&L figures are from actual brokerage records, not backtests.

---

# Author

**Kaisei Machida**

Finance & Business Analytics, Monash University Malaysia

Japanese equity swing trader (3–50 day holds, long/short)

Oct 2025 – Present
