import pandas as pd
import yfinance as yf
import sys
import os

INPUT_FILE  = "long.xlsx"
OUTPUT_FILE = "longv2.xlsx"

# ── 読み込み ──────────────────────────────────────────────────
if not os.path.exists(INPUT_FILE):
    print(f"[ERROR] {INPUT_FILE} が見つかりません")
    sys.exit(1)

df = pd.read_excel(INPUT_FILE, header=None, dtype=str)
df.columns = ["銘柄コード"] + [f"col{i}" for i in range(1, len(df.columns))]
df["銘柄コード"] = df["銘柄コード"].str.strip()

# ヘッダー行を除去
df = df[~df["銘柄コード"].str.contains(r'[^\d\.\^]', na=True) | df["銘柄コード"].str.endswith(".T")]
df = df[df["銘柄コード"].notna() & (df["銘柄コード"] != "")]

# 数字のみなら .T を付与
df["銘柄コード"] = df["銘柄コード"].apply(
    lambda t: f"{t}.T" if t.isdigit() else t
)

tickers = df["銘柄コード"].tolist()
print(f"[INFO] {len(tickers)} 銘柄の業種をyfinanceから取得中...")

# ── yfinanceから業種を直接取得 ────────────────────────────────
def get_industry(ticker: str) -> str:
    try:
        info = yf.Ticker(ticker).info
        return (
            info.get("industryDisp")
            or info.get("industry")
            or info.get("sector")
            or "未分類"
        )
    except Exception:
        return "未分類"

industry_list = []
for i, ticker in enumerate(tickers, 1):
    ind = get_industry(ticker)
    industry_list.append(ind)
    print(f"  [{i:>3}/{len(tickers)}] {ticker:<12} → {ind}")

df["業種"] = industry_list
result = df[["銘柄コード", "業種"]].reset_index(drop=True)

# ── 出力 ─────────────────────────────────────────────────────
result.to_excel(OUTPUT_FILE, index=False)
print(f"\n✅ 完了: {OUTPUT_FILE} に {len(result)} 銘柄を保存しました")
print(f"   未分類: {(result['業種'] == '未分類').sum()} 銘柄")