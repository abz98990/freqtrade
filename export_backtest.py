"""
export_backtest.py
------------------
Converts Freqtrade backtest result ZIP files into readable CSV and Excel reports.

Usage:
    uv run python export_backtest.py                  # exports latest result
    uv run python export_backtest.py --all            # exports all results
"""

import argparse
import json
import os
import zipfile
from datetime import datetime
from pathlib import Path

# ---------------------------------------------------------------------------
# Try to import optional deps; tell the user if they are missing
# ---------------------------------------------------------------------------
try:
    import pandas as pd
except ImportError:
    raise SystemExit("pandas is required. Run: uv add pandas openpyxl")

try:
    import openpyxl  # noqa: F401 – only needed for Excel writing
    EXCEL_AVAILABLE = True
except ImportError:
    import subprocess, sys
    print("[*] openpyxl not found. Installing via uv...")
    subprocess.run(["uv", "add", "openpyxl"], check=True)
    import openpyxl  # noqa: F401
    EXCEL_AVAILABLE = True

RESULTS_DIR = Path("user_data/backtest_results")
EXPORTS_BASE = Path("user_data/backtest_results/exports")


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def find_zip_files(latest_only: bool = True) -> list[Path]:
    zips = sorted(RESULTS_DIR.glob("backtest-result-*.zip"))
    if not zips:
        raise FileNotFoundError(f"No backtest result ZIP files found in {RESULTS_DIR}")
    return [zips[-1]] if latest_only else zips


def load_backtest_zip(zip_path: Path) -> dict:
    with zipfile.ZipFile(zip_path) as zf:
        # The main JSON file has the same stem as the zip
        json_name = zip_path.stem + ".json"
        with zf.open(json_name) as f:
            return json.load(f)


def trades_to_df(trades: list[dict]) -> pd.DataFrame:
    df = pd.DataFrame(trades)
    # Convert timestamps to human-readable dates (strip tz for Excel compatibility)
    for col in ["open_date", "close_date"]:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col]).dt.tz_localize(None)
    # Round floats for readability
    float_cols = ["stake_amount", "amount", "open_rate", "close_rate",
                  "profit_ratio", "profit_abs", "stop_loss_abs", "min_rate", "max_rate"]
    for col in float_cols:
        if col in df.columns:
            df[col] = df[col].round(6)
    # Rename profit_ratio to something friendlier
    df = df.rename(columns={"profit_ratio": "profit_%", "profit_abs": "profit_USDT"})
    if "profit_%" in df.columns:
        df["profit_%"] = (df["profit_%"] * 100).round(4)
    # Drop internal/less useful columns
    drop_cols = ["orders", "open_timestamp", "close_timestamp", "is_open",
                 "initial_stop_loss_abs", "initial_stop_loss_ratio", "stop_loss_ratio",
                 "funding_fees", "fee_open", "fee_close"]
    df = df.drop(columns=[c for c in drop_cols if c in df.columns])
    return df


def build_summary_df(strategy_data: dict) -> pd.DataFrame:
    metrics = strategy_data.get("results_per_enter_tag") or strategy_data.get("results_per_pair", [])
    rows = []
    for item in metrics:
        # duration_avg comes back as a timedelta string like '17:21:00'
        raw_dur = item.get("duration_avg", "0:00:00")
        if isinstance(raw_dur, str):
            parts = raw_dur.split(":")
            try:
                dur_minutes = int(parts[0]) * 60 + int(parts[1])
            except Exception:
                dur_minutes = 0
        else:
            dur_minutes = round(float(raw_dur), 1)
        rows.append({
            "Pair / Tag":        item.get("key", ""),
            "Trades":            item.get("trades", 0),
            "Avg Profit %":      round(item.get("profit_mean", 0) * 100, 4),
            "Total Profit USDT": round(item.get("profit_total_abs", 0), 4),
            "Total Profit %":    round(item.get("profit_total", 0) * 100, 4),
            "Avg Duration (min)": dur_minutes,
            "Win":               item.get("wins", 0),
            "Draw":              item.get("draws", 0),
            "Loss":              item.get("losses", 0),
            "Win%":              round(item.get("winrate", 0) * 100, 1),
            "Sharpe":            round(item.get("sharpe", 0), 3),
            "Sortino":           round(item.get("sortino", 0), 3),
        })
    return pd.DataFrame(rows)


def build_key_metrics(strategy_data: dict, strategy_name: str) -> pd.DataFrame:
    metrics = strategy_data.get("results_metrics", {})
    if not metrics:
        return pd.DataFrame()
    rows = [
        ("Strategy",            strategy_name),
        ("Backtesting From",    str(strategy_data.get("backtest_start", ""))),
        ("Backtesting To",      str(strategy_data.get("backtest_end", ""))),
        ("Total Trades",        metrics.get("trade_count", "")),
        ("Daily Avg Trades",    round(metrics.get("trades_per_day", 0), 2)),
        ("Starting Balance",    f"{metrics.get('starting_balance', 0)} USDT"),
        ("Final Balance",       f"{round(metrics.get('final_balance', 0), 3)} USDT"),
        ("Absolute Profit",     f"{round(metrics.get('profit_total_abs', 0), 3)} USDT"),
        ("Total Profit %",      f"{round(metrics.get('profit_total', 0) * 100, 3)} %"),
        ("CAGR %",              f"{round(metrics.get('cagr', 0) * 100, 2)} %"),
        ("Win Rate",            f"{round(metrics.get('winrate', 0) * 100, 1)} %"),
        ("Sharpe Ratio",        round(metrics.get("sharpe", 0), 3)),
        ("Sortino Ratio",       round(metrics.get("sortino", 0), 3)),
        ("Max Drawdown",        f"{round(metrics.get('max_drawdown', 0) * 100, 3)} %"),
        ("Max Drawdown (USDT)", f"{round(metrics.get('max_drawdown_abs', 0), 3)} USDT"),
        ("Best Pair",           metrics.get("best_pair", {}).get("key", "")),
        ("Worst Pair",          metrics.get("worst_pair", {}).get("key", "")),
        ("Best Trade",          f"{metrics.get('best_trade', {}).get('key','')} {round(metrics.get('best_trade', {}).get('profit_ratio',0)*100,2)}%"),
        ("Worst Trade",         f"{metrics.get('worst_trade', {}).get('key','')} {round(metrics.get('worst_trade', {}).get('profit_ratio',0)*100,2)}%"),
        ("Profit Factor",       round(metrics.get("profit_factor", 0), 3)),
        ("Market Change",       f"{round(metrics.get('market_change', 0) * 100, 2)} %"),
    ]
    return pd.DataFrame(rows, columns=["Metric", "Value"])


# ---------------------------------------------------------------------------
# Export functions
# ---------------------------------------------------------------------------

def export_to_csv(zip_path: Path):
    print(f"\n[*] Loading: {zip_path.name}")
    data = load_backtest_zip(zip_path)

    # Extract full timestamp from filename e.g. '2026-09-22_06-52-38'
    timestamp = zip_path.stem.replace("backtest-result-", "")
    # Use just the date portion (YYYY-MM-DD) as the folder name
    date_folder = timestamp.split("_")[0]
    output_dir = EXPORTS_BASE / date_folder
    output_dir.mkdir(parents=True, exist_ok=True)

    for strategy_name, strategy_data in data["strategy"].items():
        trades = strategy_data.get("trades", [])
        if not trades:
            print(f"  [WARN] No trades found for strategy: {strategy_name}")
            continue

        trades_df  = trades_to_df(trades)
        summary_df = build_summary_df(strategy_data)
        metrics_df = build_key_metrics(strategy_data, strategy_name)

        # CSV export — filenames include time + strategy to avoid collisions
        prefix = f"{timestamp}_{strategy_name}"
        trades_csv  = output_dir / f"{prefix}_trades.csv"
        summary_csv = output_dir / f"{prefix}_summary.csv"
        metrics_csv = output_dir / f"{prefix}_metrics.csv"

        trades_df.to_csv(trades_csv, index=False)
        summary_df.to_csv(summary_csv, index=False)
        metrics_df.to_csv(metrics_csv, index=False)

        print(f"  [PASS] CSV exports saved to folder: {output_dir}")
        print(f"    -> {trades_csv.name}")
        print(f"    -> {summary_csv.name}")
        print(f"    -> {metrics_csv.name}")

        # Excel export (multi-sheet)
        if EXCEL_AVAILABLE:
            excel_path = output_dir / f"{prefix}_full_report.xlsx"
            with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
                metrics_df.to_excel(writer, sheet_name="Key Metrics", index=False)
                trades_df.to_excel(writer, sheet_name="All Trades", index=False)
                summary_df.to_excel(writer, sheet_name="Per Pair Summary", index=False)

                # Auto-size columns for readability
                for sheet_name in writer.sheets:
                    ws = writer.sheets[sheet_name]
                    for col in ws.columns:
                        max_len = max((len(str(cell.value)) for cell in col if cell.value), default=10)
                        ws.column_dimensions[col[0].column_letter].width = min(max_len + 4, 50)

            print(f"  [PASS] Excel report saved:")
            print(f"    -> {excel_path.name}")


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(description="Export Freqtrade backtest results to CSV/Excel")
    parser.add_argument("--all", action="store_true", help="Export all backtest results, not just the latest")
    args = parser.parse_args()

    print("=" * 55)
    print("     FREQTRADE BACKTEST RESULT EXPORTER")
    print("=" * 55)

    try:
        zip_files = find_zip_files(latest_only=not args.all)
    except FileNotFoundError as e:
        print(f"\n[-] {e}")
        print("    Run a backtest first via run.py → option 2.")
        return

    for zf in zip_files:
        export_to_csv(zf)

    print(f"\n[+] All exports saved under: {EXPORTS_BASE.resolve()}")


if __name__ == "__main__":
    main()
