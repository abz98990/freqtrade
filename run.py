import os
import subprocess
import sys

def run_command(command):
    print(f"\n[+] Running: {' '.join(command)}")
    try:
        # Use shell=True on Windows if it's a batch script/executable not directly callable, 
        # but for freqtrade entrypoint, shell=False is usually fine if it's in the PATH/venv.
        subprocess.run(command, check=True)
    except subprocess.CalledProcessError as e:
        print(f"\n[-] Command failed with error code: {e.returncode}")
        sys.exit(1)
    except FileNotFoundError:
        print("\n[-] 'freqtrade' command not found.")
        print("    Make sure you have installed the project (e.g., 'pip install -e .')")
        print("    and that you have activated your virtual environment.")
        sys.exit(1)

def main():
    print("="*50)
    print("      FREQTRADE INTERACTIVE LAUNCHER")
    print("="*50)
    
    # 1. Check/Create User Directory
    user_dir = input("\nEnter user directory name [default: user_data]: ").strip()
    if not user_dir:
        user_dir = "user_data"
        
    if not os.path.exists(user_dir):
        print(f"\n[*] Creating user directory '{user_dir}'...")
        run_command(["uv", "run", "freqtrade", "create-userdir", "--userdir", user_dir])
    else:
        print(f"\n[*] User directory '{user_dir}' found.")
    
    # 2. Check/Create Config
    if os.path.exists("user.cfg"):
        config_file = "user.cfg"
        print(f"\n[*] Found existing config at '{config_file}' in main directory.")
    else:
        config_file = os.path.join(user_dir, "config.json")
        if not os.path.exists(config_file):
            print(f"\n[*] Configuration file '{config_file}' not found.")
            print("[*] Starting config wizard...")
            run_command(["uv", "run", "freqtrade", "new-config", "--config", config_file])
        else:
            print(f"[*] Found existing config at '{config_file}'.")
        
    # 3. Strategy Selection
    strategy = input("\nEnter strategy name [default: SampleStrategy]: ").strip()
    if not strategy:
        strategy = "SampleStrategy"
        
    strategy_file = os.path.join(user_dir, "strategies", f"{strategy}.py")
    if not os.path.exists(strategy_file):
        print(f"\n[*] Strategy file not found. Generating '{strategy}'...")
        run_command(["uv", "run", "freqtrade", "new-strategy", "--strategy", strategy, "--userdir", user_dir])
        
    # 4. Action Selection
    print("\nWhat would you like to do?")
    print("  1) Download Historical Data (Required before backtesting)")
    print("  2) Run Backtesting (Test strategy on historical data)")
    print("  3) Start Trading (Runs live or dry-run based on your config.json)")
    
    choice = input("\nEnter choice (1-3) [default: 3]: ").strip()
    if not choice:
        choice = "3"
        
    if choice == "1":
        exchange = input("Enter exchange (e.g., binance, kraken) [default: binance]: ").strip() or "binance"
        days = input("Enter number of days of data to download [default: 30]: ").strip() or "30"
        timeframe = input("Enter timeframe (e.g., 5m, 1h) [default: 5m]: ").strip() or "5m"
        
        run_command([
            "uv", "run", "freqtrade", "download-data",
            "--config", config_file,
            "--exchange", exchange, 
            "--days", days, 
            "--timeframes", timeframe,
            "--userdir", user_dir
        ])
        print("\n[+] Data download complete! Run this script again to backtest or trade.")
        
    elif choice == "2":
        run_command([
            "uv", "run", "freqtrade", "backtesting", 
            "--config", config_file, 
            "--strategy", strategy, 
            "--userdir", user_dir
        ])
        print("\n[*] Backtesting complete. Auto-saving results to CSV and Excel...")
        if os.path.exists("export_backtest.py"):
            run_command(["uv", "run", "python", "export_backtest.py"])
            print("[+] Results exported to user_data/backtest_results/exports/")
        else:
            print("[WARN] export_backtest.py not found — skipping auto-export.")
        
    elif choice == "3":
        print(f"\n[*] Starting bot using strategy: {strategy}")
        print("[*] Press Ctrl+C to stop the bot at any time.")
        run_command([
            "uv", "run", "freqtrade", "trade", 
            "--config", config_file, 
            "--strategy", strategy, 
            "--userdir", user_dir
        ])
        
    else:
        print("[-] Invalid choice. Exiting.")

if __name__ == "__main__":
    main()
