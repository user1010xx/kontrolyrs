import os
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
BOT = ROOT / "bot.py"


def stop_other_instances() -> None:
    if sys.platform != "win32":
        return
    my_pid = os.getpid()
    ps = (
        f"$keep = {my_pid}; "
        "Get-CimInstance Win32_Process -Filter \"Name='python.exe'\" | "
        "Where-Object { $_.CommandLine -like '*bot.py*' } | "
        "ForEach-Object { if ($_.ProcessId -ne $keep) { "
        "Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue } }"
    )
    subprocess.run(["powershell.exe", "-NoProfile", "-Command", ps], check=False)


def main() -> None:
    os.chdir(ROOT)
    stop_other_instances()
    print("Bot baslatiliyor... (durdurmak icin Ctrl+C)")
    sys.exit(subprocess.call([sys.executable, str(BOT)]))


if __name__ == "__main__":
    main()