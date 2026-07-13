import os
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
BOT = ROOT / "bot.py"


def stop_other_instances() -> None:
    """Stop other instances of THIS project's bot.py only (Windows)."""
    if sys.platform != "win32":
        return
    my_pid = os.getpid()
    # Match full path so unrelated projects named bot.py are not killed
    marker = str(BOT).replace("'", "''")
    ps = (
        f"$keep = {my_pid}; "
        f"$marker = '{marker}'; "
        "Get-CimInstance Win32_Process -Filter \"Name='python.exe' OR Name='pythonw.exe'\" | "
        "Where-Object { "
        "  $_.CommandLine -and ("
        "    $_.CommandLine -like ('*' + $marker + '*') -or "
        "    ($_.CommandLine -like '*bot.py*' -and $_.CommandLine -like '*etkinlikkt*')"
        "  )"
        "} | "
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
