from datetime import datetime
from pathlib import Path
import sys

from config import LAST_RUN_FILE

STATE_FILE = Path(LAST_RUN_FILE)

def print_last_run_datetime():
    last_run = get_last_run_datetime()
    if last_run:
        last_dt = datetime.fromisoformat(last_run)
        print("\nDernière exécution le", last_dt.strftime("%d/%m/%Y à %Hh%M"), "")
        input("Appuyez sur Entrée pour continuer...\n")

def save_run_datetime():
    now = datetime.now().isoformat(timespec="seconds")
    STATE_FILE.write_text(now)

def get_last_run_datetime():
    if not STATE_FILE.exists():
        return None
    return STATE_FILE.read_text().strip()

def check_already_run_today():
    last_run = get_last_run_datetime()

    if last_run:
        last_date = last_run.split("T")[0]
        today = datetime.now().date().isoformat()

        if last_date == today:
            print("!!! Le script a déjà été exécuté aujourd’hui !!!")
            choice = input("Exécuter quand même ? (o/n)").strip().lower()

            if choice == "o":
                return
            else:
                sys.exit()

