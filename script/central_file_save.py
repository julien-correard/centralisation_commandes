import shutil
from datetime import datetime
import os
from pathlib import Path

def save_central_file(config, BASE_DIR):

    central_file = config.central_workbook
    directory = central_file.parent
    filename = central_file.name
    base, ext = filename.rsplit(".", 1)

    if config.backup == "none":
        return

    elif config.backup == "once":
        filename = base + "_backup" + "." + ext

    elif config.backup == "date":
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = base + "_backup_" + timestamp + "." + ext
    else:
        print(f"!!! La clé [BACKUP_TYPE] du fichier config.ini est invalide."
              f" Aucune sauvegarde ne sera effectuée. !!!\n")
        return

    backup_dir = directory / "backup"
    backup_dir.mkdir(exist_ok=True)


    save_file = backup_dir / filename

    shutil.copy(config.central_workbook, save_file)

    print(f"Fichier central sauvegardé dans : {Path(save_file).relative_to(BASE_DIR)}\n")

