import shutil
from datetime import datetime
from pathlib import Path

def save_temp_backup_central_file(config):

    central_file = config.central_workbook
    directory = central_file.parent
    filename = central_file.name
    base, ext = filename.rsplit(".", 1)

    if config.backup == "none":
        return ""

    elif config.backup == "once":
        filename = base + "_backup" + "." + ext

    elif config.backup == "date":
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = base + "_backup_" + timestamp + ".tmp." + ext
    else:
        print(f"!!! La clé [BACKUP_TYPE] du fichier config.ini est invalide."
              f" Aucune sauvegarde ne sera effectuée. !!!\n")
        return

    backup_dir = directory / "backup"
    backup_dir.mkdir(exist_ok=True)


    save_file = backup_dir / filename

    shutil.copy(config.central_workbook, save_file)

    return save_file

def delete_temp_central_file(save_file):
    if save_file and save_file.exists():
        save_file.unlink()

def save_central_file(temp_file, working_directory):
    if temp_file and temp_file.exists():
        final_path = temp_file.with_suffix(temp_file.suffix.replace(".tmp", ""))
        temp_file.replace(final_path)
        print(f"L'ancien fichier central a été sauvegardé dans : {Path(final_path).relative_to(working_directory)}\n")

