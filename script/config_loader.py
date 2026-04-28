import configparser
from pathlib import Path

class Config:
    def __init__(self, parser, root_path):
        try:
            # --- CLIENT ---
            self.lowest_client_row = parser.getint("CLIENT", "LOWEST_CLIENT_ROW")
            self.highest_client_row = parser.getint("CLIENT", "HIGHEST_CLIENT_ROW")

            # --- CENTRAL ---
            self.lowest_central_row = parser.getint("CENTRAL", "LOWEST_CENTRAL_ROW")
            self.highest_central_row = parser.getint("CENTRAL", "HIGHEST_CENTRAL_ROW")

            # --- LAYOUT ---
            self.client_name_row = parser.getint("LAYOUT", "CLIENT_NAME_ROW")
            self.client_name_column = parser.getint("LAYOUT", "CLIENT_NAME_COLUMN")

            self.client_row = parser.getint("LAYOUT", "CLIENT_ROW")
            self.first_client_column = parser.getint("LAYOUT", "FIRST_CLIENT_COLUMN")

            # --- FILES ---
            self.central_workbook = root_path / parser.get("FILES", "CENTRAL_WORKBOOK")
            self.output_workbook = root_path /parser.get("FILES", "OUTPUT_WORKBOOK")
            self.fichier_correspondance_clients = root_path /parser.get("FILES", "FICHIER_CORRESPONDANCE_CLIENTS")
            self.client_path = root_path /parser.get("FILES", "CLIENT_PATH")

        except configparser.NoSectionError as e:
            raise ValueError(f"Section manquante dans config.ini : {e.section}") from e

        except configparser.NoOptionError as e:
            raise ValueError(f"Clé manquante : [{e.section}] {e.option}") from e

        except ValueError as e:
            raise ValueError(f"Erreur de type dans config.ini (int attendu ?) : {e}") from e

def load_config(root_path: Path):
    path = root_path / "script" / "config.ini"

    if not path.exists():
        raise FileNotFoundError(f"Fichier de config introuvable : {path}")

    parser = configparser.ConfigParser()
    parser.read(str(path))

    return Config(parser, root_path)