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

            # --- BACKUP ---

            self.backup = parser.get("BACKUP", "BACKUP_TYPE")

            # --- PATH ---
            self.working_directory = parser.get("PATH", "WORKING_DIRECTORY")
            self.working_directory = Path(self.working_directory) if self.working_directory else root_path

            # --- FILES ---
            self.central_workbook = parser.get("FILES", "CENTRAL_WORKBOOK")
            self.output_workbook = parser.get("FILES", "OUTPUT_WORKBOOK")
            self.fichier_correspondance_clients = root_path /parser.get("FILES", "FICHIER_CORRESPONDANCE_CLIENTS")
            self.client_path = parser.get("FILES", "CLIENT_PATH")

        except configparser.NoSectionError as e:
            raise ValueError(f"Section manquante dans config.ini : {e.section}") from e

        except configparser.NoOptionError as e:
            raise ValueError(f"Clé manquante : [{e.section}] {e.option}") from e

        except ValueError as e:
            raise ValueError(f"Erreur de type dans config.ini (int attendu ?) : {e}") from e
        
        
        self.central_workbook = self.working_directory / self.central_workbook
        self.output_workbook = self.working_directory / self.output_workbook
        self.client_path = self.working_directory / self.client_path
              

def load_config(root_path: Path):
    path = root_path / "script" / "config.ini"

    if not path.exists():
        print (f"Fichier de config introuvable : {path}")
        choice = input("\nRecréer un config.ini par défaut ? (o/n) : ")

        if choice.lower() == "o":
            create_default_config(path)
        else:
            raise FileNotFoundError(f"Fichier de config introuvable : {path}")

    parser = configparser.ConfigParser()
    parser.read(str(path))

    try:
        return Config(parser, root_path)

    except Exception as e:
        print("\n!!! LE FICHIER CONFIG.INI EST INVALIDE !!!")
        print(e)

        choice = input("\nRecréer un config.ini par défaut ? (o/n) : ")

        if choice.lower() == "o":
            backup = path.with_suffix(".broken.ini")
            path.rename(backup)

            create_default_config(path)
            parser.read(path)

            return Config(parser, root_path)

        raise

def create_default_config(path: Path):
    content = """[CLIENT]
; Première ligne utile du tableau client (premier article)
LOWEST_CLIENT_ROW = 4

; Dernière ligne utile des tableaux clients
HIGHEST_CLIENT_ROW = 300


[CENTRAL]
; Première ligne utile du tableau central
LOWEST_CENTRAL_ROW = 2

; Dernière ligne utile du tableau central
HIGHEST_CENTRAL_ROW = 300


[LAYOUT]
; Position du nom du client dans son fichier
CLIENT_NAME_ROW = 1
CLIENT_NAME_COLUMN = 3

; Ligne et première colonne des clients dans le tableau central
CLIENT_ROW = 1
FIRST_CLIENT_COLUMN = 3

[BACKUP]
; Sauvegarde du fichier central :
; none = aucune, once = ecrase la dernière sauvegarde, date = sauvegarde à chaque exécution
BACKUP_TYPE = once

[PATH]
; Répertoire où sont situés les fichiers de travail (tableau central, tableaux clients, etc.)
WORKING_DIRECTORY =

[FILES]
; Fichier modèle pour le tableau central
CENTRAL_WORKBOOK = Commandes.xlsx

; Fichier où sera enregistrée la sortie du programme
OUTPUT_WORKBOOK = Commandes.xlsx

; Repertoire contenant les tableaux clients
CLIENT_PATH = clients

FICHIER_CORRESPONDANCE_CLIENTS = script/Correspondances_clients.csv

"""

    path.write_text(content, encoding="utf-8")
    print(f"fichier {path} enregistré.")