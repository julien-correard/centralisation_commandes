import configparser
import sys
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

            self.backup_type = parser.get("BACKUP", "BACKUP_TYPE")
            self.backup_location = parser.get("BACKUP", "BACKUP_LOCATION")

            # --- PROTECTION ---
            self.protect_client_files = parser.get("PROTECTION", "PROTECT_CLIENT_FILES")


            # --- PATH ---
            self.base_path = parser.get("PATH", "BASE_PATH")
            self.base_path = Path(self.base_path) if self.base_path else root_path

            self.working_directory = parser.get("PATH", "WORKING_DIRECTORY")
            self.working_directory = Path(self.working_directory) if self.working_directory else root_path

            # --- FILES ---
            self.central_workbook = parser.get("FILES", "CENTRAL_WORKBOOK")
            self.output_workbook = parser.get("FILES", "OUTPUT_WORKBOOK")
            self.fichier_correspondance_clients = parser.get("FILES", "FICHIER_CORRESPONDANCE_CLIENTS")
            self.client_path = parser.get("FILES", "CLIENT_PATH")

        except configparser.NoSectionError as e:
            raise ValueError(f"Section manquante dans config.ini : {e.section}") from e

        except configparser.NoOptionError as e:
            raise ValueError(f"Clé manquante : [{e.section}] {e.option}") from e

        except ValueError as e:
            raise ValueError(f"Erreur de type dans config.ini (int attendu ?) : {e}") from e
        
        if self.backup_type not in ("none", "once", "date"):
            print(f"La clé [BACKUP_TYPE] du fichier config.ini est erronée")
            choice = input("Voulez vous continuer sans sauvegarder le fichier central? (o/n) ")
            if choice == "o":
                self.backup_type = "none"
            else:
                sys.exit()
        
        if self.protect_client_files.lower() not in ("y", "n"):
            print(f"La clé [PROTECT_CLIENT_FILES] du fichier config.ini est erronée")
            choice = input("Voulez vous continuer sans protéger les fichiers clients? (o/n) ")
            if choice == "o":
                self.protect_client_files = "n"
            else:
                sys.exit()
        if self.protect_client_files.lower() == "y":
            self.protect_client_files = True
        else:            self.protect_client_files = False

        
        self.central_workbook = self.working_directory / self.central_workbook
        self.output_workbook = self.working_directory / self.output_workbook
        self.client_path = self.working_directory / self.client_path
        self.fichier_correspondance_clients = self.base_path / self.fichier_correspondance_clients

        if self.backup_location == "local" :
            self.backup_path = self.base_path / "Backup"
        elif self.backup_location == "cloud":
            self.backup_path = self.working_directory / "Backup"
        elif self.backup_location != "none":
            raise ValueError(f"La clé [BACKUP_LOCATION] du fichier config.ini est erronée")
              

def load_config(root_path: Path):
    base_path = Path(sys.executable).parent
    path = base_path / "script" / "config.ini"

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
    content = r"""[CLIENT]
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

    ; Où seront enregistrés les fichiers backup?
    ; local = sur le disque dur ( = BASE_PATH)
    ; cloud = sur OneDrive ( = dans WORKING_DIRECTORY)
    BACKUP_LOCATION = local

    [PATH]
    ; Répertoire où se situe le programme. Peut être contourné en mettant un chemin en argument au lancement du programme.
    ; Si la clé est vide, utilise le répertoire d'où est lancé le programme. Ca peut bugger quoi.
    BASE_PATH =

    ; Répertoire où sont situés les fichiers de travail (tableau central, tableaux clients, etc.)
    ; Laisser vide pour utiliser le répertoire courant (pas d'espace SVP)
    WORKING_DIRECTORY =

    [PROTECTION]
    ; Protéger en écriture les cellules (sauf les quantités) dans les fichiers clients (y ou n)
    PROTECT_CLIENT_FILES = y

    [FILES]
    ; Fichier modèle pour le tableau central
    CENTRAL_WORKBOOK = Commandes.xlsx

    ; Fichier où sera enregistrée la sortie du programme
    OUTPUT_WORKBOOK = Commandes.xlsx

    ; Repertoire contenant les tableaux clients
    CLIENT_PATH = clients

    FICHIER_CORRESPONDANCE_CLIENTS = script\Correspondances_clients.csv
    """

    path.write_text(content, encoding="utf-8")
    print(f"Fichier {path} enregistré.\n\n"
          f"Veuillez renseigner les clefs [BASE_PATH] et [WORKING_DIRECTORY] "
          f"du fichier config.ini et redémarrer l'application."
          f"\n\nAppuyez sur Entrée pour quitter..."
        )
    input()
    sys.exit()