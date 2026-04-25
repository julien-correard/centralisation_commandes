import csv

from config import FICHIER_CORRESPONDANCE_CLIENTS

def load_client_map(path=FICHIER_CORRESPONDANCE_CLIENTS):
    client_map = {}

    try:
        with open(path, "r", encoding="utf-8") as file:
            for line in file:
                line = line.strip()

                if not line or line.startswith("#"):
                    continue

                client_excel, client_central = line.split(",")

                client_map[client_excel] = client_central
    except:
        raise ValueError(
            f"Le fichier {FICHIER_CORRESPONDANCE_CLIENTS} est introuvable."
        )

    return client_map