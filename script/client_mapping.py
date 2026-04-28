import csv


def load_client_map(path):
    client_map = {}
    client_list = set()

    try:
        with open(path, "r", encoding="utf-8") as file:
            for line in file:
                line = line.strip()

                if not line or line.startswith("#"):
                    continue

                client_excel, client_central = line.split(",")

                client_map[client_excel] = client_central
                client_list.add(client_excel)
    except:
        raise ValueError(
            f"Le fichier {path} est introuvable."
        )

    return client_map, client_list