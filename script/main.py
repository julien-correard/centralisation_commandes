from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

from pathlib import Path

from typing import List

from models import Article
from excel_writer import add_article
from excel_reader import get_client, read_articles
from client_mapping import load_client_map

from config_loader import load_config

        
def main():
    try:
        
        try:
            config = load_config()
        except Exception as e:
            print(e)
            exit(1)

        client_map = load_client_map(config.fichier_correspondance_clients)

        try:
            orders_workbook = load_workbook(config.central_workbook)
        except:
            raise ValueError(
                f"Le fichier {config.central_workbook} n'a pas été trouvé."
            )
        orders_sheet = orders_workbook.active

        folder = "clients"

        files = list(Path(folder).glob("*.xlsx"))

        if not files:
            raise ValueError(
                f"Aucun fichier .xlsx trouvé dans le dossier {folder}"
            )

        for file_path in Path(folder).glob("*.xlsx"):

            print(f"Traitement du fichier {file_path} : ", end="")

            client_workbook = load_workbook(file_path)
            client_sheet = client_workbook.active

            client = get_client(client_sheet, file_path, config)
            articles = read_articles(client_sheet, client, file_path, config)

            for article in articles:
                add_article(orders_sheet, client, article, client_map, config)
            print("OK")

        orders_workbook.save(config.output_workbook)
        print(f"\nFichier {config.output_workbook} enregistré avec succès.")
        input("Appuyez sur entrée pour quitter...")

    except ValueError as e:
        message = (
        f"\n\n!!! ERREUR !!!\n\n"
        f"{e}\n\n"
        f"Le fichier {config.output_workbook} n'a pas été modifié.\n"
        f"\nAppuyez sur Entrée pour quitter..."
        )
        print(message)
        input()

if __name__ == "__main__":
    main()
