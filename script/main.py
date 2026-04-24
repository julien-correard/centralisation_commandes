from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

from pathlib import Path

from typing import List

from models import Article
from excel_writer import add_article
from excel_reader import get_client, read_articles
from client_mapping import load_client_map
from date import save_run_datetime, check_already_run_today, print_last_run_datetime

from config import CLIENT_NAME_ROW, CLIENT_NAME_COLUMN, LOWEST_CLIENT_ROW, HIGHEST_CLIENT_ROW
from config import CENTRAL_WORKBOOK, OUTPUT_WORKBOOK

        
def main():
    try:
        
        client_map = load_client_map()

        orders_workbook = load_workbook(CENTRAL_WORKBOOK)
        orders_sheet = orders_workbook.active

        folder = "clients"

        check_already_run_today()
        print_last_run_datetime()

        for file_path in Path(folder).glob("*.xlsx"):

            print(f"Traitement du fichier {file_path} : ", end="")

            client_workbook = load_workbook(file_path)
            client_sheet = client_workbook.active

            client = get_client(client_sheet, file_path)
            articles = read_articles(client_sheet)

            for article in articles:
                add_article(orders_sheet, client, article, client_map)
            print("OK")

        orders_workbook.save(OUTPUT_WORKBOOK)
        print(f"\nFichier {OUTPUT_WORKBOOK} enregistré avec succès.")
        save_run_datetime()

    except ValueError as e:
        print("\n!!! ERREUR !!!\n")
        print(e)
        print(f"\nLe fichier {OUTPUT_WORKBOOK} n'a pas été modifié.")
        input("\nAppuyez sur Entrée pour quitter...")

if __name__ == "__main__":
    main()