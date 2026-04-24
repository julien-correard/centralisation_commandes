from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

from typing import List

from models import Article

from config import CLIENT_NAME_ROW, CLIENT_NAME_COLUMN, LOWEST_CLIENT_ROW, HIGHEST_CLIENT_ROW
    


def get_client(sheet, client_filename):
    value = sheet.cell(row=CLIENT_NAME_ROW, column=CLIENT_NAME_COLUMN).value

    if value is None or value == "":
        # Traduction de 1 3 par exemple à C1
        col_letter = get_column_letter(CLIENT_NAME_COLUMN)
        cell_ref = f"{col_letter}{CLIENT_NAME_ROW}"

        raise ValueError(
            f"Nom du client introuvable dans la cellule "
            f"{cell_ref} "
            f"du fichier {client_filename}.\n"
            f"Veuillez corriger le fichier et relancer le script."
        )

    return value

def read_articles(sheet):
    articles = []

    for row in range(LOWEST_CLIENT_ROW, HIGHEST_CLIENT_ROW + 1):

        col = 1
        while col <= sheet.max_column:

            name = sheet.cell(row=row, column=col).value
            quantity = sheet.cell(row=row, column=col + 1).value

            if name not in (None, "") and quantity not in (None, ""):
                try:
                    quantity_int = int(quantity)
                except (TypeError, ValueError):
                    col += 2
                    continue

                articles.append(Article(name, quantity_int))

            col += 2

    return articles