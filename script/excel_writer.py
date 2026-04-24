from typing import List

from models import Article

from config import LOWEST_CENTRAL_ROW, HIGHEST_CENTRAL_ROW, CLIENT_ROW, FIRST_CLIENT_COLUMN
from config import CENTRAL_WORKBOOK, FICHIER_CORRESPONDANCE_CLIENTS

def get_last_filled_column(sheet, row):
    last_col = None

    for col in range(1, sheet.max_column + 1):
        value = sheet.cell(row=row, column=col).value
        if value is not None and value != "":
            last_col = col

    return last_col

def get_client_column(sheet, client, client_map):

    mapped_client = client_map.get(client, client)

    last_col = get_last_filled_column(sheet, CLIENT_ROW)
    for col in range(FIRST_CLIENT_COLUMN, last_col):
        client_name = sheet.cell(row=CLIENT_ROW, column=col).value
        if mapped_client == client_name:
            return col
    
    raise ValueError(
        f"Le client {client} -> {mapped_client} "
        f"n'a pas pu être trouvé dans le fichier "
        f"{CENTRAL_WORKBOOK}.\n"
        f"Veuillez vérifier le fichier {FICHIER_CORRESPONDANCE_CLIENTS} et les fichiers .xlsx et relancer le script."
    )

def add_article(sheet, client, article: Article, client_map):

    client_column = get_client_column(sheet, client, client_map)
    
    for row in range(LOWEST_CENTRAL_ROW, HIGHEST_CENTRAL_ROW):

        order_name = sheet.cell(row=row, column=1).value
        order_quantity = sheet.cell(row=row, column=client_column).value

        if order_name == article.name:
            if order_quantity is None:
                order_quantity = 0
            else:
                try:
                    order_quantity = int(order_quantity)
                except (TypeError, ValueError):
                    order_quantity = 0

            sheet.cell(row=row, column=client_column).value = order_quantity + article.quantity
            
            return  # on sort dès qu'on a trouvé
    raise ValueError(
        f"L'article {article.name} du client {client} "
        f"n'a pas été trouvé dans le fichier {CENTRAL_WORKBOOK}.\n"
        f"Un fichier a peut être été modifié.\n"
        f"Veuillez vérifier les fichiers et relancer le script." 
    )