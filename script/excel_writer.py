from models import Article


def get_last_filled_column(sheet, row):
    last_col = None

    for col in range(1, sheet.max_column + 1):
        value = sheet.cell(row=row, column=col).value
        if value is not None and value != "":
            last_col = col

    return last_col

def get_client_column(sheet, client, client_map, config):

    mapped_client = client_map.get(client, client)

    last_col = get_last_filled_column(sheet, config.client_row)
    for col in range(config.first_client_column, last_col):
        client_name = sheet.cell(row=config.client_row, column=col).value
        if mapped_client == client_name:
            return col
    
    raise ValueError(
        f"Le client {client} "
        f"n'a pas pu être trouvé dans le fichier "
        f"{config.central_workbook}.\n"
        f"Veuillez vérifier le fichier {config._fichier_correspondance_clients} "
        f"et les fichiers .xlsx et relancer le script."
    )

def add_article(sheet, client, article: Article, client_map, config):

    client_column = get_client_column(sheet, client, client_map, config)
    
    for row in range(config.lowest_central_row, config.highest_central_row):

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

            sheet.cell(row=row, column=client_column).value = article.quantity
            
            return  # on sort dès qu'on a trouvé
    raise ValueError(
        f"L'article {article.name} du client {client} "
        f"n'a pas été trouvé dans le fichier {config.central_workbook}.\n"
        f"Un fichier a peut être été modifié.\n"
        f"Veuillez vérifier les fichiers et relancer le script." 
    )