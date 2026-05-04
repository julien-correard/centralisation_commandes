from models import Article

def clear_central_table(sheet, config):
    for row in range(config.lowest_central_row, config.highest_central_row + 1):
        for col in range(config.lowest_central_col, config.highest_central_col + 1, 2):
            sheet.cell(row=row, column=col).value = None
            
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
    for col in range(config.first_client_column, last_col + 1):
        client_name = sheet.cell(row=config.client_row, column=col).value
        if mapped_client == client_name:
            return col
    
    raise ValueError(
        f"Le client {client} "
        f"n'a pas pu être trouvé dans le fichier "
        f"{config.central_workbook}.\n"
        f"Veuillez vérifier le fichier {config.fichier_correspondance_clients} "
        f"et les fichiers .xlsx et relancer le script."
    )

def add_article(sheet, client, article: Article, client_map, config):

    client_column = get_client_column(sheet, client, client_map, config)
    
    for row in range(config.lowest_central_row, config.highest_central_row + 1):

        order_name = sheet.cell(row=row, column=1).value
        order_quantity = sheet.cell(row=row, column=client_column).value

        if order_name == article.name:
            
            sheet.cell(row=row, column=client_column).value = article.quantity
            
            return  # on sort dès qu'on a trouvé
        
    raise ValueError(
        f"L'article {article.name} du client {client} "
        f"n'a pas été trouvé dans le fichier {config.central_workbook}.\n"
        f"Un fichier a peut être été modifié.\n"
        f"Veuillez vérifier les fichiers et relancer le script." 
    )

def save_workbook(workbook, path):
    try:
        workbook.save(path)
    except Exception as e:
        raise ValueError(
            f"Impossible de sauvegarder le fichier Excel {path} :\n{e}\n"
            f"Veuillez vérifier que le fichier n'est pas ouvert et que vous avez les permissions nécessaires, puis relancez le script."
        )   