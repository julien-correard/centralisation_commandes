from openpyxl import load_workbook

from dataclasses import dataclass

@dataclass
class Article:
    name: str
    quantity: int
    
LOWEST_CLIENT_ROW = 4
HIGHEST_CLIENT_ROW = 66

def add_article(sheet, client, article: Article):
    for row in range(LOWEST_CLIENT_ROW, HIGHEST_CLIENT_ROW):

        order_name = sheet.cell(row=row, column=1).value
        order_quantity = sheet.cell(row=row, column=11).value

        if order_name == article.name:
            if order_quantity is None:
                order_quantity = 0
            
            #print(order_quantity)
            #print(article.quantity)
            sheet.cell(row=row, column=11).value = order_quantity + article.quantity
            
            return  # on sort dès qu'on a trouvé
        
def read_articles(sheet):
    
    articles = []

    # Cherche les articles dont la quantité n'est pas nulle
    for name, quantity in sheet.iter_rows(
        min_row=LOWEST_CLIENT_ROW,
        max_row=HIGHEST_CLIENT_ROW+1,
        min_col=1,
        max_col=2,
        values_only=True
    ):
        if quantity is not None and quantity != "":
            try:
                quantity_int = int(quantity)
            except (TypeError, ValueError):
                continue  # ou gérer autrement

            current_article = Article(name, quantity_int)
            articles.append(current_article)
            #print(current_article)

    return articles
            
        
def main():
    
    # Charger le fichier
    client_workbook = load_workbook("Gamm vert.xlsx")
    client_sheet = client_workbook.active

    orders_workbook = load_workbook("Commandes.xlsx")
    orders_sheet = orders_workbook.active


    
            #add_article( orders_sheet, "Gamm vert", current_article)
            
    articles = read_articles(client_sheet)

    for article in articles:
        print(article)
        add_article(orders_sheet, "Gamm vert", article)
    
    orders_workbook.save("Commandes2.xlsx")
            
    print("Done")

if __name__ == "__main__":
    main()