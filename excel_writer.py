from typing import List

from models import Article, Error
from error import print_missing_article, add_missing_article

from config import LOWEST_CLIENT_ROW, HIGHEST_CLIENT_ROW

def add_article(sheet, client, article: Article, errors: List[Error]):
    for row in range(LOWEST_CLIENT_ROW, HIGHEST_CLIENT_ROW):

        order_name = sheet.cell(row=row, column=1).value
        order_quantity = sheet.cell(row=row, column=11).value

        if order_name == article.name:
            if order_quantity is None:
                order_quantity = 0
            
            sheet.cell(row=row, column=11).value = order_quantity + article.quantity
            
            return  # on sort dès qu'on a trouvé
    print_missing_article(article)
    add_missing_article(client, article, errors)