from typing import List

from models import Article, Error

from config import CENTRAL_WORKBOOK

def print_missing_article(missing_article: Article):
    print("!!! ATTENTION : l'article " + missing_article.name + " n'a pas été trouvé dans " + CENTRAL_WORKBOOK + " !!!")
    print("Une cellule a peut être été modifiée.")
    input("Appuyez sur Entrée...")

def add_missing_article(client, article: Article, errors: List[Error]):
    errors.append(Error(article=article, client=client))

def print_errors(errors: List[Error]):
    if errors:
        print("!!! CERTAINS ARTICLES N'ONT PAS PU ETRE AJOUTES !!!")
        for error in errors:
            msg = (
                f"{error.article.quantity} {error.article.name} pour le client {error.client} "
                f"n'ont pas pu être ajoutés à {CENTRAL_WORKBOOK} "
                f"car ils n'ont pas été trouvés."
            )
            print(msg)
    if not errors:
        print("Aucune erreur.")