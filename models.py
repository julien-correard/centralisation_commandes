from dataclasses import dataclass

@dataclass
class Article:
    name: str
    quantity: int

@dataclass # classe des articles non trouvés
class Error:
    article: Article
    client: str