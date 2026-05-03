from dataclasses import dataclass
from pathlib import Path

@dataclass
class Article:
    name: str
    quantity: int

@dataclass
class Cell:
    row: int
    col: int
    file: Path