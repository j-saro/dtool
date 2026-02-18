from dataclasses import dataclass
from enum import StrEnum, auto
import os

class Unit(StrEnum):
    CHARS = auto()
    WORDS = auto()

class Boundary(StrEnum):
    STRICT = auto()
    NEAREST_WORD = auto()
    NEAREST_SENTENCE = auto()

@dataclass
class Config:
    file_path: str
    output_path: str
    count: int
    unit: Unit = Unit.CHARS
    boundary: Boundary = Boundary.STRICT

    def __post_init__(self):
        if not os.path.isfile(self.file_path):
            raise FileNotFoundError(f"Input file not found: {self.file_path}")

        output_dir = os.path.dirname(self.output_path)
        if output_dir and not os.path.isdir(output_dir):
            os.makedirs(output_dir)

        if not isinstance(self.count, int) or self.count <= 0:
            raise ValueError("Count must be a positive integer.")

        if self.unit == Unit.WORDS and self.boundary == Boundary.STRICT:
            self.boundary = Boundary.NEAREST_WORD
