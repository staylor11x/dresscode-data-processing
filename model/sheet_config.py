from dataclasses import dataclass, field
from typing import List, Tuple

@dataclass
class SheetConfig:
    level:str
    folder:str
    sheets: List[str]
    gender_cell: str
    year_cell: str


    gender_coords: Tuple[int, int] = field(init=False)
    year_coords: Tuple[int, int] = field(init=False)

    def __post_init__(self):
        self.gender_coords = self.excel_cell_to_iat(self.gender_cell)
        self.year_coords = self.excel_cell_to_iat(self.year_cell)

    @staticmethod
    def excel_cell_to_iat(cell_ref: str) -> Tuple[int, int]:
        """
        Convert Excel-style cell reference (e.g., 'A3') to (row, col) zero-indexed.
        """
        col_str = ''.join(filter(str.isalpha, cell_ref)).upper()
        row_str = ''.join(filter(str.isdigit, cell_ref))
        
        col = 0
        for char in col_str:
            col = col * 26 + (ord(char) - ord('A') + 1)
        col -= 1  # zero-based index
        
        row = int(row_str) - 1  # zero-based index

        return (row, col)