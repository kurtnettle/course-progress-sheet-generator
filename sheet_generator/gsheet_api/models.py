from dataclasses import dataclass
from enum import StrEnum


@dataclass
class GridRange:
    startRowIndex: int
    endRowIndex: int
    startColumnIndex: int
    endColumnIndex: int


@dataclass
class RGBColor:
    red: float
    blue: float
    green: float


class MergeType(StrEnum):
    MERGE_ALL = "MERGE_ALL"
    MERGE_COLUMNS = "MERGE_COLUMNS"
    MERGE_ROWS = "MERGE_ROWS"
