from dataclasses import dataclass, field
from typing import Optional, Pattern, List


@dataclass
class ColumnConfig:
    """Configuration for a single column in a subtable.
    
    The value_pattern regex determines what values are valid, including blank handling:
    - Use r".+" for required fields (no blanks)
    - Use r".*" for optional fields (blanks allowed)
    - Use r"^$|PATTERN" for "blank OR specific pattern"
    - Use r"(PATTERN)?" for optional pattern
    """
    column_letter: str
    header_pattern: Pattern[str]  # Compiled regex pattern for header
    value_pattern: Pattern[str]  # Compiled regex pattern for cell values (handles blank validation)


@dataclass
class SectionHeaderConfig:
    """Configuration for the section header detection."""
    pattern: Pattern[str]  # Compiled regex pattern for section header text
    start_column: str  # Column where header should start (e.g., 'A')
    is_merged: bool = False  # Whether the section header is in a merged cell
    merged_rows: Optional[int] = None  # Expected number of rows for merged cell
    merged_columns: Optional[int] = None  # Expected number of columns for merged cell


@dataclass
class RowValidationConfig:
    """Configuration for determining valid rows.
    
    Individual column requirements are handled by their value_pattern regex.
    This config only handles row-level validation that can't be expressed per-column.
    """
    minimum_filled_columns: int = 0  # Minimum number of non-empty columns for a valid row (0 = no minimum)


@dataclass
class SubtableSearchConfig:
    """Main configuration for searching and extracting a subtable.

    Stop conditions:
    - max_consecutive_invalid_rows: Stop after N consecutive invalid rows (includes blank rows) (0 = never)
    - max_consecutive_blank_rows: Stop after N consecutive blank rows specifically (0 = never)
    - stop_on_merged_cell: Stop when encountering a merged cell in tracked columns
    - end_pattern/end_pattern_column: Stop when pattern matches in specified column

    Note: Blank rows are considered a type of invalid row.
    """
    columns: List[ColumnConfig]
    row_validation: RowValidationConfig
    section_header: Optional[SectionHeaderConfig] = None  # Made optional - if None, search starts from row 1

    # Stop conditions based on consecutive row counts
    max_consecutive_invalid_rows: int = 0  # Stop after N consecutive invalid rows (includes blank) (0 = never)
    max_consecutive_blank_rows: int = 0  # Stop after N consecutive blank rows specifically (0 = never)

    # Other stop conditions
    stop_on_merged_cell: bool = False  # Stop when encountering a merged cell
    end_pattern: Optional[Pattern[str]] = None  # Optional pattern to match for custom end condition
    end_pattern_column: Optional[str] = None  # Column to check for end pattern

    # ADD THIS NEW FIELD - List of patterns to discover anywhere in header row
    discoverable_headers: Optional[List[Pattern[str]]] = None

    # Strict column validation - raise exception if unexpected columns found
    strict_columns: bool = False

    # Multi-subtable support
    extract_multiple: bool = False  # Enable extraction of multiple subtables with same config
    max_subtables: Optional[int] = None  # Maximum number of subtables to extract (None = unlimited)
    max_blank_rows_between_subtables: int = 50  # Max blank rows to search before giving up
    combine_subtables: bool = True  # Whether to combine into single DataFrame or return list