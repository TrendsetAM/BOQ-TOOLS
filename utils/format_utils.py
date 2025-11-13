import re


def format_number_eu(val, decimal_places=2):
    """
    Format a number with point as thousands separator and comma as decimal separator (e.g., 1.234.567,89)
    
    Args:
        val: The value to format
        decimal_places: Number of decimal places (default 2)
    """
    try:
        if val is None or val == '' or (isinstance(val, float) and (val != val)):
            return ''
        
        # Convert to float, handling various input formats
        num = float(str(val).replace(' ', '').replace('\u202f', '').replace(',', '.'))
        
        # Format with specified decimal places
        s = f"{num:,.{decimal_places}f}"
        # s is like '1,234,567.89' -> want '1.234.567,89'
        s = s.replace(',', 'X').replace('.', ',').replace('X', '.')
        return s
    except Exception:
        return str(val)


def excel_column_letter(index: int) -> str:
    """
    Convert a zero-based column index into an Excel-style column label.

    Args:
        index: Zero-based column index (0 -> 'A').

    Returns:
        Excel column label corresponding to the supplied index.

    Notes:
        Column mapping indices in the application are zero-based; see
        `MappingGenerator._create_sheet_mapping` where indices are derived
        via enumerate.
    """
    if index < 0:
        raise ValueError("Column index must be non-negative")

    letters = []
    # Excel columns are base-26 with 'A' representing 1. Since our index
    # is zero-based, we adjust by +1 when computing each digit.
    while index >= 0:
        index, remainder = divmod(index, 26)
        letters.append(chr(ord('A') + remainder))
        index -= 1

    return ''.join(reversed(letters))