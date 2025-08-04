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