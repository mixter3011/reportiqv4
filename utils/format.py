import numpy as np

def format_num(val):
    if isinstance(val, (int, float)) and not np.isnan(val):
        return round(val, 1)
    return val