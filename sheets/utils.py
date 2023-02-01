from typing import Tuple

def get_loc_from_coords(coords: Tuple[int, int]) -> str:
    col, row = coords
    if col < 1 or row < 1 or col > 9999 or row > 9999:
        raise ValueError("Invalid coordinates")

    col_name = ""
    while col > 0:
        col_name = chr((col - 1) % 26 + ord('A')) + col_name
        col = (col - 1) // 26
    return col_name.upper() + str(row)