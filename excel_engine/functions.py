def excel_sum(*args):
    total = 0
    for arg in args:
        if isinstance(arg, list):
            for row in arg:
                total += excel_sum(*row)
        else:
            try:
                total += float(arg)
            except (ValueError, TypeError):
                pass
    return total

def excel_transpose(arg):
    if not isinstance(arg, list):
        return [[arg]]
    
    if not arg:
        return [[]]
    
    # Assume all rows have same length
    rows = len(arg)
    cols = len(arg[0]) if rows > 0 else 0
    
    new_array = [[None for _ in range(rows)] for _ in range(cols)]
    for r in range(rows):
        for c in range(cols):
            new_array[c][r] = arg[r][c]
    return new_array

def excel_index(array, row_num, col_num=None):
    if not isinstance(array, list):
        if row_num == 1 and (col_num is None or col_num == 1):
            return array
        raise ValueError("#REF!")
    
    rows = len(array)
    cols = len(array[0]) if rows > 0 else 0
    
    row_num = int(row_num)
    if col_num is not None:
        col_num = int(col_num)
    
    if rows == 1:
        # 1xC row vector
        if col_num is None:
            if 1 <= row_num <= cols:
                return array[0][row_num-1]
        else:
            if 1 <= row_num <= rows and 1 <= col_num <= cols:
                return array[row_num-1][col_num-1]
    elif cols == 1:
        # Rx1 column vector
        if col_num is None:
            if 1 <= row_num <= rows:
                return array[row_num-1][0]
        else:
            if 1 <= row_num <= rows and 1 <= col_num <= cols:
                return array[row_num-1][col_num-1]
    else:
        # 2D array
        if col_num is not None:
            if 1 <= row_num <= rows and 1 <= col_num <= cols:
                return array[row_num-1][col_num-1]
        else:
            raise ValueError("#REF!")
            
    raise ValueError("#REF!")

FUNCTIONS = {
    "SUM": excel_sum,
    "TRANSPOSE": excel_transpose,
    "INDEX": excel_index
}
