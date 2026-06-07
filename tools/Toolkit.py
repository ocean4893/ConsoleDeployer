def col_name_to_index(col_name):
    """将列字母转换为数字，返回int:index-1"""
    col_name = col_name.upper()
    index = 0
    for char in col_name:
        index = index * 26 + (ord(char) - ord('A') + 1)
    return int(index - 1)