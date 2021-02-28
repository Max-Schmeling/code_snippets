def csv_to_dict(filepath, has_headers=True, column_separator=";", encoding="utf8"):
    """ 
    Converts a .csv file into a python dictionary. With the column headers 
    being the dict-keys and the corresponding values of each column being 
    the dict-values as a single list each.
    """
    with open(filepath, "r", encoding=encoding) as file:
        lines = file.readlines()
        cells = [line.strip().split(column_separator) for line in lines]
        assert len(cells) >= 2
        d = {}
        for c in range(len(cells[0])):
            column = []
            for i, row in enumerate(cells, 1):
                if (has_headers and i > 1) or (not has_headers):
                    column.append(row[c])
            if has_headers:
                d[cells[0][c]] = column
            else:
                d[f"column {c+1}"] = column
        return d
