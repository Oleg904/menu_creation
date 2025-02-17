


def open_type_menu():
    text = pd.read_excel('tm2025-sm.xlsx', sheet_name='Лист1')
    # return text.tail(1)
    # return text.head(1)
    return text.columns.tolist()
    # return text.tail(1), text.head(1)


print(open_type_menu())