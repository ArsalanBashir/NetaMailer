import xlrd

#----------------------------------------------------------------------
def open_file(path):
    """
    Open and read an Excel file
    """
    book = xlrd.open_workbook(path)
    # get the first worksheet
    first_sheet = book.sheet_by_index(0)
    name = first_sheet.col_values(2, start_rowx=1)
    email = first_sheet.col_values(3, start_rowx=1)
    bhaght_log = zip(name, email)
    return bhaght_log
