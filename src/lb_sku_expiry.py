import openpyxl

def create_db(fname):
    """ Load the sku expiry from filename. """

    expiry_db = {}
    wb = openpyxl.load_workbook(fname)
    sheet = wb.active

    for row in sheet.iter_rows(min_row=2):
        values = [cell.value for cell in row]
        sku = str(values[0])
        expiry_db[sku] = values[2]

    return expiry_db
