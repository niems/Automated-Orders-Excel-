import re
import openpyxl
from openpyxl.utils import column_index_from_string
from openpyxl.utils import get_column_letter
from ExcelModule import Spreadsheet

def main():
    order_sheet = Spreadsheet('spring_mobile_mod.xlsx')
    order_sheet.find_category_row()
    order_sheet.get_orders()
    order_sheet.print_orders()

    return None

main()
