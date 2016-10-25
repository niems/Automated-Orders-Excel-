import re
import openpyxl
from openpyxl.utils import column_index_from_string
from openpyxl.utils import get_column_letter

class Category(object):
    def __init__(self, name = None, row = -1, col = -1):
        self.name = name #name of current category
        self.row = row #current row of the category
        self.column = col #current column of the category


    def __str__(self):
        return 'Name:{0} | Row:{1} | Column:{2}'.format(self.name, self.row, self.column)


class Order(object):
    def __init__(self):
        self.first_name = None
        self.last_name = None
        self.company = None
        self.address_1 = None
        self.address_2 = None
        self.city = None
        self.state = None
        self.zip = None
        self.phone_number = '1111111111' #if no phone number is provided
        self.email = None
        self.product_and_qty = {} #format 'product' : quantity

    def __str__(self):
        info = str(self.first_name) + '\n' + str(self.last_name) + '\n' \
               + str(self.company) + '\n' + str(self.address_1) + '\n' \
               + str(self.address_2) + '\n' + str(self.city) + '\n' \
               + str(self.state) + '\n' + str(self.zip) + '\n' \
               + str(self.phone_number) + '\n' + str(self.email)

        for key, val in self.product_and_qty.items():
            info = info + '\n' + str(key) + '\n' + str(val) + '\n'

        return info
