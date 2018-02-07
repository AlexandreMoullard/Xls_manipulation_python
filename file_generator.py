from openpyxl import Workbook
from openpyxl.compat import range
from openpyxl import load_workbook
import pandas as pd
#from openpyxl.cell import get_column_letter
#from numpy import column_stack, arange, pi, sin

file_path1 = "/home/alex/Dropbox/Echange et travaux/Xls Python/Git/Files/Step_1_lot_1727.csv"


def generate(file_path):
    rf = pd.read_csv(file_path)
    product_list = list(rf['SN'])
    #{'E00210', 'E00203'}
    number_of_products = len(product_list)
    print(number_of_products)
    print(product_list)

generate(file_path1)
    #wb = {}
"""
    for i in product_list :
        wb[i] = load_workbook(filename='template.xlsx')
        dest_filename = 'PVAI_'+ i + '.xlsx'
        ws1 = wb[i].active
        ws1.title = i
        ws1['B5'] = "resultat de" + i
        ws1['C3'] = 'hello'
        wb[i].save(filename = dest_filename)
        """
