from openpyxl import Workbook
# adapted from the openpyxl docs:
# http://openpyxl.readthedocs.io/en/default/usage.html#read-an-existing-workbook
from openpyxl.compat import range
#from openpyxl.cell import get_column_letter
#from numpy import column_stack, arange, pi, sin

Product_list = {'E00210', 'E00203'}
Number_Of_Products = len(Product_list)
wb = {}

for i in Product_list :
	wb[i] = Workbook()
	dest_filename = 'PVAI_'+ i + '.xlsx'
	ws1 = wb[i].active
	ws1.title = i
	ws1['B5'] = "resultat de" + i
	ws1['C3'] = 'hello'
	wb[i].save(filename = dest_filename)