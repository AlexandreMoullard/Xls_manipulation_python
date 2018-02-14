from openpyxl import Workbook
from openpyxl import load_workbook
import file_generator
import logging

def task_manager(file_0, file_1, file_2, file_3, starting_PN):
    listed_generated_files = file_generator.generate(file_0, file_1, starting_PN)
    wb = {}
    #print(listed_generated_files)
    #loop_count = 0
    
    double_value = file_generator.double_value_check('SN', file_1)
    print(double_value)
    if double_value is not False:
        logging.warning('Le produit %s apparait en double dans les resultats du banc de test' % (double_value)) 

    for SN in listed_generated_files:
        product = SN[0]
        file_name = SN[1]
        #print(product, file_name)
        wb[product] = load_workbook(filename= file_name)
        ws1 = wb[product].active

        ws1['D13'] = product #Writing SN
        ws1['D11'] = len(listed_generated_files) #Writing number of realesed products

        row = file_generator.csv_row_search(product, 'SN', file_1)
        print(row)

        wb[product].save(filename = file_name)
        
        #loop_count += 1

if __name__ == "__main__":
    file_path1 = "/home/alex/Bureau/Git_project/Files/Step_1_lot_1727.csv"
    template = "/home/alex/Bureau/Git_project/1000691.036.AE ADAMS PVAI.xlsx"
    PN = str(1000636)
    show_return = task_manager(template, file_path1,"","", PN)
    print(show_return)