from openpyxl import Workbook
#from openpyxl.compat import range
from openpyxl import load_workbook
import pandas as pd
import re

def generate(template_file_path, testbench1_file_path, first_PN):
    rf = pd.read_csv(testbench1_file_path)
    product_list = list(rf['SN'])
    number_of_products = len(product_list)
    generated_files_list = []
    wb = {}
    loop_count_i = 0
    patern = r"(\d{7})(\.)(\d{3})"
    other_patern = r"\d{7}"
    patern1 = r"\d{1}"
    patern2 = r"\d{2}"
    patern3 = r"\d{3}"

    #Vérification de first_PN, deux cas : 1000XXX.XXX ou 1000XXX resp patern et other_patern
    if re.search(patern, first_PN):
        m = re.search(patern, first_PN)
        first_PN0 = m.group(1) + m.group(2)
        first_PN1 = m.group(3)
    elif re.search(other_patern, first_PN):
        first_PN0 = first_PN
        first_PN1 = ""

    for product in product_list :
        if re.search(patern, first_PN):
            PN = first_PN0
            PNext = int(first_PN1) + loop_count_i
            
            #Completing number of digits from integer, needs to be XXX
            if re.match(patern3, str(PNext)):
                PNext = str(PNext)
            elif re.match(patern2, str(PNext)):
                PNext = "0"+str(PNext)
            else:
                PNext = "00"+str(PNext)
        
        elif re.search(other_patern, first_PN):
            PN = ""
            PNext = str(int(first_PN0) + loop_count_i)

        wb[product] = load_workbook(filename= template_file_path)
        dest_filename = str(PN) + PNext + '.AA' + '_'+ product +'_PVAI.xlsx'
        generated_files_list.append((product, dest_filename))
        """
        ws1 = wb[product].active
        ws1.title = product
        ws1['D13'] = product
        ws1['D11'] = number_of_products
        """
        wb[product].save(filename = dest_filename)
        
        loop_count_i += 1
    return generated_files_list

if __name__ == "__main__":
    file_path1 = "/home/alex/Bureau/Git_project/Files/Step_1_lot_1727.csv"
    template = "/home/alex/Bureau/Git_project/1000691.036.AE ADAMS PVAI.xlsx"
    PN = str(1000636)
    show_return = generate(template, file_path1, PN)
    print(show_return)