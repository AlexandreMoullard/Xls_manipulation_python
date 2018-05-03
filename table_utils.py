import re
import pandas as pd
import logging
import pdb

from collections import Counter
from openpyxl    import drawing

def csv_row_search(searched_value, column_name, file, separator):
    rf = pd.read_csv(file, sep = separator)
    column_list = list(rf[column_name])
    line = 0
    for idx,value in enumerate(column_list):
        if value == searched_value:
            return idx
    return False

def tb1_check(file):
    #each sn needs to be present once
    rf = pd.read_csv(file)
    product_list = list(rf['SN'])
    return [item for item, count in Counter(product_list).items() if count > 1]
    
def tb2_check(filetb1, filetb2):
    #each sn from tb1 needs to be present twice
    cnt = Counter()
    rf1 = pd.read_csv(filetb1)
    rf2 = pd.read_csv(filetb2)
    tb1_product_list = list(set(rf1['SN'])) #deleting double
    tb2_product_list = list(rf2['SN'])
    for e in tb1_product_list:
        cnt[e] =  0
    for p in tb2_product_list:
        cnt[p] += 1
    for key in set(tb1_product_list + tb2_product_list):
        if cnt[key]  > 2:
            cnt[key] = 'seen {} times'.format(cnt[key])
        if cnt[key] == 0:
            cnt[key] = 'not seen'
        if cnt[key] == 1:
            cnt[key] = 'seen once'
        if cnt[key] == 2:
            del cnt[key]
    return dict(cnt)

def acceptance_check(filetb1, fileacc):
    #all roducts from tb1 need to be presente once
    rf1 = pd.read_csv(filetb1)
    rf2 = pd.read_csv(fileacc, sep = ';')
    tb1_product_list = set(rf1['SN'])
    acc_product_list = set(rf2['Numero de serie'])
    return list(tb1_product_list.difference(acc_product_list))

def img_import(workbook, product_name):
    #correction image problem (needs to be imported seperately)
    img = drawing.image.Image('images/srett_xls_image.png')
    img1= drawing.image.Image('images/srett_xls_image.png') #doesn't work with 1 variable
    workbook[product_name].active.add_image(img, 'A1')
    workbook[product_name].active.add_image(img1, 'A31')


def get_list_from_csv_row(file, row, strat_col, end_col):
    ls=[]
    #pdb.set_trace()
    list_range = range(strat_col, end_col)
    df = pd.read_csv(file)
    for cel in list_range:
        ls.append(df.iloc[row, cel])
    return ls
    
def generate_pn(first_pn, products):
    pn_patern = r"(\d{7})(\.)(\d{3})"
    pn_patern1 = r"\d{7}"
    tested_patern = r"\d{2}"
    tested_patern1 = r"\d{3}"
    result = []

    if re.search(pn_patern, first_pn):
        m = re.search(pn_patern, first_pn)
        first_PN0 = m.group(1) + m.group(2)
        first_PN1 = m.group(3)
    elif re.search(pn_patern1, first_pn):
        first_PN0 = first_pn
        first_PN1 = ""

    for ind,product in enumerate(products):
        if re.search(pn_patern, first_pn):
            PN = first_PN0
            PNext = str(int(first_PN1) + ind)
            
            #Completing number of digits from integer, needs to be XXX
            if re.match(tested_patern1, str(PNext)):
                PNext = str(PNext)
            elif re.match(tested_patern, str(PNext)):
                PNext = "0"+str(PNext)
            else:
                PNext = "00"+str(PNext)
        
        elif re.search(pn_patern1, first_pn):
            PN = ""
            PNext = str(int(first_PN0) + ind)

        result.append(str(PN) + str(PNext))
    
    return result   

def logging_manager():
    log = logging.getLogger()
    hdlr = logging.FileHandler('logs.txt', mode='a')
    formatter = logging.Formatter('%(asctime)s %(levelname)s %(message)s')
    hdlr.setFormatter(formatter)
    log.addHandler(hdlr) 
    log.setLevel(logging.WARNING)
    return log

if __name__ == "__main__":
    File = ""
    value = 'E00216'
    col = 'SN'
    show_return =double_value_check(col, File)
    print(show_return)