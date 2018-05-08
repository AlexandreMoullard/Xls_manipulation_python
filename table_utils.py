import logging
import pdb
import pandas as pd

from itertools   import filterfalse
from numpy import mean

def csv_row_search(searched_value, column_name, file, separator):
    rf = pd.read_csv(file, sep = separator)
    column_list = list(rf[column_name])
    line = 0
    for idx,value in enumerate(column_list):
        if value == searched_value:
            return idx
    return False

def csv_col_search(column_name, file, separator):
    rf = pd.read_csv(file, sep = separator)
    return rf.columns.get_loc(column_name)

def get_list_from_csv_row(file, row, strat_col, end_col):
    ls=[]
    list_range = range(strat_col, end_col+1)
    df = pd.read_csv(file)
    try:
        for cel in list_range:
            ls.append(df.iloc[row, cel])
        return ls
    except Exception:
        log.error('Error importing list between {a} and {b} on row {c} from file {d}'.format(a=strat_col, b=end_col, c=row, d=file))

def standard_dev(lst):
    moy = mean(lst)
    sum = 0
    for val in lst:
        sum += (val-moy)**2
    return (sum/len(lst))**(1/2)

def kick_if_noised(lst):
    std_dev    = standard_dev(lst)
    moy        = mean(lst)
    noised_val = 0
    if len(list(filter(lambda x: abs(x-moy) >= std_dev, lst))) > 3:
        return False
    else:
        lst = list(filterfalse(lambda x: abs(x-moy) >= std_dev, lst))
    return lst 

def filter_conso(lst, min_treshold, max_treshold, mode):
    filtered_lst    = []
    filtered_lst_a  = []
    filtered_lst_b  = []
    cnt_wrong_state = 0
    
    for value in lst:
        if min_treshold < value < max_treshold:
            filtered_lst.append(value)
            if mode == 'sleep' and cnt_wrong_state > 5 and not filtered_lst_a:
                filtered_lst_a = filtered_lst
                filtered_lst = []
            elif mode == 'sleep' and cnt_wrong_state > 5 and filtered_lst_a:
                filtered_lst_b = filtered_lst  
        else:
            cnt_wrong_state += 1
        
    if mode == 'acqui' and len(filtered_lst) >= 2:
        filtered_lst.pop(0)
        filtered_lst.pop()
    if mode == 'sleep' and len(filtered_lst_a) >= 2:
        filtered_lst_a.pop()
        filtered_lst_a.pop()
        filtered_lst = filtered_lst_a + filtered_lst_b

    if standard_dev(filtered_lst) >= 0.000005:
        filtered_lst = kick_if_noised(filtered_lst)
    
    return filtered_lst

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
