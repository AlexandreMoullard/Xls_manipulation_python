import pytest         as pt
import HUMS_objects   as ho
import Adams_object   as ao
import Eden_object    as eo
import file_functions as ff
import style_patch
import pdb #pdb.set_trace()

from pathlib  import Path
from openpyxl import load_workbook 

def file_gen(template_file, tb1_file, tb2_file, accept_file, first_pn=1000636, first_real_pn=1000636):
    # Objects init
    data_files    = ho.Datafiles(template_file, tb1_file, tb2_file, accept_file)
    batch         = ho.Batch()
    batch.get_products(data_files.files())
    batch.product_type_id()

    #defining batch type & generate hums objects
    product_obj_list = []
    if batch.product_type == 'E':
        for p in batch.products:
            product_obj_list.append(eo.Eden(p, data_files.files()))
    elif batch.product_type == 'A':
        for p in batch.products:
            product_obj_list.append(ao.Adams(p, data_files.files()))

    # Running functions using files
    batch.generated_pn(first_pn)

    for idx,p in enumerate(product_obj_list):
        p.generate_pv(template_file, batch.pn[idx], batch.products)
        p.get_attributs_from_acceptance(data_files.files())
        p.get_consumption(data_files.files())

        # Workbook management
        wb = {}
        wb[p.SN] = load_workbook(filename= p.pv)
        ws       = wb[p.SN].active
        ff.img_import(wb, p.SN)

        # Running functions using workbooks
        p.batch_fill_pv(batch.products, ws)
        p.pv_header(first_pn, ws)

        #filling file
        p.fill_pv(ws)


if __name__=='__main__':
    template_file  = "tested_files/1000691.036.AE ADAMS PVAI.xlsx"
    tb1_file       = "donne_dirac/Step_1_lot_16 31.csv"
    tb2_file       = "tested_files/Step_2_lot_1727.csv"
    accept_file    = "donne_dirac/HUMS_ADAMS_LOT_17 06_20180427_153620.csv"
    first_pn       = '1000691_test'
    file_gen(template_file, tb1_file, tb2_file, accept_file, first_pn)
