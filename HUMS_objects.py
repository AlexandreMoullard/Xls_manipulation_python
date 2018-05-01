import pyexcel      as pe
import pandas       as pd
import table_utils  as tu
import functools    as ft
import style_patch
import pdb

from openpyxl import load_workbook 
from operator import itemgetter

log = tu.logging_manager()

class Datafiles:
    def __init__(self, file0, file1, file2, file3):
        self.file0 = file0 # Template
        self.file1 = file1 # TB1
        self.file2 = file2 # TB2
        self.file3 = file3 # Acceptance

    def files(self):
        return [self.file0, self.file1, self.file2, self.file3]

    def file_control(self):
        tb1_list_double  = tu.tb1_check(self.file1)
        tb2_dict_error   = tu.tb2_check(self.file1, self.file2)
        acc_list_missing = tu.acceptance_check(self.file1, self.file3)
        #logging file problems
        if tb1_list_double:
            log.error('Products {} appears twice or more in {}, this file should contain only uniq product result'.format(tb1_list_double, self.file1))
        if tb2_dict_error:
            log.error('The folowing problems: {} appears in {}, this file should contain each product twice'.format(tb2_dict_error, self.file2))
        if acc_list_missing:
            log.error('The folowing products are missing: {} in {}, this file should contain the same products than the 1rs testbench step'.format(acc_list_missing, self.file3))

        return [tb1_list_double, tb2_dict_error, acc_list_missing]

class Batch:
    def __init__(self):
        self.products     = []
        self.pn           = []
        self.product_type = ''

    def get_products(self, files):
        rf = pd.read_csv(files[1])
        self.products = list(rf['SN'])

    def generated_pn(self, first_pn):
        self.pn = tu.generate_pn(first_pn, self.products)

    def product_type_id(self):
    	batch = set()
    	for p in self.products:
    			batch.add(str(p[0]))
    	if len(batch) == 1:
    		self.product_type = str(p[0])
    	else:
    		self.product_type = 'inconsistante batch'
    		log.error('The batch should contain only one type of products (EDEN or ADAMS), actual batch: {}'.format(self.products))		

class Hums:
    def __init__(self, SN, files):
        self.SN                  = SN
        self.hums_attributs      = {}
        self.test_bench_1_row    = tu.csv_row_search(self.SN, 'SN', files[1], ',')
        self.test_bench_2_row    = tu.csv_row_search(self.SN, 'SN', files[2], ',')
        self.product_row         = tu.csv_row_search(self.SN, 'Numero de serie', files[3], ';')
        self.pv                  = ''


    def get_attributs_from_acceptance(self, files):
        data_dict = pe.get_dict(file_name= files[3], delimiter=';')
        for key, value in data_dict.items():
            self.hums_attributs[key] = value[self.product_row]

    def get_consumption(self, files):
        #for sleep mode the EDEN consumption columns are: 50 to 57 + 82 to 89, taken with 1 column security margin does:
        #for sleep mode the ADAMS consumption columns are identical but - 2 columns
        css1 = 50; ces1 = 55; css2 = 83; ces2=89
        csa  = 59; cea  = 79
        csst = 99; cest = 109

        if self.SN[0] == 'A':
        	css1 = css1-2; ces1 = ces1-2 ; css2 = css2-2; ces2 = ces2-2
        	csa  = csa -2; cea  = cea -2
        	csst = csst-2; cest = cest-2

        list_conso_sleep = tu.get_list_from_csv_row(files[1], self.test_bench_1_row, css1, ces1) + tu.get_list_from_csv_row(files[1], self.test_bench_1_row, css2, ces2)
        list_conso_acq   = tu.get_list_from_csv_row(files[1], self.test_bench_1_row, csa, cea) 
        list_conso_stock = tu.get_list_from_csv_row(files[1], self.test_bench_1_row, csst, cest)
        
        #mean calculation:
        conso_sleep_raw  = ft.reduce(lambda x, y: x + y, list_conso_sleep)/len(list_conso_sleep)
        conso_acq_raw    = ft.reduce(lambda x, y: x + y, list_conso_acq)  /len(list_conso_acq)
        conso_stock_raw  = ft.reduce(lambda x, y: x + y, list_conso_stock)/len(list_conso_stock)
        
        #converting in µA: 
        self.hums_attributs['conso_sleep'] = round(conso_sleep_raw*10**6, 1)
        self.hums_attributs['conso_acq']   = round(conso_acq_raw*10**6  , 1)
        self.hums_attributs['conso_stock'] = round(conso_stock_raw*10**6, 1)
    
    def generate_pv(self, template, PN, batch):
        wb          = {}
        wb[self.SN] = load_workbook(filename= template)
        ws          = wb[self.SN].active
        tu.img_import(wb, self.SN)
        self.pv     =  PN  + '.AA_' + self.SN + '_PVAI.xlsx'
        self.batch_fill_pv(batch, ws)
        self.pv_header(PN, ws)

        wb[self.SN].save(filename = self.pv) 

    def batch_fill_pv(self, batch, worksheet):
        pv_batch_columns = ['D','E','F','G', 'H']
        pv_batch_lines   = range(13, 17)
        
        worksheet['D11'] = len(batch)
        for ind_col,let in enumerate(pv_batch_columns):
            for ind_line,num in enumerate(pv_batch_lines):
                if (ind_col*4 + ind_line) < len(batch):
                    worksheet[let + str(num)] = batch[ind_col*4 + ind_line]
    
    def pv_header(self, pn, worksheet):
        worksheet['D1']  = 'PROCES VERVAL D’ACCEPTATION\nN°' + pn
        worksheet['D31'] = 'PROCES VERVAL D’ACCEPTATION INDIVIDUEL\nN°' + pn
        worksheet['D38'] = 'NUMERO DE SERIE : ' + self.SN


class Eden(Hums):
    def __init__(self, SN, files):
        Hums.__init__(self, SN, files)
    def fill_pv(self, worksheet):
        pass


class Adams(Hums):
    def __init__(self, SN, files):
        Hums.__init__(self, SN, files)

    def fill_pv(self, ws):
        wb         = {}
        wb[self.SN]= load_workbook(filename= self.pv)
        ws         = wb[self.SN].active
        tu.img_import(wb, self.SN)

        #Conso and other tests
        self.writing_tester(ws, 'F107', 'conso_sleep', ' µA')
        self.writing_tester(ws, 'F108', 'conso_acq', ' µA')
        self.writing_tester(ws, 'F109', 'conso_stock', ' µA')
        self.writing_tester(ws, 'F111', 'Compression joints piles', ' mm')
        self.writing_tester(ws, 'F137', 'Poids', 'g')
        self.writing_tester(ws, 'F139', 'Resultat du test gabarit')
        self.writing_tester(ws, 'F140', 'Resultat de la valeur de mesure du HUMS libre au milliohmetre', ' mΩ')
        self.writing_tester(ws, 'F141', 'Resultat de la valeur de mesure du HUMS monte au milliohmetre', ' mΩ')        

        #Temperature
        self.writing_tester(ws, 'F154', 'Temperature Hums 2', '°C', 1)
        self.writing_tester(ws, 'F155', 'Temperature Hums 3', '°C', 1)
        self.writing_tester(ws, 'F156', 'Temperature Hums 1', '°C', 1)

        self.writing_tester(ws, 'D154', 'Temperature Ref 2', '°C', 1)
        self.writing_tester(ws, 'D155', 'Temperature Ref 3', '°C', 1)
        self.writing_tester(ws, 'D156', 'Temperature Ref 1', '°C', 1)

        #Humidity
        self.writing_tester(ws, 'F158', 'Humidite Hums 2', '%', 1)
        self.writing_tester(ws, 'F159', 'Humidite Hums 3', '%', 1)
        self.writing_tester(ws, 'F160', 'Humidite Hums 1', '%', 1)

        self.writing_tester(ws, 'D158', 'Humidite Ref 2', '%', 1)
        self.writing_tester(ws, 'D159', 'Humidite Ref 3', '%', 1)
        self.writing_tester(ws, 'D160', 'Humidite Ref 1', '%', 1)

        #Vibrations
        self.writing_tester(ws, 'F166', 'Resultat de la mesure de vibration sur X', ' grms', 2)
        self.writing_tester(ws, 'F167', 'Resultat de la mesure de vibration sur Y', ' grms', 2)
        self.writing_tester(ws, 'F168', 'Resultat de la mesure de vibration sur Z', ' grms', 2)

        self.writing_tester(ws, 'D166', 'Resultat de la valeur de reference de vibration sur X', ' grms', 2)
        self.writing_tester(ws, 'D167', 'Resultat de la valeur de reference de vibration sur Y', ' grms', 2)
        self.writing_tester(ws, 'D168', 'Resultat de la valeur de reference de vibration sur Z', ' grms', 2)

        #Chocs
        self.writing_tester(ws, 'F170', 'Ecart max de l\'analyse SRC X', '%', 2)
        self.writing_tester(ws, 'F172', 'Ecart max de l\'analyse SRC Y', '%', 2)
        self.writing_tester(ws, 'F173', 'Ecart max de l\'analyse SRC Z', '%', 2)

        ws['D170'] = '-'
        ws['D171'] = '-'
        ws['D172'] = '-'
        ws['D173'] = '-'
        
        try:
            max_transverse = max(self.hums_attributs['Shock (transverse X) axe Y'],self.hums_attributs['Shock (transverse X) axe Z'])    
            ws['F171'] = str(round(max_transverse, 2)) + '%'
        except Exception:
            max_transverse = 0
            log.error('Max transverse calculation error on product: {}'.format(self.SN))

        
        #Ecretage
        try:
            ecret_min_low = min(self.hums_attributs['Ecretage -20 axe X'], self.hums_attributs['Ecretage -20 axe Y'], self.hums_attributs['Ecretage -20 axe Z'])
            ecret_min_hig = min(self.hums_attributs['Ecretage 70 axe X'] , self.hums_attributs['Ecretage 70 axe Y'] , self.hums_attributs['Ecretage 70 axe Z'] )
            ws['D218'] = str(round(ecret_min_low, 1)) + ' g'
            ws['D219'] = str(round(ecret_min_hig, 1)) + ' g'
        except Exception:
            ecret_min_low = 0
            ecret_min_hig = 0
            log.error('Ecretage minimal calculation error on product: {}'.format(self.SN))
        
        #Retrieving axis of mininimum ecretage
        if ecret_min_low and ecret_min_hig:
            dic_ecrt_low  = {key:self.hums_attributs[key] for key in ['Ecretage -20 axe X', 'Ecretage -20 axe Y', 'Ecretage -20 axe Z']}
            dic_ecrt_high = {key:self.hums_attributs[key] for key in ['Ecretage 70 axe X' , 'Ecretage 70 axe Y' , 'Ecretage 70 axe Z'] }
            ws['F218']    = str(max(dic_ecrt_low.items(),  key=itemgetter(1))[0][13:])
            ws['F219']    = str(max(dic_ecrt_high.items(), key=itemgetter(1))[0][12:])

        wb[self.SN].save(filename = self.pv)

    def writing_tester(self, worksheet, cell_value, attribut, unit='', rounding=0):
        try :
            if self.hums_attributs[attribut]:
                if rounding:
                    worksheet[cell_value] = str(round(self.hums_attributs[attribut], rounding)) + unit
                else:
                    worksheet[cell_value] = str(self.hums_attributs[attribut]) + unit
            else:
                log.error('No value found for: {} on product : {}'.format(attribut, self.SN))
        except Exception:
            log.error('Error during calculation of: {} on product : {}'.format(attribut, self.SN))
            raise