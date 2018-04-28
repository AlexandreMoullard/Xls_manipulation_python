import pyexcel      as pe
import pandas       as pd
import table_utils  as tu
import functools    as ft
import style_patch
import pdb

from openpyxl import load_workbook 
from operator import itemgetter

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
        #logging tbd here
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
    		#logging here		

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

        ws['F107'] = str(self.hums_attributs['conso_sleep']) + ' µA'
        ws['F108'] = str(self.hums_attributs['conso_acq']) + ' µA'
        ws['F109'] = str(self.hums_attributs['conso_stock']) + ' µA'
        ws['F111'] = str(self.hums_attributs['Compression joints piles']) + ' mm'
        ws['F137'] = str(self.hums_attributs['Poids']) + 'g'
        ws['F139'] = str(self.hums_attributs['Resultat du test gabarit'])
        ws['F140'] = str(self.hums_attributs['Resultat de la valeur de mesure du HUMS libre au milliohmetre']) + ' mΩ'
        ws['F141'] = str(self.hums_attributs['Resultat de la valeur de mesure du HUMS monte au milliohmetre']) + ' mΩ'

        #Temperature
        ws['F154'] = str(round(self.hums_attributs['Temperature Hums 2'], 1)) + '°C'
        ws['F155'] = str(round(self.hums_attributs['Temperature Hums 3'], 1)) + '°C'
        ws['F156'] = str(round(self.hums_attributs['Temperature Hums 1'], 1)) + '°C'

        ws['D154'] = str(round(self.hums_attributs['Temperature Ref 2'], 1)) + '°C'
        ws['D155'] = str(round(self.hums_attributs['Temperature Ref 3'], 1)) + '°C'
        ws['D156'] = str(round(self.hums_attributs['Temperature Ref 1'], 1)) + '°C'

        #Humidity
        ws['F158'] = str(round(self.hums_attributs['Humidite Hums 2'], 1)) + '%'
        ws['F159'] = str(round(self.hums_attributs['Humidite Hums 1'], 1)) + '%'
        ws['F160'] = str(round(self.hums_attributs['Humidite Hums 3'], 1)) + '%'

        ws['D158'] = str(round(self.hums_attributs['Humidite Ref 2'], 1)) + '%'
        ws['D159'] = str(round(self.hums_attributs['Humidite Ref 1'], 1)) + '%'
        ws['D160'] = str(round(self.hums_attributs['Humidite Ref 3'], 1)) + '%'

        #Vibrations
        ws['F166'] = str(round(self.hums_attributs['Resultat de la mesure de vibration sur X'], 2)) + ' grms'
        ws['F167'] = str(round(self.hums_attributs['Resultat de la mesure de vibration sur Y'], 2)) + ' grms'
        ws['F168'] = str(round(self.hums_attributs['Resultat de la mesure de vibration sur Z'], 2)) + ' grms'

        ws['D166'] = str(round(self.hums_attributs['Resultat de la valeur de reference de vibration sur X'], 2)) + ' grms'
        ws['D167'] = str(round(self.hums_attributs['Resultat de la valeur de reference de vibration sur Y'], 2)) + ' grms'
        ws['D168'] = str(round(self.hums_attributs['Resultat de la valeur de reference de vibration sur Z'], 2)) + ' grms'

        #Chocs
        ws['F170'] = str(round(self.hums_attributs['Ecart max de l\'analyse SRC X'], 2)) + '%'
        ws['F171'] = str(round(max(self.hums_attributs['Shock (transverse X) axe Y'],self.hums_attributs['Shock (transverse X) axe Z']), 2)) + '%'
        ws['F172'] = str(round(self.hums_attributs['Ecart max de l\'analyse SRC Y'], 2)) + '%'
        ws['F173'] = str(round(self.hums_attributs['Ecart max de l\'analyse SRC Z'], 2)) + '%'

        ws['D170'] = '-'
        ws['D171'] = '-'
        ws['D172'] = '-'
        ws['D173'] = '-'

        #Ecretage
        ws['D218'] = str(round(min(self.hums_attributs['Ecretage -20 axe X'], self.hums_attributs['Ecretage -20 axe Y'], self.hums_attributs['Ecretage -20 axe Z']), 1)) + ' g'
        ws['D219'] = str(round(min(self.hums_attributs['Ecretage 70 axe X'] , self.hums_attributs['Ecretage 70 axe Y'] , self.hums_attributs['Ecretage 70 axe Z'] ), 1)) + ' g'

        dic_ecrt_low  = {key:self.hums_attributs[key] for key in ['Ecretage -20 axe X', 'Ecretage -20 axe Y', 'Ecretage -20 axe Z']}
        dic_ecrt_high = {key:self.hums_attributs[key] for key in ['Ecretage 70 axe X' , 'Ecretage 70 axe Y' , 'Ecretage 70 axe Z'] }
        ws['F218']    = str(max(dic_ecrt_low.items(), key=itemgetter(1))[0][12:])
        ws['F219']    = str(max(dic_ecrt_high.items(), key=itemgetter(1))[0][12:])

        wb[self.SN].save(filename = self.pv)