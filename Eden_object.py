import file_functions  as ff
import HUMS_objects    as ho
import functools       as ft
import pdb #pdb.set_trace()
import logging

from operator import itemgetter
from openpyxl import load_workbook

log = logging.getLogger(__name__)

class Eden(ho.Hums):
    def __init__(self, SN, files):
        ho.Hums.__init__(self, SN, files)
        
        #defining test results of eden
        #4 types of tets:  S = status, T = thresholds, R = tolerence ref, P = % ref
        self.waited_test_result['T_conso_veil']  = [215, 250]
        self.waited_test_result['T_conso_perio'] = [4250, 4500]
        self.waited_test_result['T_conso_stock'] = [10, 20]
        self.waited_test_result['T_compr_joint'] = [73.2, 74.2]
        self.waited_test_result['T_poids']       = [0, 1250]
        self.waited_test_result['T_conti_cable'] = [0, 25]
        self.waited_test_result['T_conti_capot'] = [0, 5]
        self.waited_test_result['T_long_cable']  = [340, 360]
        self.waited_test_result['R_pressure']    = 50
        self.waited_test_result['R_temperature'] = 1
        self.waited_test_result['R_humidity']    = 5
        self.waited_test_result['P_vibration']   = 10
        self.waited_test_result['R_choc']        = 1
        self.waited_test_result['P_choc']        = 10
        self.waited_test_result['T_choc_trans']  = 20

    def fill_pv(self, ws):
        wb         = {}
        wb[self.SN]= load_workbook(filename= self.pv)
        ws         = wb[self.SN].active
        ff.img_import(wb, self.SN)

        self.fill_eden(ws)
        self.test_eden(ws)

        wb[self.SN].save(filename = self.pv)


    def fill_eden(self, ws):
        #Conso and other tests
        self.writing_tester(ws, 'F107', 'conso_sleep', ' µA')
        self.writing_tester(ws, 'F108', 'conso_acq', ' µA')
        self.writing_tester(ws, 'F109', 'conso_stock', ' µA')
        self.writing_tester(ws, 'F111', 'Compression joints piles', ' mm')
        self.writing_tester(ws, 'F162', 'Poids', 'g')
        self.writing_tester(ws, 'F166', 'Continuite electrique capot', ' mΩ')
        self.writing_tester(ws, 'F165', 'Continuite electrique cable', ' mΩ')
        self.writing_tester(ws, 'F158', 'Longueur de cable', ' mm')
    

        #Temperature
        self.writing_tester(ws, 'F171', 'Temperature Hums 2', '°C', 1)
        self.writing_tester(ws, 'F172', 'Temperature Hums 3', '°C', 1)
        self.writing_tester(ws, 'F173', 'Temperature Hums 1', '°C', 1)

        self.writing_tester(ws, 'D171', 'Temperature Ref 2', '°C', 1)
        self.writing_tester(ws, 'D172', 'Temperature Ref 3', '°C', 1)
        self.writing_tester(ws, 'D173', 'Temperature Ref 1', '°C', 1)

        #Humidity
        self.writing_tester(ws, 'F175', 'Humidite Hums 2', '%', 1)
        self.writing_tester(ws, 'F176', 'Humidite Hums 3', '%', 1)
        self.writing_tester(ws, 'F177', 'Humidite Hums 1', '%', 1)

        self.writing_tester(ws, 'D175', 'Humidite Ref 2', '%', 1)
        self.writing_tester(ws, 'D176', 'Humidite Ref 3', '%', 1)
        self.writing_tester(ws, 'D177', 'Humidite Ref 1', '%', 1)

        #Pressure
        self.writing_tester(ws, 'F179', 'Pression Hums 2', ' hPa', 1)
        self.writing_tester(ws, 'F180', 'Pression Hums 3', ' hPa', 1)
        self.writing_tester(ws, 'F181', 'Pression Hums 1', ' hPa', 1)

        self.writing_tester(ws, 'D179', 'Pression Ref 2', ' hPa', 1)
        self.writing_tester(ws, 'D180', 'Pression Ref 3', ' hPa', 1)
        self.writing_tester(ws, 'D181', 'Pression Ref 1', ' hPa', 1)

        #Vibrations
        self.writing_tester(ws, 'F117', 'Resultat de la mesure de vibration sur X', ' grms', 2)
        self.writing_tester(ws, 'F118', 'Resultat de la mesure de vibration sur Y', ' grms', 2)
        self.writing_tester(ws, 'F119', 'Resultat de la mesure de vibration sur Z', ' grms', 2)

        self.writing_tester(ws, 'D117', 'Resultat de la valeur de reference de vibration sur X', ' grms', 2)
        self.writing_tester(ws, 'D118', 'Resultat de la valeur de reference de vibration sur Y', ' grms', 2)
        self.writing_tester(ws, 'D119', 'Resultat de la valeur de reference de vibration sur Z', ' grms', 2)

        #Chocs
        self.writing_tester(ws, 'F188', 'Valeur mesuree des chocs sur l\'axe X', ' g', 2)
        self.writing_tester(ws, 'D188', 'Valeur de reference des chocs sur l\'axe X', ' g', 2)
        self.writing_tester(ws, 'F189', 'Shock (transverse) axe Y', ' g', 2)
        self.writing_tester(ws, 'F190', 'Shock (transverse) axe Z', ' g', 2)
        self.writing_tester(ws, 'F194', 'Duree choc axe X (ms)', ' ms', 1)

        self.writing_tester(ws, 'F195', 'Valeur mesuree des chocs sur l\'axe Y', ' g', 2)
        self.writing_tester(ws, 'D195', 'Valeur de reference des chocs sur l\'axe Y', ' g', 2)
        self.writing_tester(ws, 'F196', 'Shock (transverse) axe X', ' g', 2)
        self.writing_tester(ws, 'F197', 'Shock (transverse) axe Z', ' g', 2)
        self.writing_tester(ws, 'F198', 'Duree choc axe Y (ms)', ' ms', 1)

        self.writing_tester(ws, 'F199', 'Valeur mesuree des chocs sur l\'axe Z', ' g', 2)
        self.writing_tester(ws, 'D199', 'Valeur de reference des chocs sur l\'axe Z', ' g', 2)
        self.writing_tester(ws, 'F200', 'Shock (transverse) axe X', ' g', 2)
        self.writing_tester(ws, 'F201', 'Shock (transverse) axe Y', ' g', 2)
        self.writing_tester(ws, 'F202', 'Duree choc axe Z (ms)', ' ms', 1)


    def test_eden(self, ws):
        #Thresholded values checked
        self.threshold_check(ws, 'T_conso_veil' , 'conso_sleep', 'G107')
        self.threshold_check(ws, 'T_conso_perio', 'conso_acq'  , 'G108')
        self.threshold_check(ws, 'T_conso_stock', 'conso_stock', 'G109')
        self.threshold_check(ws, 'T_compr_joint', 'Compression joints piles', 'G111')
        self.threshold_check(ws, 'T_poids', 'Poids', 'G162')
        self.threshold_check(ws, 'T_conti_capot', 'Continuite electrique capot', 'G166')
        self.threshold_check(ws, 'T_conti_cable', 'Continuite electrique cable', 'G165')
        self.threshold_check(ws, 'T_long_cable', 'Longueur de cable', 'G158')
        #Tolerenced values checked
        self.tolerence_check(ws, 'R_temperature', 'Temperature Hums 2', 'Temperature Ref 2', 'G171')
        self.tolerence_check(ws, 'R_temperature', 'Temperature Hums 3', 'Temperature Ref 3', 'G172')
        self.tolerence_check(ws, 'R_temperature', 'Temperature Hums 1', 'Temperature Ref 1', 'G173')
        self.tolerence_check(ws, 'R_humidity', 'Humidite Hums 2', 'Humidite Ref 2', 'G175')
        self.tolerence_check(ws, 'R_humidity', 'Humidite Hums 3', 'Humidite Ref 3', 'G176')
        self.tolerence_check(ws, 'R_humidity', 'Humidite Hums 1', 'Humidite Ref 1', 'G177')
        self.tolerence_check(ws, 'R_pressure', 'Pression Hums 2', 'Pression Ref 2', 'G179')
        self.tolerence_check(ws, 'R_pressure', 'Pression Hums 3', 'Pression Ref 3', 'G180')
        self.tolerence_check(ws, 'R_pressure', 'Pression Hums 1', 'Pression Ref 1', 'G181')
        self.tolerence_check(ws, 'P_vibration', 'Resultat de la mesure de vibration sur X', 'Resultat de la valeur de reference de vibration sur X', 'G117', 'PERCENT')
        self.tolerence_check(ws, 'P_vibration', 'Resultat de la mesure de vibration sur Y', 'Resultat de la valeur de reference de vibration sur Y', 'G118', 'PERCENT')
        self.tolerence_check(ws, 'P_vibration', 'Resultat de la mesure de vibration sur Z', 'Resultat de la valeur de reference de vibration sur Z', 'G119', 'PERCENT')
        
        self.tolerence_check(ws, 'R_choc', 'Valeur mesuree des chocs sur l\'axe X', 'Valeur de reference des chocs sur l\'axe X', 'G188', 'BOTH', 'P_choc')
        self.tolerence_check(ws, 'P_choc_trans', 'Shock (transverse) axe Y', 'Shock (transverse) axe X', 'G189', 'PERCENT')
        self.tolerence_check(ws, 'P_choc_trans', 'Shock (transverse) axe Z', 'Shock (transverse) axe X', 'G190', 'PERCENT')
        self.tolerence_check(ws, 'R_choc', 'Valeur mesuree des chocs sur l\'axe Y', 'Valeur de reference des chocs sur l\'axe Y', 'G195', 'BOTH', 'P_choc')
        self.tolerence_check(ws, 'P_choc_trans', 'Shock (transverse) axe X', 'Shock (transverse) axe Y', 'G196', 'PERCENT')
        self.tolerence_check(ws, 'P_choc_trans', 'Shock (transverse) axe Z', 'Shock (transverse) axe Y', 'G197', 'PERCENT')
        self.tolerence_check(ws, 'R_choc', 'Valeur mesuree des chocs sur l\'axe Z', 'Valeur de reference des chocs sur l\'axe Z', 'G199', 'BOTH', 'P_choc')
        self.tolerence_check(ws, 'P_choc_trans', 'Shock (transverse) axe X', 'Shock (transverse) axe Z', 'G200', 'PERCENT')
        self.tolerence_check(ws, 'P_choc_trans', 'Shock (transverse) axe Y', 'Shock (transverse) axe Z', 'G201', 'PERCENT')












