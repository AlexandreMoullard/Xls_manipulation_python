import table_utils  as tu

log = tu.logging_manager()

class Mixin:

	def fill_adams(self, ws):
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


