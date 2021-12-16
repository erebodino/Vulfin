from datetime import datetime, timedelta
import pandas as pd

class CorrectorExcel:
    
    def __init__(self,abs_path_excel : str):
        self.frame = pd.read_excel(abs_path_excel)
        self.corrector_fecha()
    
    def corrector_fecha(self):
        self.frame['Fecha'] = pd.to_datetime(self.frame['Fecha']).dt.date
        
    def frame_return(self) -> pd.DataFrame:
        return self.frame
        
    def corrector_marcada(self) -> pd.DataFrame:
        patron = self.frame.iloc[0,4]
        for x in range (0,self.frame.shape[0]):

            cero = pd.to_datetime(('{} 00:00').format(self.frame.iloc[x,3]))
            for y in range (4,14):
                celda = self.frame.iloc[x,y]
                if 'datetime.time' in str(type(celda)).split()[1]: # Chequea si en el type de la celda la misma es una hora (12:30)
                    try:
                        hora_corregida = pd.to_datetime(('{} {}').format(str(celda).split()[0],str(celda).split()[1]))
                    except:
                        hora_corregida = pd.to_datetime(('{} {}').format(self.frame.iloc[x,3],str(celda)))
                    self.frame.iloc[x,y] = hora_corregida
                
                elif 'datetime.datetime' in str(type(celda)).split()[1]: #Cheque si el type de la celda es el formato dd/mm/yyyy hh:mm:ss, si es asi rearma la fecha.
                    if str(self.frame.iloc[x,0]) == '782':
                        pass

                    if pd.to_datetime(str(celda).split()[0]) == pd.to_datetime(self.frame.iloc[x,3])- timedelta(days=1):
                        hora_corregida = pd.to_datetime(('{} {}').format(str(celda).split()[0],str(celda).split()[1]))
                    else:
                        hora_corregida = pd.to_datetime(('{} {}').format(str(celda).split()[0],str(celda).split()[1]))
                        #hora_corregida = pd.to_datetime(('{} {}').format(self.frame.iloc[x,3],str(celda).split()[1]))
                    self.frame.iloc[x,y] = hora_corregida

                elif type(celda) != type(patron): #En el resto de los casos los pasa a 00:00
                    self.frame.iloc[x,y] = cero

        
        return self.frame

    
    
    
if __name__ == '__main__':
    a = r'J:\Emma\14. Vulcano\RelojRRHH\Proyecto\Archivos de trabajo\A completar\Excel 397 - copia (3) - copia.xlsx'
    corrector = CorrectorExcel(abs_path_excel=a)
    frame = corrector.frame_return()
    frameCorregido = corrector.corrector_marcada()

