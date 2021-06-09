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
        for x in range (0,self.frame.shape[0] - 1):
            cero = pd.to_datetime(('{} 00:00').format(self.frame.iloc[x,3]))
            for y in range (4,14):
                celda = self.frame.iloc[x,y]
                if type(celda) != type(patron):
                    self.frame.iloc[x,y] = cero
        
        return self.frame

    
    
    
if __name__ == '__main__':
    a = r'J:\Emma\14. Vulcano\RelojRRHH\Proyecto\Archivos de trabajo\A completar\Excel 397 - copia (3) - copia.xlsx'
    corrector = CorrectorExcel(abs_path_excel=a)
    frame = corrector.frame_return()
    frameCorregido = corrector.corrector_marcada()

