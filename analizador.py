class Analizador:
    """
    Clase que viene a hacer de limpieza para el dataframe, se encarga de leer linea por linea
    e ir acomodando el dataFrame.
    
    El criterio actual de turnos es el siguiente:
        -1er Turno del dia de 00:00 a 08:00
        -2do turno del dia de 08:00 a 16:00 o 16:48 dependiendo del operario (rotativo o no)
        -3er turno del dia de 16:00 a 00:00


    """    
    
    def __init__(self,frameOriginal,frameEnAnalisis):
        self.frameOriginal = frameOriginal
        self.frameEnAnalisis = frameEnAnalisis
        
        
    def limpiador(self):
        """
        Funcion que itera por cada uno de las lineas y limpia.
        Se necesita que los registros comienzen un dia antes del dia de analisis.

        Returns
        -------
        TYPE
            DESCRIPTION.

        """
        import pandas as pd
        import datetime
        for renglon in range(len(self.frameEnAnalisis)):
            dia = self.frameEnAnalisis.iloc[renglon,3]
            ayer = dia - datetime.timedelta(days=1)
            
            turnoMañanaIngreso = pd.to_datetime(('{} 8:00').format(dia))
            turnoMañanaSalida = pd.to_datetime(('{} 16:48').format(dia))
            turnoMañanaSalidaSinComer = pd.to_datetime(('{} 16:00').format(dia))
            turnoTardeIngreso = pd.to_datetime(('{} 16:00').format(dia))
            turnoNocheIngreso = pd.to_datetime(('{} 00:00').format(dia))
            turnoNocheSalida = pd.to_datetime(('{} 8:00').format(dia))
            #turnoMañanaIngreso = pd.to_datetime(('{} 8:00').format(dia))
            cero = pd.to_datetime(('{} 00:00').format(dia))
            medioDia = pd.to_datetime(('{} 12:00').format(dia))
            
            #--------------------Turnos de ayer----------------------
            turnoMañanaIngresoAyer = pd.to_datetime(('{} 8:00').format(ayer))
            turnoMañanaSalidaAyer = pd.to_datetime(('{} 16:48').format(ayer))
            turnoMañanaSalidaSinComerAyer = pd.to_datetime(('{} 16:00').format(ayer))
            turnoTardeIngresoAyer = pd.to_datetime(('{} 16:00').format(ayer))
            turnoNocheIngresoAyer = pd.to_datetime(('{} 00:00').format(ayer))
            turnoNocheSalidaAyer = pd.to_datetime(('{} 8:00').format(ayer))
            #turnoMañanaIngreso = pd.to_datetime(('{} 8:00').format(dia))
            ceroAyer = pd.to_datetime(('{} 00:00').format(ayer))
            medioDiaAyer = pd.to_datetime(('{} 12:00').format(ayer))
            
            
            for posicion in range(4,14,2):# iteracion sobre las columnas del dataFrame, arrancando con el primer ingreso.
                
                if (self.frameEnAnalisis.iloc[renglon,posicion] >= turnoMañanaIngreso and self.frameEnAnalisis.iloc[renglon,posicion] < medioDia) and self.frameOriginal.iloc[renglon,posicion +1] > turnoTardeIngresoAyer:
                    # condiciones sobre si el primer registro del dataFrame para ver si pertenece al primer turno del dia (NOCHE)
                    #print('pasando 1',dia)
                    fechaIngreso = self.frameOriginal.iloc[renglon,posicion +1]                                       
                    self.frameEnAnalisis.iloc[renglon,posicion +1] =  self.frameEnAnalisis.iloc[renglon,posicion]
                    self.frameEnAnalisis.iloc[renglon,posicion] = fechaIngreso
                
                elif self.frameEnAnalisis.iloc[renglon,posicion] == self.frameEnAnalisis.iloc[renglon -1,posicion +1] :
                    #Correcion del dataFrame cuando hay solo un registro en 1 dia y pertenece al ultimo turno
                    #pone todo en cero esa linea y la deja como falta.
                    #print('pasando 2',dia)
                    if (self.frameEnAnalisis.iloc[renglon,posicion +1] and self.frameEnAnalisis.iloc[renglon,posicion +2]) != cero:
                        self.frameEnAnalisis.iloc[renglon,posicion] = self.frameEnAnalisis.iloc[renglon,posicion+ 1]
                        self.frameEnAnalisis.iloc[renglon,posicion +1] = self.frameEnAnalisis.iloc[renglon,posicion+ 2]
                    
                    elif renglon == (len(self.frameEnAnalisis) -1):
                        self.frameEnAnalisis.iloc[renglon,posicion] = self.frameEnAnalisis.iloc[renglon,posicion+ 1]
                        self.frameEnAnalisis.iloc[renglon,posicion +1] = self.frameEnAnalisis.iloc[renglon,posicion+ 2]
                        
                    break
                             
                    
                if (self.frameEnAnalisis.iloc[renglon,posicion] > medioDia and self.frameEnAnalisis.iloc[renglon,posicion] < turnoTardeIngreso) and self.frameEnAnalisis.iloc[renglon,posicion +1] == cero:
                    #Condicion para ver si pertenece al turno tarde y no hay mas registros en ese dia
                    #print('pasando 3',dia)
                    fechaSalida = self.frameOriginal.iloc[renglon +1,posicion]
                    self.frameEnAnalisis.iloc[renglon,posicion +1] =  fechaSalida
                    #self.frameEnAnalisis.iloc[renglon +1,posicion] = pd.to_datetime(('{} 00:00').format(dia))
                    
        
                    
                elif self.frameEnAnalisis.iloc[renglon,posicion] > turnoTardeIngreso and self.frameEnAnalisis.iloc[renglon,posicion +1] == cero:
                    #Condicion para ver si corresponde a un ingreso nocturno que hace horas extras y no hay mas
                    #registros en la linea.
                    #print('pasando 4',dia)
                    fechaIngreso = self.frameEnAnalisis.iloc[renglon,posicion]
                    
                    self.frameEnAnalisis.iloc[renglon +1,posicion +1] = self.frameEnAnalisis.iloc[renglon +1,posicion]
                    self.frameEnAnalisis.iloc[renglon +1,posicion] = fechaIngreso
                    
                    self.frameEnAnalisis.iloc[renglon,posicion] = pd.to_datetime(('{} 00:00').format(dia))
                    break
        
        return self.frameEnAnalisis
                
                
                
            
        