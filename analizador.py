import docx
import pandas as pd
import os
import win32com.client
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
        
        
    def limpiador(self,inyeccion):
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
            if inyeccion:
                dia = self.frameEnAnalisis.iloc[renglon,3]
                ayer = dia - datetime.timedelta(days=1)
                mañana = dia + datetime.timedelta(days=1)
                
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
                
                turnoTardeIngresoTomorrow = pd.to_datetime(('{} 16:00').format(mañana))
                turnoMañanaIngresoTomorrow = pd.to_datetime(('{} 8:00').format(mañana))
                turnoNocheIngresoTomorrow = pd.to_datetime(('{} 00:00').format(mañana))
                
                
                for posicion in range(4,14,2):# iteracion sobre las columnas del dataFrame, arrancando con el primer ingreso.
                    
                    if (self.frameEnAnalisis.iloc[renglon,posicion] >= turnoMañanaIngreso and self.frameEnAnalisis.iloc[renglon,posicion] < medioDia) and \
                        (self.frameEnAnalisis.iloc[renglon,posicion +1] > turnoTardeIngreso or self.frameEnAnalisis.iloc[renglon,posicion +1] == cero):
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
                                 
                        
                    if (self.frameEnAnalisis.iloc[renglon,posicion] > medioDia and self.frameEnAnalisis.iloc[renglon,posicion] < turnoTardeIngreso) and self.frameEnAnalisis.iloc[renglon,posicion +1] == cero and \
                        self.frameEnAnalisis.iloc[renglon +1 ,posicion] > turnoNocheIngresoTomorrow and \
                        self.frameEnAnalisis.iloc[renglon +1 ,posicion +1] <= turnoMañanaIngresoTomorrow :
                        #Condicion para ver si pertenece al turno tarde y no hay mas registros en ese dia
                        #print('pasando 3',dia)
                        fechaSalida = self.frameOriginal.iloc[renglon +2,posicion]
                        self.frameEnAnalisis.iloc[renglon,posicion +1] =  fechaSalida
                        self.frameEnAnalisis.iloc[renglon +1,posicion] =  self.frameEnAnalisis.iloc[renglon +1,posicion +1]
                        self.frameEnAnalisis.iloc[renglon +1,posicion +1] =  self.frameEnAnalisis.iloc[renglon +1,posicion +2]
                        
                        #self.frameEnAnalisis.iloc[renglon +1,posicion] = pd.to_datetime(('{} 00:00').format(dia))
                        
            
                        
                    elif self.frameEnAnalisis.iloc[renglon,posicion] >= turnoTardeIngreso and self.frameEnAnalisis.iloc[renglon,posicion +1] == cero and \
                        self.frameEnAnalisis.iloc[renglon +1,posicion] >= turnoMañanaIngresoTomorrow and self.frameEnAnalisis.iloc[renglon +1,posicion +1] >= turnoTardeIngresoTomorrow:
                        #Condicion para ver si corresponde a un ingreso nocturno que hace horas extras y no hay mas
                        #registros en la linea.
                        #print('pasando 4',dia)
                        fechaIngreso = self.frameOriginal.iloc[renglon +1,posicion]
                        self.frameEnAnalisis.iloc[renglon +1,posicion +2] = self.frameOriginal.iloc[renglon +2,posicion +1]                        
                        self.frameEnAnalisis.iloc[renglon +1,posicion +1] = self.frameOriginal.iloc[renglon +2,posicion]
                        self.frameEnAnalisis.iloc[renglon +1,posicion] = fechaIngreso                        
                        self.frameEnAnalisis.iloc[renglon,posicion] = pd.to_datetime(('{} 00:00').format(dia))
                        break
                    
                    # if  self.frameEnAnalisis.iloc[renglon,posicion] >= turnoMañanaIngreso and \
                    #     self.frameEnAnalisis.iloc[renglon,posicion +1] > turnoTardeIngreso and \
                    #     self.frameEnAnalisis.iloc[renglon -1,posicion +2] > turnoTardeIngresoAyer :
                    #     # condiciones sobre si el primer registro del dataFrame para ver si pertenece al primer turno del dia (NOCHE)
                    #     #print('pasando 1',dia)
                    #     fechaIngreso = self.frameEnAnalisis.iloc[renglon -1,posicion +2]                                       
                    #     self.frameEnAnalisis.iloc[renglon,posicion +1] =  self.frameEnAnalisis.iloc[renglon,posicion]
                    #     self.frameEnAnalisis.iloc[renglon,posicion] = fechaIngreso
                    #     self.frameEnAnalisis.iloc[renglon -1,posicion +2] = cero
                    #     break
            else:
                    dia = self.frameEnAnalisis.iloc[renglon,3]
                    ayer = dia - datetime.timedelta(days=1)
                    mañana = dia + datetime.timedelta(days=1)
                    margen = datetime.timedelta(minutes=5)
                    
                    turnoMañanaIngreso = pd.to_datetime(('{} 6:55').format(dia))
                    turnoMañanaSalida = pd.to_datetime(('{} 15:00').format(dia))
                    turnoMañanaSalidaSinComer = pd.to_datetime(('{} 15:00').format(dia))
                    turnoTardeIngreso = pd.to_datetime(('{} 15:00').format(dia))
                    turnoNocheIngreso = pd.to_datetime(('{} 23:00').format(dia))
                    turnoNocheSalida = pd.to_datetime(('{} 7:00').format(dia))
                    #turnoMañanaIngreso = pd.to_datetime(('{} 8:00').format(dia))
                    cero = pd.to_datetime(('{} 00:00').format(dia))
                    medioDia = pd.to_datetime(('{} 12:00').format(dia))
                    
                    #--------------------Turnos de ayer----------------------
                    turnoMañanaIngresoAyer = pd.to_datetime(('{} 7:00').format(ayer))
                    turnoMañanaSalidaAyer = pd.to_datetime(('{} 15:48').format(ayer))
                    turnoMañanaSalidaSinComerAyer = pd.to_datetime(('{} 15:00').format(ayer))
                    turnoTardeIngresoAyer = pd.to_datetime(('{} 15:00').format(ayer))
                    turnoNocheIngresoAyer = pd.to_datetime(('{} 23:00').format(ayer))
                    turnoNocheSalidaAyer = pd.to_datetime(('{} 7:00').format(ayer))
                    #turnoMañanaIngreso = pd.to_datetime(('{} 8:00').format(dia))
                    ceroAyer = pd.to_datetime(('{} 00:00').format(ayer))
                    medioDiaAyer = pd.to_datetime(('{} 12:00').format(ayer))
                    
                    #--------------------Turnos de mañana----------------------
                    turnoMañanaIngresoTomorrow = pd.to_datetime(('{} 7:00').format(mañana))
                    turnoMañanaSalidaTomorrow = pd.to_datetime(('{} 15:48').format(mañana))
                    turnoMañanaSalidaSinComerTomorrow = pd.to_datetime(('{} 15:00').format(mañana))
                    turnoTardeIngresoTomorrow = pd.to_datetime(('{} 15:00').format(mañana))
                    turnoNocheIngresoTomorrow = pd.to_datetime(('{} 23:00').format(mañana))
                    turnoNocheSalidaTomorrow = pd.to_datetime(('{} 7:00').format(mañana))
                    #turnoMañanaIngreso = pd.to_datetime(('{} 8:00').format(dia))
                    ceroTomorrow = pd.to_datetime(('{} 00:00').format(mañana))
                    medioDiaTomorrow = pd.to_datetime(('{} 12:00').format(mañana))
                    
                    
                    for posicion in range(4,12,2):# iteracion sobre las columnas del dataFrame, arrancando con el primer ingreso.

                        if (self.frameEnAnalisis.iloc[renglon,posicion] > cero and self.frameEnAnalisis.iloc[renglon,posicion] < turnoMañanaIngreso) and self.frameEnAnalisis.iloc[renglon,posicion +1] > medioDia:
                            # condiciones sobre si el primer registro del dataFrame para ver si pertenece al turno mañana y la salida es dsp del medio dia
                            #print('pasando 1',dia)
                            break
                        
                        elif (self.frameEnAnalisis.iloc[renglon,posicion] > medioDia and self.frameEnAnalisis.iloc[renglon,posicion] < turnoTardeIngreso) and self.frameEnAnalisis.iloc[renglon,posicion +1] > turnoTardeIngreso:
                            # verifica si el primer registro es entre el medio dia y el turno de ingreso de la tarde (posibles horas extras), y la salida es maor al ingreso del turno tarde.
                            #print('pasando 1',dia)
                            break
                        
                        elif self.frameEnAnalisis.iloc[renglon,posicion] > turnoTardeIngreso and self.frameEnAnalisis.iloc[renglon,posicion +1] == cero  and \
                            (self.frameOriginal.iloc[renglon +2,posicion] > turnoMañanaIngresoTomorrow and self.frameOriginal.iloc[renglon +2,posicion] < medioDiaTomorrow):
                            #Si el primer registro es del turno tarde y no hay mas registros lo lleva al dia de mañana y corre todo las columnas en 1.

                            self.frameEnAnalisis.iloc[renglon +1,posicion +2] = self.frameEnAnalisis.iloc[renglon +1,posicion +1]
                            self.frameEnAnalisis.iloc[renglon +1,posicion +1] = self.frameEnAnalisis.iloc[renglon +1,posicion]                                      
                            self.frameEnAnalisis.iloc[renglon +1,posicion] = self.frameEnAnalisis.iloc[renglon,posicion]
                            self.frameEnAnalisis.iloc[renglon,posicion] = cero
                            break
                                     
                            
                        elif  self.frameEnAnalisis.iloc[renglon,posicion] >= turnoMañanaIngreso and \
                              self.frameEnAnalisis.iloc[renglon,posicion +1] > turnoTardeIngreso and \
                              self.frameEnAnalisis.iloc[renglon -1,posicion +2] > turnoTardeIngresoAyer:
                            #Condicion para ver si pertenece al turno tarde y la fecha tiene que ser corrida en 1 posicion hacia abajo.
                                print('pasando 3',dia)
                                fechaIngreso = self.frameEnAnalisis.iloc[renglon -1,posicion +2]
                                self.frameEnAnalisis.iloc[renglon,posicion +2] = self.frameEnAnalisis.iloc[renglon,posicion +1]
                                self.frameEnAnalisis.iloc[renglon,posicion +1] = self.frameEnAnalisis.iloc[renglon,posicion]              
                                self.frameEnAnalisis.iloc[renglon,posicion] =  fechaIngreso
                                self.frameEnAnalisis.iloc[renglon -1,posicion +2] = ceroAyer
                                break


                            
                
                            
                        # elif self.frameEnAnalisis.iloc[renglon,posicion] > turnoTardeIngreso and self.frameEnAnalisis.iloc[renglon,posicion +1] == cero:
                        #     #Condicion para ver si corresponde a un ingreso nocturno que hace horas extras y no hay mas
                        #     #registros en la linea.
                        #     #print('pasando 4',dia)
                        #     fechaIngreso = self.frameEnAnalisis.iloc[renglon,posicion]
                            
                        #     self.frameEnAnalisis.iloc[renglon +1,posicion +1] = self.frameEnAnalisis.iloc[renglon +1,posicion]
                        #     self.frameEnAnalisis.iloc[renglon +1,posicion] = fechaIngreso
                            
                        #     self.frameEnAnalisis.iloc[renglon,posicion] = pd.to_datetime(('{} 00:00').format(dia))
                        #     break
                    
        
        return self.frameEnAnalisis

class CalculadorHoras:
    
    def horasTrabajadas(self,frame):
        import pandas as pd


        frame['H.Norm'] = 0
        frame['H. 50'] = 0
        frame['H. 100'] = 0
        
        for fila in range(len(frame)):
                horas_trabajadas = 0
                fecha = frame.iloc[fila,3]
                dia = frame.iloc[fila,2]
        
                for idx in range(4,14,2):
                    if frame.iloc[fila,idx + 1] == pd.to_datetime(('{} 00:00').format(fecha)) and frame.iloc[fila,idx] != pd.to_datetime(('{} 00:00').format(fecha)):
                        horas_trabajadas = 0
                        break
                    else:
                        horas_trabajadas += (frame.iloc[fila,idx + 1] - frame.iloc[fila,idx]).seconds
                
                horas_trabajadas = round(horas_trabajadas /3600,2)
                frame.iloc[fila,14] = horas_trabajadas

        return frame
    
    def horasExtrasTrabajadas(self,frame,frameEmpleados,feriados=None,toleranciaHoraria=1):
        
        indices = [frame[frame['H.Norm.'] != 0].index[x] for x in range(len(frame[frame['H.Norm.Emp'] != 0]))]
        fecha = frame.iloc[fila,3]
        dia = frame.iloc[fila,2]
        
        for indice in indices:
            minutosExtras100 = 0
            minutosExtras50 = 0
            fecha = frame.iloc[indice,3]
            ingreso = '08:00'
            salida = '16:00'
            horaSalidaSabado = '13:00'
            salidaSabado = pd.to_datetime(('{} {}').format(fecha,horaSalidaSabado))
            ceroHoy = pd.to_datetime(('{} {00:00}').format(fecha))
            horaIngreso = pd.to_datetime(pd.to_datetime(('{} {}').format(fecha,ingreso)) - datetime.timedelta(minutes=toleranciaHoraria))
            horaSalida = pd.to_datetime(pd.to_datetime(('{} {}').format(fecha,salida)) + datetime.timedelta(minutes=toleranciaHoraria))
            
            ingresoOperario = frame.iloc[indice,4]
            salidaOperario = 0
            
            for x in range(5,15,2):
                if frame.iloc[indice,x + 1] ==  ceroHoy:
                    salidaOperario = frame.iloc[indice,x + 1]
                    break
            
            if (dia == 'Sábado' and salidaOperario > salidaSabado):
                minutosExtras100 += ((salidaOperario - salidaSabado).second)/3600
                frame.iloc[indice,15] = frame.iloc[indice,15] - minutosExtras100
                break
            
            elif dia in feriados:
                minutosExtras100 += frame.iloc[indice,15]
                frame.iloc[indice,15] = 0
                frame.iloc[indice,17] = minutosExtras100
                break
            
            else:
                if salidaOperario > horaSalida:
                    minutosExtras50 += ((salidaOperario - horaSalida).second)/3600
                elif ingresoOperario < horaIngreso:
                    minutosExtras50 += ((horaIngreso - ingresoOperario).second)/3600
            
            frame.iloc[indice,15] = frame.iloc[indice,15] - minutosExtras50
            frame.iloc[indice,16] = minutosExtras50
        
        return frame
    
def informeNoFichadas(frame,fechaInicio,fechaFin,mediosDias,feriados):

      
        len_noMarca = len(frame[frame['H.Norm'] == 0])
        print(len_noMarca)
        doc = docx.Document()
        doc.add_heading(('Olvidos de fichaje entre {} y {}').format(fechaInicio,fechaFin), 0)
        c = doc.add_paragraph('Personal que no ha fichado: \n')
        for x in range(len_noMarca):
            legajo = frame.iloc[frame[frame['H.Norm'] == 0].index[x],0]
            nombre = frame.iloc[frame[frame['H.Norm'] == 0].index[x],1]
            dia = frame.iloc[frame[frame['H.Norm'] == 0].index[x],2]
            fecha = frame.iloc[frame[frame['H.Norm'] == 0].index[x],3]
            print(x,nombre)
            
            hora = frame.iloc[frame[frame['H.Norm'] == 0].index[x],4]
            hora_2 = frame.iloc[frame[frame['H.Norm'] == 0].index[x],6]
            if dia not in feriados and dia in mediosDias: 
                
                if hora_2 == pd.to_datetime(('{} 00:00').format(fecha)):                   
                    if hora >= pd.to_datetime(('{} 10:00').format(fecha)):
                        c.add_run(('\n\tEl dia {} el empleado {:10s} no ficho Ingreso.').format(dia,nombre))
                    else:
                        c.add_run(('\n\tEl dia {} el empleado {:10s} no ficho Salida.').format(dia,nombre))
    
                else:
                    if hora_2 >= pd.to_datetime(('{} 10:00').format(fecha)):
                        c.add_run(('\n\tEl dia {} el empleado {:10s} no ficho Re-ingreso.').format(dia,nombre))
                    else:   
                        c.add_run(('\n\tEl dia {} el empleado {:10s} no ficho Salida.').format(dia,nombre))
                        
            else:
                if dia not in feriados: 
                
                    if hora_2 == pd.to_datetime(('{} 00:00').format(fecha)):                   
                        if hora >= pd.to_datetime(('{} 15:00').format(fecha)):
                            c.add_run(('\n\tEl dia {} el empleado {:10s} no ficho Ingreso.').format(dia,nombre))
                        else:
                            c.add_run(('\n\tEl dia {} el empleado {:10s} no ficho Salida.').format(dia,nombre))
        
                    else:
                        if hora_2 >= pd.to_datetime(('{} 15:00').format(fecha)):
                            c.add_run(('\n\tEl dia {} el empleado {:10s} no ficho Re-ingreso.').format(dia,nombre))
                        else:   
                            c.add_run(('\n\tEl dia {} el empleado {:10s} no ficho Salida.').format(dia,nombre)) 
        fechaInicio = str(fechaInicio)
        fechaFin = str(fechaFin)
        fini = '\\Informe No fichaje del {} al {} .docx'.format(fechaInicio.replace('/','-'),fechaFin.replace('/','-'))
        pathToWord = os.getcwd() + fini
        fini_pdf = ('\\Informe No fichaje del {} al {} .pdf').format(fechaInicio.replace('/',' '),fechaFin.replace('/','-'))
        pathToPDF = os.getcwd() + fini_pdf
        doc.save(pathToWord)
        
        
        wordFilename = pathToWord
        
        pdfFilename = pathToPDF
        
        
        wdFormatPDF = 17 # Word's numeric code for PDFs.
        wordObj = win32com.client.Dispatch('Word.Application')
        docObj = wordObj.Documents.Open(wordFilename)
        docObj.SaveAs(pdfFilename, FileFormat=wdFormatPDF)
        docObj.Close()
        wordObj.Quit()
        os.remove(pathToWord)

                
                
            
                
        
        
    
                
                
                
            
        