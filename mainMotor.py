import os
import logging.config
import traceback
import sys
import pandas as pd
import pyinputplus as pyip
import numpy as np
import win32com.client
import docx
import colorama
colorama.init()

from paths import empleados_text,pathExcelTemporal,nombreExcelTemporal,pathExcelInforme,pathTXT,areas,formaDePago
from analizador import Analizador,CalculadorHoras,informeNoFichadas,ingresoNoFichadas
from createDB import ManagerSQL
from queryes import queryConsultaEmpleados,insertRegistros,selectAll,selectSome,insertEmpleado,deleteEmpleado,actualizarEmpleado
from openpyxl import load_workbook
from time import sleep
from termcolor import colored
from clean_data import fechasDeCalculo


logging.config.fileConfig('logger.ini', disable_existing_loggers=False)
logger = logging.getLogger(__name__)


from paths import(nombreInformeNoFichadasWord,nombreInformeNoFichadasPDF,pathInformesNoFichadas,
pathInformesFaltasTardanzas,
nombreInformeFaltasTardanzasWord,
nombreInformeFaltasTardanzasPDF)
class Motor:
    def mainLoop(self):
        print('-'*100)
        print(' '*50,'Bienvenido',' '*50)
        print('-'*100)
    
    
        tareas = ['Ordenado de registros','Creación de informes','Gestion de Base de datos','Salir']
        tareasOrdenado = ['Limpieza de registros','Actualización de registros','Volver','Salir']
        tareasInformes = ['Ingreso de fechas','Volver','Salir']
        tareasBD = ['Insertar registro','Actualizar registro','Eliminar registro','Descargar','Volver','Salir']
        
        continuar = True
        
        
        while continuar:
            respuesta = pyip.inputMenu(tareas,prompt='¿Que desea hacer?\n',lettered=True)
            print('\n')
            continuarOrdenado = True
            continuarCreacion = True
            continuarGestionBD = True
            
            if respuesta == 'Ordenado de registros':
                while continuarOrdenado:
                    
                    ordenadoRespuesta = pyip.inputMenu(tareasOrdenado,prompt='\n¿Que desea hacer?\n',lettered=True)
                    print('\n')
                
                    if ordenadoRespuesta == 'Limpieza de registros':
                        fechaInicio,fechaFin,feriados,mediosDias = fechasDeCalculo()
                        frame = frameFichadas()
                        
                        if frame.empty:
                            print('No existen archivos que limpiar\n')
                        else:
                            legajos = frameAnalisisIndividual(frame,fechaInicio,fechaFin)
                            limpiezaDeRegistros(legajos, fechaInicio, fechaFin)               
                    elif ordenadoRespuesta == 'Actualización de registros':
                        actualizacionRegistros()                
                    elif ordenadoRespuesta == 'Volver':
                        print('Volviendo al PRIMER MENU')
                        continuarOrdenado = False                    
                    elif ordenadoRespuesta == 'Salir':
                        continuarOrdenado = False
                        continuar = False
            
            
            elif respuesta == 'Creación de informes':
                while continuarCreacion:
                    informesRespuesta = pyip.inputMenu(tareasInformes,prompt='\n¿Que desea hacer?\n',lettered=True)
                    print('\n')
                    if informesRespuesta == 'Ingreso de fechas':
                        fechaInicio,fechaFin,feriados,mediosDias = fechasDeCalculo()
                        frameCorregido = seleccionInformes(fechaInicio, fechaFin,feriados = feriados,mediosDias= mediosDias)
                        if frameCorregido.empty:
                            pass 
                        else:
                            informeFaltasTardanzas(frameCorregido,fechaInicio,fechaFin,
                                    feriados=feriados,medioDias = mediosDias)
                        continuarCreacion = True                
                    
                    elif informesRespuesta == 'Volver':
                        print('Volviendo al PRIMER MENU')
                        continuarCreacion = False
                        
                    elif informesRespuesta == 'Salir':
                        continuarCreacion = False
                        continuar = False
    
            
            
            elif respuesta == 'Gestion de Base de datos':
                while continuarGestionBD:
                    baseDeDatosRespuesta = pyip.inputMenu(tareasBD,prompt='\n¿Que desea hacer?\n',lettered=True)
                    print('\n')
                    ['Insertar registro','Actualizar registro','Eliminar registro','Descargar','Volver','Salir']
                    if baseDeDatosRespuesta == 'Insertar registro':
                        decision = False

                        while not decision:

                            legajo,nombre,apellido,area,pago = datosOperario(areas, formaDePago)
                            print(legajo,nombre,apellido,area,pago)
                            decision = repreguntar()

                        managerSQL = ManagerSQL()
                        insercionBDLegajos(managerSQL, legajo, nombre, apellido, area, pago, insertEmpleado)
                        continuarGestionBD = True                
                    
                    elif baseDeDatosRespuesta == 'Actualizar registro':
                        campos = ['LEG','APELLIDO','NOMBRE','AREA','TIPO_DE_PAGO']
                        valorAreas =[]
                        decision = False
                        while not decision:
                            legajo = pyip.inputInt(prompt='Ingrese el LEGAJO del empleado a actualizar',min=0)
                            campo = pyip.inputMenu(campos,prompt='Elija que campo va a actualizar\n',lettered=True)
                            valor = actualizarValor(campo)
                            decision = repreguntar()
                        managerSQL = ManagerSQL()
                        actualizaBDLegajos(managerSQL, legajo, campo, valor, actualizarEmpleado)
                        print('\n')
                        continuarGestionBD = True
                    
                    elif baseDeDatosRespuesta == 'Eliminar registro':
                        decision = False
                        while not decision:
                            legajo = pyip.inputInt(prompt='Ingrese el LEGAJO del empleado a eliminar',min=0)
                            decision = repreguntar()
                        print('\n')
                        managerSQL = ManagerSQL()
                        deleteBDLegajos(managerSQL, legajo, deleteEmpleado)     
                        continuarGestionBD = True
                        
                    elif baseDeDatosRespuesta == 'Descargar':
                        manager = ManagerSQL()
                        sql_conection = manager.conexion()
                        consultaEmpleados = pd.read_sql(queryConsultaEmpleados,sql_conection)
                        archivo = pyip.inputCustom(prompt='Ingrese el nombre que desea ponerle al EXCEL: \n',
                                customValidationFunc=validador)
                        print('\n')        
                        nombre = os.path.join(os.getcwd(),pathExcelInforme,archivo)
                        nombre = nombre+ '.xlsx'
                        consultaEmpleados = consultaEmpleados.sort_values(by=['LEG'])
                        consultaEmpleados.to_excel(nombre,sheet_name='Registros',index=False)
                        
                        continuarGestionBD = True
                    
                    elif baseDeDatosRespuesta == 'Volver':
                        print('Volviendo al PRIMER MENU')
                        continuarGestionBD = False
                        
                    elif baseDeDatosRespuesta == 'Salir':
                        continuarGestionBD = False
                        continuar = False
    
            
            elif respuesta == 'Salir':
                continuar= False


        
    
    
    
if __name__ == '__main__':
    try:
        motor = Motor()
        motor.mainLoop()
        sys.exit()
    except Exception:
        logger.error("excepcion desconocida: %s", traceback.format_exc())
        sleep(5)