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

from paths import empleados_text,pathExcelTemporal,nombreExcelTemporal,pathExcelInforme,pathTXT,\
areas,formaDePago,rotativosInyeccion,rotativosSoplado,motivos
from analizador import Analizador,CalculadorHoras,informeNoFichadas,ingresoNoFichadas
from createDB import ManagerSQL
from queryes import (queryConsultaEmpleados,insertRegistros,selectAll,
                     selectSome,insertEmpleado,deleteEmpleado,
                     actualizarEmpleado,selectDeleteRegistro,updateRegistro)
from openpyxl import load_workbook
from time import sleep
from termcolor import colored
from datetime import timedelta
import datetime
import openpyxl
from openpyxl.styles import PatternFill
from openpyxl.worksheet.datavalidation import DataValidation

from paths import(nombreInformeNoFichadasWord,nombreInformeNoFichadasPDF,pathInformesNoFichadas,
pathInformesFaltasTardanzas,
nombreInformeFaltasTardanzasWord,
nombreInformeFaltasTardanzasPDF,
valoresListaDesplegable,toleranciaHoraria)

colorama.init()


from Vulcano import(updateRegistroQuery,
                    actualizaLinea,
                    coloreadorRegistroModificar,
                    edicionRegistros,
                    agregadoListaDesplegable,
                    fechasDeCalculo,
                    ingreso_egreso,
                    creacionFrameVacio,
                    empleadosFrame,
                    insercionBD,
                    insercionBDLegajos,
                    deleteBDLegajos,
                    actualizaBDLegajos,
                    frameFichadas,
                    logicaRotativos,
                    coloreadorExcel,
                    frameAnalisisIndividual,
                    limpiezaDeRegistros,
                    analizadorFramesCorregidos,
                    actualizacionRegistros,
                    validador,
                    calculosExtrasRotativos,
                    hojaTotalizadora,
                    agregadoColumnas,
                    retTarRotativos,
                    seleccionInformes,
                    cambioPorMotivos,
                    calculosAdicionalesTotalizados,
                    informeFaltasTardanzas,
                    escritorInformeFaltasTardanzas,
                    datosOperario,
                    repreguntar,
                    actualizarValor)

logging.config.fileConfig('logger.ini', disable_existing_loggers=False)
logger = logging.getLogger(__name__)




class Motor:
    def mainLoop(self):
        print('-'*100)
        print(' '*50,'Bienvenido',' '*50)
        print('-'*100)
    
    
        tareas = ['Ordenado de registros','Creación de informes','Gestion de Base de datos','Salir']
        tareasOrdenado = ['Limpieza de registros','Actualización de registros','Volver','Salir']
        tareasInformes = ['Ingreso de fechas','Volver','Salir']
        tareasBD = ['Agregar empleado','Actualizar empleado','Eliminar empleado','Modificar Registro','Descargar','Volver','Salir']
        
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
                        
                        decision = False
                        while not decision:
                            fechaInicio,fechaFin = fechasDeCalculo(completo=False)
                            decision = repreguntar()
                        frame = frameFichadas()
                        
                        if frame.empty:
                            noArchivos = colored('No existen archivos que limpiar\n','grey',on_color='on_red')
                            print(noArchivos)
                        else:
                            legajos = frameAnalisisIndividual(frame,fechaInicio,fechaFin)
                            if legajos.empty:
                                msgRechazos = colored('No existen archivos que limpiar, archivo guardado en Rechazos.\n','grey',on_color='on_red')
                                print(msgRechazos)
                                len_noMarca = limpiezaDeRegistros(legajos, fechaInicio, fechaFin) 
                            else:
                                msgPersistenErrores = colored('\nRegistros erroneos y duplicados eliminados, excel a completar creado.\n','grey',on_color='on_red')
                                len_noMarca = limpiezaDeRegistros(legajos, fechaInicio, fechaFin)
                                if len_noMarca != 0:
                                    print(msgPersistenErrores)
                            
                    elif ordenadoRespuesta == 'Actualización de registros':
                        
                        decision = False
                        while not decision:
                            fechaInicio,fechaFin= fechasDeCalculo(completo=False)
                            decision = repreguntar()
                        actualizacionRegistros(fechaInicio,fechaFin)
                        
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
                        decision = False
                        while not decision:
                            fechaInicio,fechaFin,feriados,mediosDias = fechasDeCalculo()
                            decision = repreguntar()
                        frameCorregido = seleccionInformes(fechaInicio, fechaFin,feriados = feriados,mediosDias= mediosDias)
                        if frameCorregido.empty:
                            pass #Internamente ya hay un msj
                        else:
                            empleados = informeFaltasTardanzas(frameCorregido,fechaInicio,fechaFin,
                                    feriados=feriados,medioDias = mediosDias)
                            
                            empleadosExtras = calculosExtrasRotativos(frameCorregido)
                            
                            calculosAdicionalesTotalizados(frameCorregido, fechaInicio, fechaFin, feriados, empleados,empleadosExtras)
                            
                            escritorInformeFaltasTardanzas(empleados, fechaInicio, fechaFin)
                            
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
                    ['Agregar empleado','Actualizar empleado','Eliminar empleado','Modificar registro','Descargar','Volver','Salir']
                    if baseDeDatosRespuesta == 'Agregar empleado':
                        
                        decision = False
                        while not decision:
                            legajo,nombre,apellido,area,pago = datosOperario(areas, formaDePago)
                            decision = repreguntar()

                        managerSQL = ManagerSQL()
                        insercionBDLegajos(managerSQL, legajo, nombre, apellido, area, pago, insertEmpleado)
                        continuarGestionBD = True                
                    
                    elif baseDeDatosRespuesta == 'Actualizar empleado':
                        
                        campos = ['LEG','APELLIDO','NOMBRE','AREA','TIPO_DE_PAGO']
                        decision = False
                        while not decision:
                            legajo = pyip.inputInt(prompt='Ingrese el LEGAJO del empleado a actualizar:\n',min=0)
                            campo = pyip.inputMenu(campos,prompt='Elija que campo va a actualizar:\n',lettered=True)
                            valor = actualizarValor(campo)
                            decision = repreguntar()
                        managerSQL = ManagerSQL()
                        actualizaBDLegajos(managerSQL, legajo, campo, valor, actualizarEmpleado)
                        print('\n')
                        continuarGestionBD = True
                    
                    elif baseDeDatosRespuesta == 'Eliminar empleado':
                        
                        decision = False
                        while not decision:
                            legajo = pyip.inputInt(prompt='Ingrese el LEGAJO del empleado a eliminar:\n',min=0)
                            decision = repreguntar()
                        print('\n')
                        managerSQL = ManagerSQL()
                        deleteBDLegajos(managerSQL, legajo, deleteEmpleado)     
                        continuarGestionBD = True
                        
                    elif baseDeDatosRespuesta == 'Modificar Registro':
                        
                        decision = False
                        while not decision:
                            legajo = pyip.inputInt(prompt='Ingrese el LEGAJO:\n',min=0)
                            fecha = pyip.inputDate('Ingrese la fecha a actualizar DD/MM/AAAA: ',formats=["%d/%m/%Y"])
                            decision = repreguntar()
                        print('\n')
                        managerSQL = ManagerSQL()
                        registro = edicionRegistros(managerSQL,selectDeleteRegistro, fecha, legajo)
                        if registro.empty:
                            print('\n No existen registros para esa fecha y legajo\n')
                        else:
                            cantidad = coloreadorRegistroModificar(registro, fecha)
                            campo,fechaHora = actualizaLinea(registro, cantidad)
                            updateRegistroQuery(managerSQL, campo, fechaHora, legajo, fecha)
                            registro = edicionRegistros(managerSQL,selectDeleteRegistro, fecha, legajo)
                            print('Nuevo registro: \n')
                            cantidad = coloreadorRegistroModificar(registro, fecha)                           


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
    frameParaVer = None
    try:
        motor = Motor()
        motor.mainLoop()
        sys.exit()
    except Exception:
        logger.error("excepcion desconocida: %s", traceback.format_exc())
        sleep(5)