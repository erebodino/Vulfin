#PATHS

areas = ['ARMADO','INYECCION','MOTORES','DEPOSITO','COMERCIAL','MECANIZADO',
         'ADMINISTRACION','METALURGICA','LOGISTICA','SOPLADO','INYECCION MOLIENDA',
         'MECANIZADO ROTATIVO', 'MOTORES ALUMINIO', 'MANGUERAS','ARMADO ROTATIVO'
         ]
formaDePago = ['Mensual','Jornal']
pathTXT ='Archivos de trabajo\TXTs'
empleados_text = r'J:\Emma\14. Vulcano\RelojRRHH\Proyecto\TXTs\siragusaCOMPLETO.txt'

pathExcelTemporal = 'Archivos de trabajo\A completar'
nombreExcelTemporal = 'Excel desde {} hasta {}.xlsx'



pathInformesNoFichadas = 'Archivos de trabajo\Informes\Informe No Fichadas'
nombreInformeNoFichadasWord = 'Informe No fichaje del {} al {}'
nombreInformeNoFichadasPDF = 'Informe No fichaje del {} al {}'


pathInformesFaltasTardanzas = 'Archivos de trabajo\Informes\Informe Tardanzas y faltas'
nombreInformeFaltasTardanzasWord = 'Informe faltas,tardanzas y retiros del {} al {} .docx'
nombreInformeFaltasTardanzasPDF = 'Informe faltas, tardanzas y retiros del {} al {} .pdf'


pathExcelInforme = 'Archivos de trabajo\Informes\Excel'


rotativosInyeccion = ['INYECCION','MECANIZADO ROTATIVO','MOTORES ALUMINIO','SOPLADO', 'MANGUERAS', 'ARMADO ROTATIVO']
rotativosSoplado = ['NARANJA']

valoresListaDesplegable = '"HS ENFERMEDAD,HS ACCIDENTE,FERIADO,FALLECIMIENTO FAMILIAR,LIC S GOCE,SUSPENSION,LIC POR PATERNIDAD,EXAMEN,FALTA FAMILIAR ENFERMO,\
RETIRO EN HS,HS NO TRABAJADAS,LLEG TARDE,FALTAS INJUSTIFICADAS,FALTA JUSTIFICADA,VACUNACION COVID"'

motivos = {'HS ENFERMEDAD':0,
           'HS ACCIDENTE':1,
           'FERIADO':2,
           'FALLECIMIENTO FAMILIAR':3,
           'LIC S GOCE':4,
           'SUSPENSION':5,
           'LIC POR PATERNIDAD':6,
           'EXAMEN':7,
           'FALTA FAMILIAR ENFERMO':8,
           'RETIRO EN HS':9,
           'HS NO TRABAJADAS':10,
           'LLEG TARDE':11,
           'FALTAS INJUSTIFICADAS':12,
           'FALTA JUSTIFICADA':13,
           'VACUNACION COVID':14}

toleranciaHoraria = 1

if __name__ == '__main__':
    pass

