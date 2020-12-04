from openpyxl import Workbook
from openpyxl.worksheet.datavalidation import DataValidation
dv = DataValidation(type="list", formula1='"Dog,Cat,Bat"', allow_blank=True)
dv.error ='Your entry is not in the list'
dv.errorTitle = 'Invalid Entry'
dv.prompt = 'Please select from the list'
dv.promptTitle = 'List Selection'

wb = openpyxl.load_workbook(r'E:\MachinaData\Escritorio\prueba.xlsx')
ws = wb.active
ws.add_data_validation(dv)
dv.add('B1:B1048576')
wb.save(r'E:\MachinaData\Escritorio\prueba.xlsx')


