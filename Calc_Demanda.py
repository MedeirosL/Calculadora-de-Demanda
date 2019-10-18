from openpyxl import Workbook
from openpyxl import load_workbook
from openpyxl.cell import Cell
from openpyxl.comments import Comment
 
demandaLida = []
demandaContratada = []*12

wb = load_workbook(filename = 'data.xlsx',data_only=True)
wb.active
sheetranges=wb['Planilha1']
for i in range (2,14):
    cedula = "C"+ str(i)
    demandaLida.append(sheetranges[cedula].value)
    #print(sheetranges[cedula].value)
print (sheetranges['H14'].value)
sheetranges['B2'].value=1999
print (sheetranges['B2'].value)
wb.save('data.xlsx')
wb.close()
print (sheetranges['H14'].value)
#teste
#print(max(demandaLida))