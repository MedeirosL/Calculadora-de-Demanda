from openpyxl import Workbook
from openpyxl import load_workbook
from openpyxl.cell import Cell
from openpyxl.comments import Comment
 
demandaLidaP = []
demandaLidaFP = []
demandaLida = []
demandaContratada = [1,1,1,1,1,1,1,1,1,1,1,1]
consumoP = []
consumoFP = []
totalAnual=0
consumoFPValor=0
consumoPValor=0
demandaValor=0
ultrapassagemValor=0

wb = load_workbook(filename = 'data.xlsx',data_only=True)
wb.active
sheetranges=wb['Planilha1']
for i in range (2,14):
    #cedula = "I"+ str(i)
    demandaLidaFP.append(sheetranges['H'+str(i)].value)
    demandaLidaP.append(sheetranges['I'+str(i)].value)
    consumoFP.append(sheetranges['J'+str(i)].value)
    consumoP.append(sheetranges['K'+str(i)].value)
    if demandaLidaP[i-2]>=demandaLidaFP[i-2]:
        demandaLida.append(demandaLidaP[i-2])
    else:
        demandaLida.append(demandaLidaFP[i-2])
kW_Verde=sheetranges['C5'].value
kWhP_Verde=sheetranges['C3'].value
kWhFP_Verde=sheetranges['C4'].value


for j in range (0,12):
    demandaContratada[j]=60
    demandaValor += kW_Verde*demandaContratada[j]
    consumoFPValor += kWhFP_Verde*consumoFP[j]
    consumoPValor += kWhP_Verde*consumoP[j]    
    if demandaLida[j]>demandaContratada[j]*1.05:
        ultrapassagemValor+=(demandaLida[j]-demandaContratada[j])*kW_Verde*2
valorTotal = demandaValor + consumoFPValor + consumoPValor + ultrapassagemValor