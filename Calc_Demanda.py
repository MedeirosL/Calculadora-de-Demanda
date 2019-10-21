from openpyxl import Workbook
from openpyxl import load_workbook
from openpyxl.cell import Cell
from openpyxl.comments import Comment

demandaLidaP = []
demandaLidaFP = []
demandaLida = []
demandaContratada = [30,30,30,30,30,30,30,30,30,30,30,30]
demandaContratadaMelhor = []
consumoP = []
consumoFP = []
totalAnual=0
consumoFPValor=0
consumoPValor=0
demandaValor=0
ultrapassagemValor=0
valorTotal = []
valorTotalMelhor=9999999.99
k=4
l=0
m=0

wb = load_workbook(filename = 'data.xlsx',data_only=True)
wb.active
sheetranges=wb['Planilha1']
for i in range (2,14):
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

for i in range (30,100):
    for j in range (0,12):
        demandaContratada[j]=i
        if demandaLida[j]>demandaContratada[j]:
            demandaValor += kW_Verde*demandaLida[j]
        else:
            demandaValor +=kW_Verde*demandaContratada[j]
        consumoFPValor += kWhFP_Verde*consumoFP[j]
        consumoPValor += kWhP_Verde*consumoP[j]    
        if demandaLida[j]>demandaContratada[j]*1.05:
            ultrapassagemValor+=(demandaLida[j]-demandaContratada[j])*kW_Verde*2
    valorTotal.append(demandaValor + consumoFPValor + consumoPValor + ultrapassagemValor)
    consumoFPValor=0
    consumoPValor=0
    demandaValor=0
    ultrapassagemValor=0
melhorContrato = 30 + int(valorTotal.index(min(valorTotal)))
print ("O melhor contrato HSV é de:", melhorContrato,"kW")
valorTotal.clear()

demandaContratada=[30,30,30,30,30,30,30,30,30,30,30,30]
for k in range (4,9):
    for m in range (0,60):
        for l in range (0,9):
            for i in range (30,100-m):
                for j in range (0,12):
                    if j<k:
                        demandaContratada[j+l-k+4]=demandaContratada[j+l-k+4]+1
                    if demandaLida[j]>demandaContratada[j]:
                        demandaValor += kW_Verde*demandaLida[j]
                    else:
                        demandaValor +=kW_Verde*demandaContratada[j]
                    consumoFPValor += kWhFP_Verde*consumoFP[j]
                    consumoPValor += kWhP_Verde*consumoP[j]    
                    if demandaLida[j]>demandaContratada[j]*1.05:
                        ultrapassagemValor+=(demandaLida[j]-demandaContratada[j])*kW_Verde*2
                valorTotal.append(demandaValor + consumoFPValor + consumoPValor + ultrapassagemValor)
                consumoFPValor=0
                consumoPValor=0
                demandaValor=0
                ultrapassagemValor=0
            demandaContratada=[30+m,30+m,30+m,30+m,30+m,30+m,30+m,30+m,30+m,30+m,30+m,30+m]
        l=0
indexHSVMulti=valorTotal.index(min(valorTotal))    
valorTotal.clear()   
demandaContratada=[30,30,30,30,30,30,30,30,30,30,30,30]
for k in range (4,9):
    for m in range (0,60):
        for l in range (0,9):
            for i in range (30,100-m):
                for j in range (0,12):
                    if j<k:
                        demandaContratada[j+l-k+4]=demandaContratada[j+l-k+4]+1
                    if demandaLida[j]>demandaContratada[j]:
                        demandaValor += kW_Verde*demandaLida[j]
                    else:
                        demandaValor +=kW_Verde*demandaContratada[j]
                    consumoFPValor += kWhFP_Verde*consumoFP[j]
                    consumoPValor += kWhP_Verde*consumoP[j]    
                    if demandaLida[j]>demandaContratada[j]*1.05:
                        ultrapassagemValor+=(demandaLida[j]-demandaContratada[j])*kW_Verde*2
                valorTotal.append(demandaValor + consumoFPValor + consumoPValor + ultrapassagemValor)
                if len(valorTotal)==indexHSVMulti+1:
                    print ("\nO melhor contrato HSV Multipatamar é:")
                    print (demandaContratada)
                consumoFPValor=0
                consumoPValor=0
                demandaValor=0
                ultrapassagemValor=0
            
            demandaContratada=[30+m,30+m,30+m,30+m,30+m,30+m,30+m,30+m,30+m,30+m,30+m,30+m]
            l+=1
        m+=1
        l=0