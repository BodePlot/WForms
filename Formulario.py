import xlsxwriter
import numpy as np

CreateFile=xlsxwriter.Workbook("Acta.xlsx")
CreateSheet=CreateFile.add_worksheet()

#Specify number of nodes

Nodes=int(input("Insert number of nodes"))

#Data
ID1=np.linspace(1,Nodes-1,Nodes-1)
ID2=np.linspace(2,Nodes,Nodes-1)

#Define Headers (same for all forms)
CreateSheet.write("A1","ID")
CreateSheet.write("B1","ID")
CreateSheet.write("C1","NC")
CreateSheet.write("D1","C")
CreateSheet.write("E1","ID")
CreateSheet.write("F1","NC")
CreateSheet.write("G1","C")
CreateSheet.write("H1","ID")
CreateSheet.write("I1","NC")
CreateSheet.write("J1","C")
CreateSheet.write("K1","NC")
CreateSheet.write("L1","C")
CreateSheet.write("M1","FOTO")
CreateSheet.write("N1","Observaciones")

for item in range(len(ID1)):
    CreateSheet.write(item+1,0,ID1[item])
    CreateSheet.write(item+1,1,ID2[item])
    CreateSheet.write(item+1,3,"-")

#Link last node with first one

CreateSheet.write(Nodes,0,Nodes)
CreateSheet.write(Nodes,1,1)


CreateFile.close()