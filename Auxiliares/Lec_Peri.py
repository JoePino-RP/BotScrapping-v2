import pandas as pd
from openpyxl import load_workbook

archivo = 'Status_Excel_Pruebas.xlsx'
df = pd.read_excel(archivo, sheet_name = 'Status')
#print(df)
ref = df.loc[df['Status']==0]

#
print(ref)
ref['Status'].iloc[0] = 1
print(ref['Status'].iloc[0])

print (ref)
print(df)

"""archivo = []

aux = []

for i in range (len(ref)):

      archivo = ref.iloc[i]


      aux.append(archivo)

print (aux[0])
busq = (len(aux))
print(type(aux))
aux [0][0]=1
dataframeaux = pd.DataFrame(data = aux)
print(type(dataframeaux))
print(dataframeaux['Status'], dataframeaux['Cargo'],sep="\n")"""
"""
with pd.ExcelWriter(archivo,mode ="w",engine='xlsxwriter') as writer:
      sheet_name = 'Status'
      dataframeaux.to_excel(writer, sheet_name=sheet_name, index=False)

      

      writer.save()"""

"""
aux.loc[0,'Status']=1
print (aux[0])
print("MEREQUETENGUE")
""""""
with pd.ExcelWriter(archivo, mode='w', engine='xlsxwriter') as writer:
      sheet_name = 'Promos'
      df.to_excel(writer, sheet_name=sheet_name, index=False)"""