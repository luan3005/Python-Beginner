# Python-Beginner
Problem to identify which excel is from 5 different sheets by reading the column names

-- The problem is that received 5 excel sheets in a disorderly manner I have to identify what sheet it is by opening the files and reading the last column names. Can anyone help me to reduce this work overload please?

import datetime as dt

import pandas as pd
import xlsxwriter

df_1_3ultimas = df_1.iloc[:, -6:]

df_2_2ultimas = df_2.iloc[:, -6:]

df_3_6ultimas = df_3.iloc[:, -6:]

df_4_4ultimas = df_4.iloc[:, -6:]

df_5_4ultimas = df_5.iloc[:, -6:]


lista_1 = list(df_1_3ultimas.columns.values)

lista_2 = list(df_2_2ultimas.columns.values)

lista_3 = list(df_3_6ultimas.columns.values)

lista_4 = list(df_4_4ultimas.columns.values)

lista_5 = list(df_5_4ultimas.columns.values)




writer_1 = pd.ExcelWriter(path_output+"\\"+"BG Lvl 2 BRA SHP QTD-NET v1_"+dt.datetime.now().strftime("%d-%m-%y")+"_"+".xlsx",engine="xlsxwriter")

writer_2 = pd.ExcelWriter(path_output+"\\"+"CDNAInventory_"+dt.datetime.now().strftime("%d%m%y")+"_"+".xlsx",engine="xlsxwriter")

writer_3 = pd.ExcelWriter(path_output+"\\"+"CDNATradeSellout_"+dt.datetime.now().strftime("%d%m%y")+"_"+".xlsx",engine="xlsxwriter")

writer_4 = pd.ExcelWriter(path_output+"\\"+"STH To agent flag_end user v1_"+dt.datetime.now().strftime("%d%m%y")+"_"+".xlsx",engine="xlsxwriter")

writer_5 = pd.ExcelWriter(path_output+"\\"+"Sellto w taxid_v1_"+dt.datetime.now().strftime("%d%m%y")+"_"+".xlsx",engine="xlsxwriter")


if (("Shipments Net USD" in lista_1[3]) & ("Shipments Net Local Currency" in lista_1[4]) & ("Shipments Product Units" in lista_1[5])):
    df_1.to_excel(writer_1)
    writer_1.save()
    
elif (("Inv SellFrom Raw Store No" in lista_1[4]) & ("Total Inventory Product Units" in lista_1[5])):
    df_2.to_excel(writer_1)
    writer_1.save()
    
elif (("SellThru NDP Local Currency" in lista_1[0]) & ("SellThru NDP USD" in lista_1[1]) & ("SellThru Product Units" in lista_1[2]) & ("SellTo NDP Local Currency" in lista_1[3]) & ("SellTo NDP USD" in lista_1[4]) & ("SellTo Product Units" in lista_1[5])):
    df_3.to_excel(writer_1)
    writer_1.save()
    
elif (("POS EU Raw Name" in lista_1[2]) & ("SellThru NDP Local Currency" in lista_1[3]) & ("SellThru NDP USD" in lista_1[4]) & ("SellThru Product Units" in lista_1[5])):
    df_4.to_excel(writer_1)
    writer_1.save()
    
elif (("POS EU Raw Name" in lista_1[2]) & ("SellTo NDP Local Currency" in lista_1[3]) & ("SellTo NDP USD" in lista_1[4]) & ("SellTo Product Units" in lista_1[5])):
    df_5.to_excel(writer_1)
    writer_1.save()
else:
    print("não encontrou em nenhuma coluna - tipo1")

if (("Shipments Net USD" in lista_2[3]) & ("Shipments Net Local Currency" in lista_2[4]) & ("Shipments Product Units" in lista_2[5])):
    df_1.to_excel(writer_2)
    writer_2.save()
    
elif (("Inv SellFrom Raw Store No" in lista_2[4]) & ("Total Inventory Product Units" in lista_2[5])):
    df_2.to_excel(writer_2)
    writer_2.save()
    
elif (("SellThru NDP Local Currency" in lista_2[0]) & ("SellThru NDP USD" in lista_2[1]) & ("SellThru Product Units" in lista_2[2]) & ("SellTo NDP Local Currency" in lista_2[3]) & ("SellTo NDP USD" in lista_2[4]) & ("SellTo Product Units" in lista_2[5])):
    df_3.to_excel(writer_2)
    writer_2.save()
    
elif (("POS EU Raw Name" in lista_2[2]) & ("SellThru NDP Local Currency" in lista_2[3]) & ("SellThru NDP USD" in lista_2[4]) & ("SellThru Product Units" in lista_2[5])):
    df_4.to_excel(writer_2)
    writer_2.save()
    
elif (("POS EU Raw Name" in lista_2[2]) & ("SellTo NDP Local Currency" in lista_2[3]) & ("SellTo NDP USD" in lista_2[4]) & ("SellTo Product Units" in lista_2[5])):
    df_5.to_excel(writer_2)
    writer_2.save()
else:
    print("não encontrou em nenhuma coluna - tipo2")

if (("Shipments Net USD" in lista_3[3]) & ("Shipments Net Local Currency" in lista_3[4]) & ("Shipments Product Units" in lista_3[5])):
    df_1.to_excel(writer_3)
    writer_3.save()
    
elif (("Inv SellFrom Raw Store No" in lista_3[4]) & ("Total Inventory Product Units" in lista_3[5])):
    df_2.to_excel(writer_3)
    writer_3.save()
    
elif (("SellThru NDP Local Currency" in lista_3[0]) & ("SellThru NDP USD" in lista_3[1]) & ("SellThru Product Units" in lista_3[2]) & ("SellTo NDP Local Currency" in lista_3[3]) & ("SellTo NDP USD" in lista_3[4]) & ("SellTo Product Units" in lista_3[5])):
    df_3.to_excel(writer_3)
    writer_3.save()
    
elif (("POS EU Raw Name" in lista_3[2]) & ("SellThru NDP Local Currency" in lista_3[3]) & ("SellThru NDP USD" in lista_3[4]) & ("SellThru Product Units" in lista_3[5])):
    df_4.to_excel(writer_3)
    writer_3.save()
    
elif (("POS EU Raw Name" in lista_3[2]) & ("SellTo NDP Local Currency" in lista_3[3]) & ("SellTo NDP USD" in lista_3[4]) & ("SellTo Product Units" in lista_3[5])):
    df_5.to_excel(writer_3)
    writer_3.save()
else:
    print("não encontrou em nenhuma coluna - tipo3")

if (("Shipments Net USD" in lista_4[3]) & ("Shipments Net Local Currency" in lista_4[4]) & ("Shipments Product Units" in lista_4[5])):
    df_1.to_excel(writer_4)
    writer_4.save()
    
elif (("Inv SellFrom Raw Store No" in lista_4[4]) & ("Total Inventory Product Units" in lista_4[5])):
    df_2.to_excel(writer_4)
    writer_4.save()
    
elif (("SellThru NDP Local Currency" in lista_4[0]) & ("SellThru NDP USD" in lista_4[1]) & ("SellThru Product Units" in lista_4[2]) & ("SellTo NDP Local Currency" in lista_4[3]) & ("SellTo NDP USD" in lista_4[4]) & ("SellTo Product Units" in lista_4[5])):
    df_3.to_excel(writer_4)
    writer_4.save()
    
elif (("POS EU Raw Name" in lista_4[2]) & ("SellThru NDP Local Currency" in lista_4[3]) & ("SellThru NDP USD" in lista_4[4]) & ("SellThru Product Units" in lista_4[5])):
    df_4.to_excel(writer_4)
    writer_4.save()
    
elif (("POS EU Raw Name" in lista_4[2]) & ("SellTo NDP Local Currency" in lista_4[3]) & ("SellTo NDP USD" in lista_4[4]) & ("SellTo Product Units" in lista_4[4])):
    df_5.to_excel(writer_4)
    writer_4.save()
else:
    print("não encontrou em nenhuma coluna - tipo4")
    
if (("Shipments Net USD" in lista_5[3]) & ("Shipments Net Local Currency" in lista_5[4]) & ("Shipments Product Units" in lista_5[5])):
    df_1.to_excel(writer_5)
    writer_5.save()
    
elif (("Inv SellFrom Raw Store No" in lista_5[4]) & ("Total Inventory Product Units" in lista_5[5])):
    df_2.to_excel(writer_5)
    writer_5.save()
    
elif (("SellThru NDP Local Currency" in lista_5[0]) & ("SellThru NDP USD" in lista_5[1]) & ("SellThru Product Units" in lista_5[2]) & ("SellTo NDP Local Currency" in lista_5[3]) & ("SellTo NDP USD" in lista_5[4]) & ("SellTo Product Units" in lista_5[5])):
    df_3.to_excel(writer_5)
    writer_5.save()
    
elif (("POS EU Raw Name" in lista_5[2]) & ("SellThru NDP Local Currency" in lista_5[3]) & ("SellThru NDP USD" in lista_5[4]) & ("SellThru Product Units" in lista_5[5])):
    df_4.to_excel(writer_5)
    writer_5.save()
    
elif (("POS EU Raw Name" in lista_5[2]) & ("SellTo NDP Local Currency" in lista_5[3]) & ("SellTo NDP USD" in lista_5[4]) & ("SellTo Product Units" in lista_5[5])):
    df_5.to_excel(writer_5)
    writer_5.save()
    
else:
    print("não encontrou em nenhuma coluna - tipo5")
