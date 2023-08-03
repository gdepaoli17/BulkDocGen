import os,sys
import xlwings as xw
from docxtpl import DocxTemplate
import pandas as pd




# wb = xw.Book.caller()
# sht_newhires = wb.sheets['NewHires']
doc = DocxTemplate('C:\\Users\\gino.depaoli\\Documents\\Python Scripts\\New Hires\\CS New User Start sheet.docx')
# context = {'full_name': "Caroline Isra"}

df = pd.read_excel('C:\\Users\\gino.depaoli\\Documents\\Python Scripts\\New Hires\\New Hires Test.xlsx')
# print(df)
# df_dict = df.to_dict()
# print()
# print(df_dict)



for index,row in df.iterrows():
        
    # print(index, row)
    # print()
    
    context = {
        'full_name': row['full_name'],
        'user_name': row['user_name'],
        'AS400': row['AS400'],
        'email': row['email'],
        'computer_name': row['computer_name'],
        'location': row['location'],
        'ext': row['ext'],
        'cs_agent_id2': row['cs_agent_id2'],
        'emp_id': row['emp_id'],
    }
    
    # print(row['full_name'])
    # print(context)
    doc.render(context)
    doc.save(f"{row['full_name']}.docx")
    

