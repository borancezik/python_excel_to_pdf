from multiprocessing import context
from docxtpl import DocxTemplate
import pandas as pd
import time
import os

from win32com import client
word_app = client.Dispatch("Word.Application")

sayı1 = 4.376
data_frame = pd.read_excel("data2.xlsx")

for r_index, row in data_frame.iterrows():
    cust_name = row['FİRMA_ADI']
    
    tpl = DocxTemplate("template.docx")
    df_to_doct = data_frame.to_dict()
    x = data_frame.to_dict(orient='records')
    context = x
    #print(context[r_index])
    tpl.render(context[r_index])
    tpl.save('Doc\\'+str(r_index)+".docx")

    time.sleep(1)

    ROOT_DIR = os.path.dirname(os.path.abspath(__file__))
    print(ROOT_DIR)

    doc = word_app.Documents.Open(ROOT_DIR+'\\Doc\\'+str(r_index)+'.docx')
    print(str(r_index+1)+". belge PDf'e dönüştürülüyor")
    doc.SaveAs(ROOT_DIR+'\\PDF\\'+str(r_index)+'.pdf',FileFormat=17)

    time.sleep(0.5)

print("PDF oluşturma işlemi bitmiştir.!!!!")


