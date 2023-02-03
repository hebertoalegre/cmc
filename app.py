# basics
import sqlite3
from docxtpl import DocxTemplate, InlineImage
import numpy as np
import pandas as pd
import datetime
from datetime import datetime

import locale
locale.setlocale(locale.LC_TIME, "es_GT")

pd.set_option("styler.format.thousands", ",")
# connecting to sqlite
conn = sqlite3.connect(r'C:\Users\DELL\Desktop\cmc\cmc.db')
curr = conn.cursor()

# selecting dataframe
query ='''
SELECT * FROM cmc WHERE fecha >= '2021-01-01'
'''
data = pd.read_sql(query, con=conn)
data.valor = data.valor.astype(float)
data.fecha = pd.to_datetime([datetime.strptime(i, '%Y-%m-%d') for i in data.fecha])

data = data.dropna(axis=0)
var = ['IMAE', 'PIB trimestral en constantes','IPC general','Tasa de política monetaria','Exportaciones totales','Importaciones totales',' Tipo de cambio de venta fin de mes',
       'ITCER global','ITCER con USA','Remesas, ingreso','RIN del Banco Central','Ingresos totales GC','Gastos totales GC','Deuda total / PIB']

# creating template
doc=DocxTemplate(r'C:\Users\DELL\Desktop\cmc\PLANTILLA2.docx')
# adding images to template
image_equal = InlineImage(doc,'equal.jpg')
image_up = InlineImage(doc,'down.jpg')
image_down = InlineImage(doc,'up.jpg')


#extracting the data to jsonfy#
todos = {}
for i in range(len(var)):
    x = data[data['variable'].str.contains(var[i])]
    if i ==1:
        conditional  = x.groupby(['pais', 'variable']).nth([-1]) -x.groupby(['pais', 'variable']).nth([-5])
        conditional = conditional.reset_index()['valor'].tolist()
        conditional =  [image_up if val>0 else image_equal if val==0 else image_down  for val in conditional]
        x.loc[:,'change'] = round(100*x.groupby('pais')['valor'].pct_change(periods = 4),2)
        x = x.groupby(['pais','variable']).tail(1).stack()
        x = x.reset_index().iloc[:,1:]
        fecha = [d.date().strftime('a %B %Y') for d in x[x['level_1']=='fecha'][0].tolist()]
        items = [item for item in zip(fecha,x[x['level_1']=='change'][0].tolist(), conditional)]
        
    elif any(loc == i for loc in [3,6,13]):
        conditional  = x.groupby(['pais', 'variable']).nth([-1]) -x.groupby(['pais', 'variable']).nth([-13])
        conditional = conditional.reset_index()['valor'].tolist()
        conditional =  [image_up if val>0 else image_equal if val==0 else image_down  for val in conditional]
        x = x.groupby(['pais','variable']).tail(1).stack()
        x = x.reset_index().iloc[:,1:]
        fecha = [d.date().strftime('a %B %Y') for d in x[x['level_1']=='fecha'][0].tolist()]
        items = [item for item in zip(fecha,x[x['level_1']=='valor'][0].tolist(), conditional)]

    elif any(loc==i for loc in [4,5,9,11,12]):
        conditional  = x.groupby(['pais', 'variable']).nth([-1]) -x.groupby(['pais', 'variable']).nth([-5])
        conditional = conditional.reset_index()['valor'].tolist()
        conditional =  [image_up if val>0 else image_equal if val==0 else image_down  for val in conditional]
        x.loc[:,'change'] = round(100*x.groupby([x.pais, x.fecha.dt.year])['valor'].cumsum().pct_change(periods = 12),2)
        x.loc[:,'change_abs'] = x.groupby([x.pais, x.fecha.dt.year])['valor'].cumsum()
        x = x.groupby(['pais','variable']).tail(1).stack()
        x = x.reset_index().iloc[:,1:]
        fecha = [d.date().strftime('a %B %Y') for d in x[x['level_1']=='fecha'][0].tolist()]
        items = [item for item in zip(fecha, x[x['level_1']=='change'][0].tolist(), conditional, x[x['level_1']=='change_abs'][0].tolist())]

    else:
        conditional  = x.groupby(['pais', 'variable']).nth([-1]) -x.groupby(['pais', 'variable']).nth([-5])
        conditional = conditional.reset_index()['valor'].tolist()
        conditional =  [image_up if val>0 else image_equal if val==0 else image_down  for val in conditional]
        x.loc[:,'change'] = round(100*x.groupby('pais')['valor'].pct_change(periods = 12),2)
        x = x.groupby(['pais','variable']).tail(1).stack()
        x = x.reset_index().iloc[:,1:]
        fecha = [d.date().strftime('a %B %Y') for d in x[x['level_1']=='fecha'][0].tolist()]
        items = [item for item in zip(fecha,x[x['level_1']=='change'][0].tolist(), conditional)]

        
    titulo = x[x['level_1']=='variable'][0].unique().tolist()[0]
    todos[titulo] ={'valores': items}
    

output = {'contenido': todos}
output['columnas'] = [i.strip()for i in data.pais.unique().tolist()]
doc.render(output)
doc.save("editable_cmc.docx")

    


 



