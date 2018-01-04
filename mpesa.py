from __future__ import print_function
import pandas as pd
import subprocess,sys,os,re
from tabula import read_pdf
from decimal import Decimal
from pandas import ExcelWriter
 

if sys.version_info >= (3,):
    import tkinter as tk
    from tkinter import filedialog as fd
else:
    import Tkinter as tk
    import tkFileDialog as fd

def pdftotext(pdf, page=None):
        """Retrieve all text from a PDF file.
   
        Arguments:
            pdf Path of the file to read.
            page: Number of the page to read. If None, read all the pages.
   
        Returns:
            A list of lines of text.
        """
        if page is None:
            args = ['pdf2text.py', pdf]
        else:
            args = ['pdf2txt.py', '-p', str(page), pdf]
        try:
            txt = subprocess.check_output(args,shell=True,universal_newlines=True)
            lines = txt.splitlines()
        except subprocess.CalledProcessError:
            lines = []
        return lines

def process_mpesa(pdf,num):    
    file_actual = pdf.rsplit('/',1)
    print ('File:',file_actual[1])
    print ('Number:',num)
#       
    df = read_pdf(pdf,pages='all',encoding='ISO-8859-1')
       
    df.drop(df.columns[[1,3]], axis=1,inplace=True)
    df.drop(df.index[:1], inplace=True)
    
    columns = df.columns
    df.rename(columns={columns[0]:'date_time',\
                       columns[1]:'transaction_details',\
                       columns[2]:'money_in',\
                       columns[3]:'money_out',\
                       columns[4]:'balance'},inplace=True);
    
    df['transaction_details'] = df.transaction_details.str.replace('Completed' , '')
    
    for index, row in df.iterrows():
        if row['date_time'] is not None:
            if len(str(row['date_time'])) > 16:
                df['date_time'] = df['date_time'].map(lambda x: str(x)[:16])

    df = df.drop(df[df.date_time.str.contains('Operator')].index)
    df = df.drop(df[df.date_time.str.contains('Date &')].index)
    
    df['date_time'] = pd.to_datetime(df['date_time'], format="%Y-%m-%d %H:%M")

    
    df['money_in'] = df.money_in.str.replace(',' , '')
    df['money_out'] = df.money_out.str.replace(',' , '')
    df['balance'] = df.balance.str.replace(',' , '')
    
    for index, row in df.iterrows():
        if row['money_out'] is not None:
            if re.search(" ", str(row['money_out'])):
                out,bal = row['money_out'].split(' ')
                df.set_value(index,'money_out',out)
                df.set_value(index,'balance',bal)
    
    df['money_in'] = pd.to_numeric(df['money_in'])
    df['money_out'] = pd.to_numeric(df['money_out'])
    df['balance'] = pd.to_numeric(df['balance'])
    
    for index, row in df.iterrows():
        if row['money_in'] is not None:
            if row['money_out'] > 0:
                df.set_value(index,'money_out',Decimal('nan'))
                df.set_value(index,'balance',row['money_out'])

    df = df[pd.notnull(df['transaction_details'])]
    
    for index,rows in df.iterrows():
        if pd.isnull(rows['date_time']):
            previous  = (index - 1)
            if pd.notnull(rows['transaction_details']):
                trans = df.get_value(previous,'transaction_details') + ' ' + df.get_value(index,'transaction_details')
                df.set_value(index,'transaction_details',str(None))
                next_ = (index + 1)
                try:
                    next_date = df.get_value(next_,'date_time')
                    if pd.isnull(next_date):
                        trans = trans + ' ' + df.get_value(next_,'transaction_details')
                        df.set_value(next_,'transaction_details',str(None))
                except KeyError:
                    pass
                df.set_value(previous,'transaction_details',trans)
    df = df.drop(df[df.transaction_details == 'None'].index).sort_values(by='date_time',ascending=True)
    df = df.drop(df[df.transaction_details == 'None None'].index)
          
    df_in = df[pd.notnull(df['money_in'])]
    df_in_c = df_in.groupby('transaction_details')['money_in'].agg(['sum','count']).reset_index().sort_values(by=['sum','count'],ascending=False)
    df_in_tills = df_in_c[df_in_c['transaction_details'].str.contains("Till")].sort_values(by='count',ascending=False)
    df_in_others = df_in_c[df_in_c['transaction_details'].str.contains("Funds")].sort_values(by='count',ascending=False)
        
    df_out = df[pd.notnull(df['money_out'])]
    df_out_c = df_out.groupby('transaction_details')['money_out'].agg(['sum','count']).reset_index().sort_values(by=['sum','count'],ascending=True)
    df_out_tills = df_out_c[df_out_c['transaction_details'].str.contains("Agent")].sort_values(by='count',ascending=False)
    df_out_others = df_out_c[df_out_c['transaction_details'].str.contains("Transfer of funds")].sort_values(by='count',ascending=False)
    df_out_bus = df_out_c[df_out_c['transaction_details'].str.contains("Payment")].sort_values(by='count',ascending=False)
        
    writer = ExcelWriter(file_actual[0]+'/'+ 'mpesa_' + number + '.xlsx')
    df.to_excel(writer,'mpesa')
    df_in_c.to_excel(writer,'mpesa_in')
    df_in_tills.to_excel(writer,'mpesa_in_tills')
    df_in_others.to_excel(writer,'mpesa_in_others')
    df_out_c.to_excel(writer,'mpesa_out')
    df_out_tills.to_excel(writer,'mpesa_out_tills')
    df_out_others.to_excel(writer,'mpesa_out_others')
    df_out_bus.to_excel(writer,'mpesa_payments')
    writer.save()
        
    os.rename(pdf,file_actual[0]+'/'+number+'.pdf') ## Python 3 we can use replace
    print ('Finished processing: ',file_actual[1])
 
    
if __name__ == "__main__":
    tk.Tk().withdraw()
    filez = fd.askopenfilenames(title='Choose a file')
    lst = list(filez)
    for statement in lst:
        number = pdftotext(statement, 1)[6]
        process_mpesa(statement,number)