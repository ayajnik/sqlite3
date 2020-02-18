import pyodbc as cnn
import pandas as pd

cnxn = pyodbc.connect('DRIVER={SQL Server};SERVER=SQLSRV01;DATABASE=DATABASE;UID=USER;PWD=PASSWORD')

# Copy to Clipboard for paste in Excel sheet
def copia (argumento):
    df=pd.DataFrame(argumento)
    df.to_clipboard(index=False,header=True)


tableResult = pd.read_sql("SELECT * FROM YOURTABLE", cnxn) 

# Copy to Clipboard
copia(tableResult)

# Or create a Excel file with the results
df=pd.DataFrame(tableResult)
df.to_excel("FileExample.xlsx",sheet_name='Results')