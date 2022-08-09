from ast import Break
from msilib.schema import Directory
import pandas as pd
import ctypes
import os, glob
from tkinter import filedialog

working_dir = filedialog.askdirectory()
path = "lines.csv"
new_file = open(path,'w')

for data_file in glob.glob(os.path.join(working_dir, '*.txt')):
    filename,ext = os.path.splitext(os.path.basename(data_file))
    for line in reversed(list(open(data_file))):
        if "OFFSET_ADC" in line:
            line = line.split()
            new_file.write(filename+","+line[4]+","+line[7]+'\n')
            break

new_file.close()
new_file = open(path,'r')

df = pd.read_csv(new_file, sep=',', engine='python', header=None)

## Configuração do arquivo .xlsx
writer = pd.ExcelWriter("lines.xlsx", engine='xlsxwriter')

df.to_excel(writer, sheet_name='Sheet1', startrow=1, header=False, index=False)

workbook = writer.book
worksheet = writer.sheets['Sheet1']

(max_row, max_col) = df.shape
column_settings = []

worksheet.add_table(0, 0, max_row, max_col - 1, {'columns': column_settings})
worksheet.set_column(0, max_col - 1, 12)
writer.save()

MessageBox = ctypes.windll.user32.MessageBoxW
MessageBox(None, 'Processo finalizado.\nArquivos .csv e .xlsx salvos em:\n\n'+os.getcwd(), 'Sucesso', 0)