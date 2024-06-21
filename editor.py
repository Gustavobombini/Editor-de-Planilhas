import configparser
import pandas as pd
from datetime import datetime
import glob
import os

config = configparser.ConfigParser()
config.read('config.ini', encoding='utf-8')
now = datetime.now()
file_path = config['GERAL']['entrada']
file_path_drop = config['GERAL']['saida']
encoding = 'latin1'  
delimiter = ';'      
rename = config['RENOMEAR']['rename'].split(",")
columnsDell = config['COLUNAS']['columnsDell'].split(',')
rowDell = config['LINHAS']['rowDell'].split(',')
timestamp = now.strftime("%d-%m-%Y_%H-%M-%S")
taxa1 = config['TAXA']['colum1']
taxa2 = config['TAXA']['colum2']
tabelasTomoney = config['CONVERT']['columns'].split(',')
pattern = os.path.join(file_path, '*')
files = glob.glob(pattern)

def clean_and_convert(column):
    column = column.str.replace(',', '.', regex=False)  # Substituir vírgulas por pontos
    column = column.str.replace('[^\d.]', '', regex=True) 
    return pd.to_numeric(column, errors='coerce').fillna(0)  

def format_money(value):
    return f'R$ {value:,.2f}'.replace(',', 'v').replace('.', ',').replace('v', '.')


for file in files:
    print(file)

    pd.options.display.max_rows = 9999
    df = pd.read_csv(f'{file}', encoding=encoding, delimiter=delimiter)

    print(df.columns)


    if taxa1 in df.columns and taxa2 in df.columns:
        df[taxa1] = clean_and_convert(df[taxa1].astype(str))
        df[taxa2] = clean_and_convert(df[taxa2].astype(str))

        df['TAXA'] = (df[taxa2] / df[taxa1])*100
        df['TAXA'] = df['TAXA'].round(2)

    else:
        print(f"Colunas {taxa1} ou {taxa2} não encontradas para cálculo da nova coluna.")


    for value in rename:
        data = value.split(":")
        if data[0] in df.columns:
            df = df.rename({data[0] : data[1]},  axis='columns')
        else:
            print(f"Coluna '{value}' não encontrada e não foi Renomeada.")

    for coluna in tabelasTomoney:
        if coluna in df.columns:
            df[coluna] = clean_and_convert(df[coluna].astype(str))
            df[coluna] = pd.to_numeric(df[coluna], errors='coerce').fillna(0)
            df[coluna] = df[coluna].apply(lambda x: format_money(x))
        else:
            print(f"Coluna '{coluna}' não encontrada e não foi convertida.")

    for value in columnsDell:
            if value in df.columns:
                df.drop(columns=value, inplace=True)
            else:
                print(f"Coluna '{value}' não encontrada e não foi removida.")

    for value in rowDell:
        df = df[~df.isin([value]).any(axis=1)]

    df.to_excel(f'{file_path_drop}{timestamp}.xls', index=False, engine='openpyxl')
