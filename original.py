import pandas as pd
import numpy as np
import math
import openpyxl
from datetime import datetime
import tkinter as tk
from interface import InterfaceUI

root = tk.Tk()
app = InterfaceUI(root)
root.mainloop()

excel_file = app.entrada_planilha
output = app.entrada_pasta + '/' + 'arquivoCRI.txt'

# Lê os parâmetros necessários para montar o txt
# sheet_name - pega os dados da planilha certa | usecols - informa quais as colunas que devem ser lidas | header=1 usa os nomes da linha 2 como títulos das colunas
df = pd.read_excel(excel_file, sheet_name='Página 1', usecols=[
                   'CRI / CRA', 'IF', 'Data de pagamento', 'Juros PU', 'Amortização PU', 'Saldo Devedor PU', 'Indexador'], header=1, engine='openpyxl')

#### Filtrando para apenas CRIs baseados em IPCA
df = df[
    (df['CRI / CRA'] == 'CRI') &
    (df['Indexador'] == 'IPCA') &
    (df['Juros PU'].notnull()) &
    (df['Juros PU'] != 0)
]

print(df)


# Abre o txt para leitura
with open(output, 'w') as file:

    # Escreve o header padrão
    file.write('CRI  0PUEVNOMESIMP            ') %ALTERAR NOMESIMP, respeitando o número adequado de caractéres
    hoje = datetime.now().date()
    hoje_formatado = hoje.strftime('%Y%m%d')
    file.write(hoje_formatado)
    file.write('00002                                ')

    # Loop para ver todas as linhas com IPCA
    for index, row in df.iterrows():

        # Escrevendo a linha que informa os dados do ativo:
        file.write('\nCRI  1PUEV')

        # Lendo os valores da planilha
        cod_if = row.get('IF')
        file.write(cod_if)
        file.write('    ')
        if (row.get('Amortização PU') == 0 or pd.isna(row.get('Amortização PU'))):
            file.write('0001\n')
        else:
            file.write('0002\n')

            # Lê os dados dos valores e formata
        juros = row.get('Juros PU')
        parte_inteira_j = int(juros)
        parte_fracionaria_j = juros - parte_inteira_j
        inteira_formatada_j = str(parte_inteira_j).zfill(10)
        fracionaria_formatada_j = f"{parte_fracionaria_j:.8f}"[2:]

        amort = row.get('Amortização PU')
        if pd.isna(amort) or amort == 0:
            # não tem amortização → coloca tudo zerado
            parte_inteira_a = 0
            parte_fracionaria_a = 0
        else:
            parte_inteira_a = int(amort)
            parte_fracionaria_a = amort - parte_inteira_a
        inteira_formatada_a = str(parte_inteira_a).zfill(10)
        fracionaria_formatada_a = f"{parte_fracionaria_a:.8f}"[2:]

        res = row.get('Saldo Devedor PU')
        parte_inteira_r = int(res)
        parte_fracionaria_r = res - parte_inteira_r
        inteira_formatada_r = str(parte_inteira_r).zfill(10)
        fracionaria_formatada_r = f"{parte_fracionaria_r:.8f}"[2:]

        # Pagamento de Juros
        file.write('CRI  2PUEV')

        # Lendo e escrevendo a data de pagamento
        data_pag = row.get('Data de pagamento')
        data_pag_formatada = data_pag.strftime('%Y%m%d')
        file.write(data_pag_formatada)
        file.write('001')

        # Escreve os valores do pagamento de juros
        file.write(inteira_formatada_j)
        file.write(fracionaria_formatada_j)
        file.write(
            '                                                                        ')

        # Pagamento da amortização
        if (row.get('Amortização PU') != 0 and not pd.isna(row.get('Amortização PU'))):
            file.write('\nCRI  2PUEV')

            # Lendo e escrevendo a data de pagamento
            data_pag = row.get('Data de pagamento')
            data_pag_formatada = data_pag.strftime('%Y%m%d')
            file.write(data_pag_formatada)
            file.write('011')

            # Escreve os valores do pagamento da Amortização
            file.write(inteira_formatada_a)
            file.write(fracionaria_formatada_a)
            file.write('                  ')
            file.write(inteira_formatada_r)
            file.write(fracionaria_formatada_r)
            file.write('000000000000000000000000000000000000')


#### Filtrando para apenas CRIs baseados em DI
df = pd.read_excel(excel_file, sheet_name='Página 1', usecols=[
                   'CRI / CRA', 'IF', 'Data de pagamento', 'Juros PU', 'Amortização PU', 'Saldo Devedor PU', 'Indexador'], header=1, engine='openpyxl')

df = df[
    (df['CRI / CRA'] == 'CRI') &
    ((df['Indexador'] == 'TAXA DI') | (df['Indexador'] == 'DI')) &
    (df['Juros PU'].notna()) &
    (df['Juros PU'] != 0)
]

print(df)

# Abre o txt para leitura
with open(output, 'a') as file:

    # Loop para ver todas as linhas com DI
    for index, row in df.iterrows():

        # Escrevendo a linha que informa os dados do ativo:
        file.write('\nCRI  1PUEV')

        # Lendo os valores da planilha
        cod_if = row.get('IF')
        file.write(cod_if)
        file.write('    ')
        if (row.get('Amortização PU') == 0 or pd.isna(row.get('Amortização PU'))):
            file.write('0001\n')
        else:
            file.write('0002\n')

            # Lê os dados dos valores e formata
        juros = row.get('Juros PU')
        parte_inteira_j = int(juros)
        parte_fracionaria_j = juros - parte_inteira_j
        inteira_formatada_j = str(parte_inteira_j).zfill(10)
        fracionaria_formatada_j = f"{parte_fracionaria_j:.8f}"[2:]

        amort = row.get('Amortização PU')
        if pd.isna(amort) or amort == 0:
            # não tem amortização → coloca tudo zerado
            parte_inteira_a = 0
            parte_fracionaria_a = 0
        else:
            parte_inteira_a = int(amort)
            parte_fracionaria_a = amort - parte_inteira_a
        inteira_formatada_a = str(parte_inteira_a).zfill(10)
        fracionaria_formatada_a = f"{parte_fracionaria_a:.8f}"[2:]

        res = row.get('Saldo Devedor PU')
        parte_inteira_r = int(res)
        parte_fracionaria_r = res - parte_inteira_r
        inteira_formatada_r = str(parte_inteira_r).zfill(10)
        fracionaria_formatada_r = f"{parte_fracionaria_r:.8f}"[2:]

        # Pagamento de Juros
        file.write('CRI  2PUEV')

        # Lendo e escrevendo a data de pagamento
        data_pag = row.get('Data de pagamento')
        data_pag_formatada = data_pag.strftime('%Y%m%d')
        file.write(data_pag_formatada)
        file.write('001')

        # Escreve os valores do pagamento de juros
        file.write(inteira_formatada_j)
        file.write(fracionaria_formatada_j)
        file.write(
            '                                                                        ')

        # Pagamento da amortização
        if (row.get('Amortização PU') != 0 and not pd.isna(row.get('Amortização PU'))):
            file.write('\nCRI  2PUEV')

            # Lendo e escrevendo a data de pagamento
            data_pag = row.get('Data de pagamento')
            data_pag_formatada = data_pag.strftime('%Y%m%d')
            file.write(data_pag_formatada)
            file.write('011')

            # Escreve os valores do pagamento da Amortização
            file.write(inteira_formatada_a)
            file.write(fracionaria_formatada_a)
            file.write('                  ')
            file.write(inteira_formatada_r)
            file.write(fracionaria_formatada_r)
            file.write('                                    ')


output = app.entrada_pasta + '/' + 'arquivoCRA.txt'
print(output)


# Lê os parâmetros necessários para montar o txt
# sheet_name - pega os dados da planilha certa | usecols - informa quais as colunas que devem ser lidas | header=1 usa os nomes da linha 2 como títulos das colunas
df = pd.read_excel(excel_file, sheet_name='Página 1', usecols=[
                   'CRI / CRA', 'IF', 'Data de pagamento', 'Juros PU', 'Amortização PU', 'Saldo Devedor PU', 'Indexador'], header=1, engine='openpyxl')

#### Filtrando para apenas CRAs baseados em IPCA
df = df[
    (df['CRI / CRA'] == 'CRA') &
    (df['Indexador'] == 'IPCA') &
    (df['Juros PU'].notna()) &
    (df['Juros PU'] != 0)
]

print(df)

# Abre o txt para leitura
with open(output, 'w') as file:

    # Escreve o header padrão
    file.write('CRA  0PUEVNOMESIMP            ') %ALTERAR NOMESIMP, respeitando o número adequado de caractéres
    hoje = datetime.now().date()
    hoje_formatado = hoje.strftime('%Y%m%d')
    file.write(hoje_formatado)
    file.write('00002                                ')

    # Loop para ver todas as linhas com IPCA
    for index, row in df.iterrows():

        # Escrevendo a linha que informa os dados do ativo:
        file.write('\nCRA  1PUEV')

        # Lendo os valores da planilha
        cod_if = row.get('IF')
        file.write(cod_if)
        file.write('   ')
        if (row.get('Amortização PU') == 0 or pd.isna(row.get('Amortização PU'))):
            file.write('0001\n')
        else:
            file.write('0002\n')

            # Lê os dados dos valores e formata
        juros = row.get('Juros PU')
        parte_inteira_j = int(juros)
        parte_fracionaria_j = juros - parte_inteira_j
        inteira_formatada_j = str(parte_inteira_j).zfill(10)
        fracionaria_formatada_j = f"{parte_fracionaria_j:.8f}"[2:]

        amort = row.get('Amortização PU')
        if pd.isna(amort) or amort == 0:
            # não tem amortização → coloca tudo zerado
            parte_inteira_a = 0
            parte_fracionaria_a = 0
        else:
            parte_inteira_a = int(amort)
            parte_fracionaria_a = amort - parte_inteira_a
        inteira_formatada_a = str(parte_inteira_a).zfill(10)
        fracionaria_formatada_a = f"{parte_fracionaria_a:.8f}"[2:]

        res = row.get('Saldo Devedor PU')
        parte_inteira_r = int(res)
        parte_fracionaria_r = res - parte_inteira_r
        inteira_formatada_r = str(parte_inteira_r).zfill(10)
        fracionaria_formatada_r = f"{parte_fracionaria_r:.8f}"[2:]

        # Pagamento de Juros
        file.write('CRA  2PUEV')

        # Lendo e escrevendo a data de pagamento
        data_pag = row.get('Data de pagamento')
        data_pag_formatada = data_pag.strftime('%Y%m%d')
        file.write(data_pag_formatada)
        file.write('001')

        # Escreve os valores do pagamento de juros
        file.write(inteira_formatada_j)
        file.write(fracionaria_formatada_j)
        file.write(
            '                                                                        ')

        # Pagamento da amortização
        if row.get('Amortização PU') != 0 and not pd.isna(row.get('Amortização PU')):
            file.write('\nCRA  2PUEV')

            # Lendo e escrevendo a data de pagamento
            data_pag = row.get('Data de pagamento')
            data_pag_formatada = data_pag.strftime('%Y%m%d')
            file.write(data_pag_formatada)
            file.write('011')

            # Escreve os valores do pagamento da Amortização
            file.write(inteira_formatada_a)
            file.write(fracionaria_formatada_a)
            file.write('                  ')
            file.write(inteira_formatada_r)
            file.write(fracionaria_formatada_r)
            file.write('000000000000000000000000000000000000')


#### Filtrando para apenas CRAs baseados em DI
df = pd.read_excel(excel_file, sheet_name='Página 1', usecols=[
                   'CRI / CRA', 'IF', 'Data de pagamento', 'Juros PU', 'Amortização PU', 'Saldo Devedor PU', 'Indexador'], header=1, engine='openpyxl')

df = df[
    (df['CRI / CRA'] == 'CRA') &
    ((df['Indexador'] == 'TAXA DI') | (df['Indexador'] == 'DI')) &
    (df['Juros PU'].notna()) &
    (df['Juros PU'] != 0)
]

print(df)

# Abre o txt para leitura
with open(output, 'a') as file:

    # Loop para ver todas as linhas com DI
    for index, row in df.iterrows():

        # Escrevendo a linha que informa os dados do ativo:
        file.write('\nCRA  1PUEV')

        # Lendo os valores da planilha
        cod_if = row.get('IF')
        file.write(cod_if)
        file.write('   ')
        if (row.get('Amortização PU') == 0 or pd.isna(row.get('Amortização PU'))):
            file.write('0001\n')
        else:
            file.write('0002\n')

            # Lê os dados dos valores e formata
        juros = row.get('Juros PU')
        parte_inteira_j = int(juros)
        parte_fracionaria_j = juros - parte_inteira_j
        inteira_formatada_j = str(parte_inteira_j).zfill(10)
        fracionaria_formatada_j = f"{parte_fracionaria_j:.8f}"[2:]

        amort = row.get('Amortização PU')
        if pd.isna(amort) or amort == 0:
            # não tem amortização → coloca tudo zerado
            parte_inteira_a = 0
            parte_fracionaria_a = 0
        else:
            parte_inteira_a = int(amort)
            parte_fracionaria_a = amort - parte_inteira_a
        inteira_formatada_a = str(parte_inteira_a).zfill(10)
        fracionaria_formatada_a = f"{parte_fracionaria_a:.8f}"[2:]

        res = row.get('Saldo Devedor PU')
        parte_inteira_r = int(res)
        parte_fracionaria_r = res - parte_inteira_r
        inteira_formatada_r = str(parte_inteira_r).zfill(10)
        fracionaria_formatada_r = f"{parte_fracionaria_r:.8f}"[2:]

        # Pagamento de Juros
        file.write('CRA  2PUEV')

        # Lendo e escrevendo a data de pagamento
        data_pag = row.get('Data de pagamento')
        data_pag_formatada = data_pag.strftime('%Y%m%d')
        file.write(data_pag_formatada)
        file.write('001')

        # Escreve os valores do pagamento de juros
        file.write(inteira_formatada_j)
        file.write(fracionaria_formatada_j)
        file.write(
            '                                                                        ')

        # Pagamento da amortização
        if row.get('Amortização PU') != 0 and not pd.isna(row.get('Amortização PU')):
            file.write('\nCRA  2PUEV')

            # Lendo e escrevendo a data de pagamento
            data_pag = row.get('Data de pagamento')
            data_pag_formatada = data_pag.strftime('%Y%m%d')
            file.write(data_pag_formatada)
            file.write('011')

            # Escreve os valores do pagamento da Amortização
            file.write(inteira_formatada_a)
            file.write(fracionaria_formatada_a)
            file.write('                  ')
            file.write(inteira_formatada_r)
            file.write(fracionaria_formatada_r)

            file.write('                                    ')
