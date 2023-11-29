'''Esse codigo é um complemento do main1.py, ele serve para imprimir os dados
da planilha sem prescisar abrir a mesma.'''

import openpyxl

#Carrega o Arquivo e selecionar pagina
book = openpyxl.load_workbook('Planilha de Alunos.xlsx')
alunos_page = book['Alunos']
#imprimindo os dados de cada linha, passando da linha 2 a 5, pois esta delimitado no codigo.
for rows in alunos_page.iter_rows(min_rows=2, max_rows=5):
   print(rows[0].value,rows[1].value,rows[2].value)
