
import os
from openpyxl import Workbook
import pdfplumber
import re


diretorio = 'pdfs'
arquivos = os.listdir(diretorio)
arquivos_qtd = len(arquivos)

if arquivos_qtd == 0:
    raise Exception('Nenhum arquivo PDF encontrado no diretório.')

planilha = Workbook()
aba = planilha.active
aba.title = 'Sicovimed'
aba['A1'] = 'CNPJ'
aba['B1'] = 'Vencimento'
aba['C1'] = 'Valor'
aba['D1'] = 'Número do Documento'
aba['E1'] = 'Status'

ultima_linha = 1
while aba['A' + str(ultima_linha)].value is not None:
    ultima_linha += 1
    
for arquivo in arquivos:
    with pdfplumber.open(diretorio + '/' + arquivo) as pdf:
        primeira_pagina = pdf.pages[0]
        pdf_texto = primeira_pagina.extract_text()
        
    pdf_cnpj = r'\b\d{2}\.?\d{3}\.?\d{3}/?\d{4}-?\d{2}'
    pdf_vencimento = r'VENCIMENTO[\s\S]*?(\d{2}/\d{2}/\d{4})'
    pdf_valor = r'VALOR\s+DO\s+DOCUMENTO[:\s\n]*([\d\.]+,\d{2})'
    pdf_numero_documento = r'NRO\.DOC:\s*(\w+)'

    match_cnpj = re.search(pdf_cnpj, pdf_texto)
    match_vencimento = re.search(pdf_vencimento, pdf_texto)
    match_valor = re.search(pdf_valor, pdf_texto)
    match_numero_documento = re.search(pdf_numero_documento, pdf_texto)

    if match_cnpj:
        cnpj = match_cnpj.group()
        aba['A{}'.format(ultima_linha)] = cnpj
    else:
        aba['A{}'.format(ultima_linha)] = ' CNPJ não encontrado'

    if match_vencimento:
        vencimento = match_vencimento.group(1)
        aba['B{}'.format(ultima_linha)] = vencimento
    else:
        vencimento = ' Vencimento não encontrado'
        aba['B{}'.format(ultima_linha)] = 'vencimento não encontrado'
        
    if match_valor:
        valor = match_valor.group(1)
        aba['C{}'.format(ultima_linha)] = valor
    else:
        valor = ' valor não encontrado'
        aba['C{}'.format(ultima_linha)] = 'valor não encontrado'

    if match_numero_documento:
        numero_documento = match_numero_documento.group(1)
        aba['D{}'.format(ultima_linha)] = numero_documento
    else:
        numero_documento = ' Número do Documento não encontrado'
        aba['D{}'.format(ultima_linha)] = 'número do documento não encontrado'

    aba['E{}'.format(ultima_linha)] = 'Completa'
    
    ultima_linha += 1
    
planilha.save ('tabela.xlsx')
os.startfile('tabela.xlsx')