# -*- coding: utf-8 -*-
"""
Created on Fri Jun  3 10:17:25 2022

@author: x290372
"""

from pdf2image import convert_from_path
from PIL import Image
from pytesseract import pytesseract
from pathlib import Path
from os.path import getmtime
import pandas as pd
import os
import re

Image.MAX_IMAGE_PIXELS = 933120000


def pathLocal():
    absFilePath = os.path.abspath(__file__)
    filePath = os.path.split(absFilePath)
    path = filePath[0]
    return path


def extractTextRegex(pattern, text):
    result = ''
    pattern = re.compile(pattern)
    search = pattern.finditer(text)
    for a in search:
        result = a.group(0)
    return result


def menu():

    ### CONFIGURAÇÕES ###
    path_1 = r'\\fspib01\PIBFS\Santander\Asegs'
    path_2 = r'\\fspib01\pibfs2\santander2\Asegs'

    file_1 = 'Procura_no_diretorio1.xlsx'
    file_2 = 'Procura_no_diretorio2.xlsx'
    file_3 = 'Procura_no_diretorio3.xlsx'
    file_4 = 'Procura_no_diretorio4.xlsx'

    ### INICIO MENU ###
    print('---------- ESCOLHA O DIRETORIO ----------')
    print('1 - ' + path_1)
    print('2 - ' + path_2)
    proposalPath = int(input('Opção: '))

    if proposalPath == 1:
        proposalPath = path_1
    elif proposalPath == 2:
        proposalPath = path_2
    else:
        print('Opção Invalida.')

    print('---------- ESCOLHA O ARQUIVO ----------')
    print('1 - ' + file_1)
    print('2 - ' + file_2)
    print('3 - ' + file_3)
    print('4 - ' + file_4)
    file_name = int(input('Opção: '))

    if file_name == 1:
        file_name = pathLocal() + '\\input\\' + file_1
    elif file_name == 2:
        file_name = pathLocal() + '\\input\\' + file_2
    elif file_name == 3:
        file_name = pathLocal() + '\\input\\' + file_3
    elif file_name == 4:
        file_name = pathLocal() + '\\input\\' + file_4
    else:
        print('Opção Invalida.')

    pagina = input('Informe a pagina inicial: ')

    ### MOSTRA AS OPÇÕES ###

    print('\nCONFIGURAÇÃO REALIZADA')
    print('Local: ' + proposalPath)
    print('Arquivo: ' + file_name)
    print('Pagina: ' + pagina)

    return proposalPath, file_name, pagina


### DEFINE VARIAVEIS ###
tesseractPath = pathLocal() + r'\dependencies\Tesseract-OCR\tesseract.exe'
popplerPath = pathLocal() + r'\dependencies\poppler-22.01.0\Library\bin'
tempPath = pathLocal() + r'\temp'
outputPath = pathLocal() + r'\output'
proposalPath, file_name, pagina = menu()

### INICIO SCRIPT ###
for a in range(int(pagina), 200):

    print('\nIniciando leitura da pagina ' + str(a))

    # Faz a leitura do arquivo de entrada
    df = pd.read_excel(file_name)

    # percorre por todas as linhas coletando o nome da pasta
    for index, row in df.iterrows():

        ### SALVA A PROPOSTA ATUAL ###
        ProposalName = str(int(row['Proposta']))
        print("--- Buscando Proposta: " + str(ProposalName) + " ---")

        ### VALIDA SE A LINHA JÁ FOI LIDA ###
        if row['Tag'] == 'x':
            print("--- Proposta já lida ---\n")
            continue

        ### VALIDA SE A LINHA JÁ FOI LIDA (PAGINA ATUAL) ###
        if row['pagina'] >= a:
            print("--- Pagina referente a essa proposta já lida ---\n")
            continue

        ### CRIA BACKUP A CADA 1000 REGISTROS ###
        if index % 1000 == 0:
            df.to_excel(file_name.replace(
                'Procura', str(index) + ' - Procura'), index=False)

        # coleta os ultimos 3 arquivos mais recentes da pasta
        directory = Path(proposalPath + '\\' + ProposalName)
        files = directory.glob('*.pdf')
        sorted_files = sorted(files, key=getmtime, reverse=True)
        count = 0
        pula = 0

        try:
            if len(sorted_files) == 0:
                df.loc[df["Proposta"] == int(ProposalName), "Tag"] = "x"
                df.loc[df["Proposta"] == int(
                    ProposalName), "Obs"] = "Não foi localizado o diretorio."
                df.to_excel(file_name, index=False)
                print("--- Não foi localizado o diretorio ---\n")
                continue

            for pdf in sorted_files:
                if count < 3:
                    if pula > 0:
                        break

                    PDFPath = str(pdf)
                    PDFName = PDFPath.split('\\')
                    PDFName = PDFName[7]
                    numberLoop = a

                    while True:
                        try:
                            image = convert_from_path(
                                PDFPath, 500, poppler_path=popplerPath, first_page=numberLoop, last_page=numberLoop)

                            # Gera imagem png
                            imgName = PDFName.replace('pdf', 'png')
                            imgPath = tempPath + '\\' + imgName
                            image[0].save(imgPath, "PNG")

                            # Coleta o texto do png
                            resultText = ''
                            img = Image.open(tempPath + '\\' + imgName)
                            pytesseract.tesseract_cmd = tesseractPath
                            text = pytesseract.image_to_string(img)
                            resultText = resultText + text

                            # procura a pagina correta que contem todas as informações
                            iden_santander = extractTextRegex(
                                r'BANCO SANTANDER', resultText)
                            iden_registro_geral = extractTextRegex(
                                r'REGISTRO GERAL', resultText)

                            if iden_santander and iden_registro_geral:

                                obs = f'Achou! - Proposta: {ProposalName} - Arquivo: {PDFName} - Pagina: {numberLoop}'
                                print(f'{obs}')

                                df.loc[df["Proposta"] == int(
                                    ProposalName), "Tag"] = "x"
                                df.loc[df["Proposta"] == int(
                                    ProposalName), "Obs"] = obs
                                df.loc[df["Proposta"] == int(
                                    ProposalName), "pagina"] = numberLoop
                                df.loc[df["Proposta"] == int(
                                    ProposalName), "txt"] = resultText
                                df.to_excel(file_name, index=False)

                                pula = 1
                                break

                            else:
                                obs = f'Não achei nessa pagina - Proposta: {ProposalName} - Arquivo: {count} - Pagina: {numberLoop}'
                                print(f'{obs}')

                                df.loc[df["Proposta"] == int(
                                    ProposalName), "Obs"] = obs
                                df.loc[df["Proposta"] == int(
                                    ProposalName), "pagina"] = numberLoop
                                df.to_excel(file_name, index=False)

                                break

                        except IndexError:

                            if count == 2:
                                obs = f'Procurei em todo o arquivo e não achei! - Proposta: {ProposalName} - Arquivo: {PDFName} - Pagina: {numberLoop}'
                                print(f'{obs}')

                                df.loc[df["Proposta"] == int(
                                    ProposalName), "Obs"] = obs
                                df.loc[df["Proposta"] == int(
                                    ProposalName), "pagina"] = numberLoop
                                df.loc[df["Proposta"] == int(
                                    ProposalName), "Tag"] = "x"
                                df.to_excel(file_name, index=False)

                            obs = f'Procurei em todo o arquivo e não achei! - Proposta: {ProposalName} - Arquivo: {PDFName} - Pagina: {numberLoop}'
                            print(f'{obs}')

                            df.loc[df["Proposta"] == int(
                                ProposalName), "Obs"] = obs
                            df.loc[df["Proposta"] == int(
                                ProposalName), "pagina"] = numberLoop
                            df.to_excel(file_name, index=False)

                            break

                    count += 1

                else:
                    print("--- Próx Propota ---")
                    break

        except Image.DecompressionBombError:
            print("Não foi possivel tratar a imagem dessa proposta.")
            df.loc[df["Proposta"] == int(
                ProposalName), "Obs"] = 'Erro de Memoria - '
            df.loc[df["Proposta"] == int(ProposalName), "pagina"] = numberLoop
            df.to_excel(file_name, index=False)
            continue

        except:
            print("Erro não conhecido.")
            continue
