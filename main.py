from openpyxl import load_workbook
from docx import Document
from functions_EXCEL import *
from functions_WORD import * 
import os
import pathlib
from win32com.client import Dispatch
import sys
import time
import pathlib

def excluirImagensPATH(imagem):
    if os.path.exists(imagem):
        os.remove(imagem)

def pegarGraficosExcel(app, workbook_file_name, workbook):
    app.DisplayAlerts = False

    for i, sheet in enumerate(workbook.Worksheets):
        for chartObject in sheet.ChartObjects():
            chartObject.Chart.Export(rf"{str(pathlib.Path().resolve())}\chart{str(i+1)}.png")
            i +=1
    workbook.Close(SaveChanges=False, Filename=workbook_file_name)

## Inicialização do app =========
print(r"""
________            _____________                           __________  
___  __ \_________________  /__(_)______ _______________    ___  ___  \ 
__  /_/ /_  ___/  _ \  __  /__  /__  __ `__ \  _ \  ___/    __  / _ \  |
_  ____/_  /   /  __/ /_/ / _  / _  / / / / /  __/ /__      _  / , _/ / 
/_/     /_/    \___/\__,_/  /_/  /_/ /_/ /_/\___/\___/      | /_/|_| /  
                                                             \______/   
Iniciando processo de formatação...
      """)

try: 
    arquivos = os.listdir('./')
    for arquivo in arquivos:
        if arquivo.endswith(".docx"):
            ARQUIVO_WORD = arquivo
        if arquivo.endswith(".xlsm") or arquivo.endswith(".xlsx"):
            ARQUIVO_EXCEL = arquivo
    f = open(ARQUIVO_WORD, 'rb')
    g = open(ARQUIVO_EXCEL, 'rb')
    documentoWord = Document(f)
    ws = load_workbook(g)
except:
    print("ERRO: Erro ao identificar arquivos para formatação.")
    time.sleep(10)
    sys.exit(1)


planilha = ws["Listagem"]
WORD_criarTabelaListagem(planilha, documentoWord)
WORD_criarTabelasOS(documentoWord, ws)
WORD_addValoresTabelaOS(ws, documentoWord)

PASTA_RESULTADOS = "RELATÓRIOS FORMATADOS"
os.makedirs(PASTA_RESULTADOS, exist_ok=True)
caminhoWord = os.path.join(PASTA_RESULTADOS, ARQUIVO_WORD)     

try:
    app = Dispatch("Excel.Application")
    workbook_file_name = rf"{str(pathlib.Path().resolve())}\SEARA - FÁB_BRASÍLIA - RAT_202403946 (04-24).xlsm"
    workbook = app.Workbooks.Open(Filename=workbook_file_name)
    pegarGraficosExcel(app, workbook_file_name, workbook)

    for i in documentoWord.paragraphs:
        if i.text == "[grafico_status]":
            WORD_addGraficos(i, 2)
except:
    print("ERRO: Erro ao inserir imagens dos gráficos do arquivo WORD.")
    time.sleep(10)
    sys.exit(1)
documentoWord.save(caminhoWord)
        
excluirImagensPATH("chart2.png")
excluirImagensPATH("chart3.png")
excluirImagensPATH("chart4.png")             

print("\nArquivos formatados com sucesso!\n")
time.sleep(10)
