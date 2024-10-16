from data_table import valoresOS

def EXCEL_retornarPlanilhasOS(workbook):
    listaOS = []
    for i in workbook.sheetnames:
        if i != "Dados" and i !="Listagem" and i != "Gráficos" and i !="":
            listaOS.append(i)
    return listaOS

def EXCEL_addValoresTabelasOS(planilhaOStext, ws):
    objvalores = {}
    planilhaOS = ws[planilhaOStext]
    for i in valoresOS:
        objvalores[i] = planilhaOS[valoresOS[i]].value if planilhaOS[valoresOS[i]].value != None else ""
    return objvalores

def EXCEL_pegarValorTabelaListagem(planilha, coluna, linha):
    textoCelula = str(planilha[f"{coluna}{linha+1}"].value) if planilha[f"{coluna}{linha+1}"].value != "None" else "" 
    return textoCelula
