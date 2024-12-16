def EXCEL_pegarValorTabelaListagem(planilha, coluna, linha):
    textoCelula = str(planilha[f"{coluna}{linha+1}"].value) if str(planilha[f"{coluna}{linha+1}"].value) != "None" else "" 
    return textoCelula

def EXCEL_retornarOSdaListagem(planilha):
    os = []
    for item in planilha["G"]:
        os.append(item.value if item.value != None else "")
    os = [item for item in os if not item==""]
    os = [item for item in os if not item==" "]
    os = [item for item in os if not item=="OS"]
    infos = EXCEL_retornarInfosListagem(planilha)
    if len(infos) != 0:    
        for i in range(0,len(os)):
            infos[i].append(os[i])
    return infos

def EXCEL_retornarInfosListagem(planilha):
    infos = []
    coluna = "F" if EXCEL_segundoAcompanhamento(planilha) else "E"
    linha = 1
    for item in planilha[coluna]:
        if item.value == "Aceitável" or item.value == "Crítico" or item.value == "Alerta" or item.value == "Crítica":
            infos.append([item.value, planilha[f"B{linha}"].value, planilha[f"C{linha}"].value, planilha[f"D{linha}"].value])
        linha+=1
    return infos

def EXCEL_segundoAcompanhamento(planilha):
    if planilha["F1"].value == "2°Acompanhamento":
        return False
    return True

def EXCEL_organizarNumerosOS(planilha):
    os = EXCEL_retornarOSdaListagem(planilha)
    osOrganizado = []
    if len(os) != 0:
        for i in os:
            if type(i[4]) is str and "e" in i[4]:
                nmSeparado = i[4].split(" e ")
                for j in nmSeparado:
                    osOrganizado.append([int(j) , i])
            if type(i[4]) is str and "a" in i[4]:
                nmSeparado = i[4].split(" a ")
                for j in range(int(nmSeparado[0]), int(nmSeparado[1])+1):
                    osOrganizado.append([int(j), i])
            if type(i[4]) is int:
                nm = i.pop()
                i.append("remv")
                osOrganizado.append([int(nm),i])
    def obter_numero(item):
        return item[0]
    lista_ordenada = sorted(osOrganizado, key=obter_numero)
    return lista_ordenada
        