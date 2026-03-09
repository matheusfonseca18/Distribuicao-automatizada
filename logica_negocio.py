import json
import os
import random
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.worksheet.table import Table, TableStyleInfo
from CTkMessagebox import CTkMessagebox
from tkinter import filedialog

# lê os arquivos JSON
def ler_arquivos_json(nome_arquivo):
    if os.path.exists(nome_arquivo):
        with open(nome_arquivo, "r") as f:
            return json.load(f)
    return {}

def salvar_arquivo_json(nome_arquivo, dados):
    if os.path.exists(nome_arquivo):
        with open(nome_arquivo, "w") as f:
            json.dump(dados, f, indent=4)

def distribuir_por_turno(lista_colaboradores, lista_atividades, historico, data_atual):
    atividades_ordenadas = sorted(lista_atividades, key=lambda x: int(x['grupo']))
    disponiveis = list(lista_colaboradores) # Cópia da lista
    alocacao = {atv['nome']: [] for atv in lista_atividades}
    origem_do_colaborador = {}
    mapa_links = {atv['nome']: [l.strip() for l in atv.get('link', "").split(",") if l.strip()] for atv in lista_atividades}

    for atividade in atividades_ordenadas:
        nome_atv = atividade['nome']
        qtd_necessaria = int(atividade['pessoas'])
        links_raiz = mapa_links.get(nome_atv, [])

        def calcular_peso_escolha(colab):
            # Histórico específico da atividade (Fator Principal)
            dados_h = historico.get(colab, {}).get(nome_atv, {"qtd": 0, "ultima_data": ""})
            
            # Total de atividades que a pessoa já fez
            total_geral = sum(info["qtd"] for info in historico.get(colab, {}).values())
            
            # PENALIDADE PESADA: Se ele já fez ESSA atividade muitas vezes,o score sobe drasticamente (multiplicador de 100)
            score = (dados_h["qtd"] * 100) + total_geral
            
            # Bloqueio de repetição diária (se trabalhou nela por último, vai pro fim da fila)
            if dados_h["ultima_data"] == data_atual:
                score += 500
            
            return score

        # REORDENA a cada nova atividade para garantir que o mais apto no momento seja eleito
        disponiveis.sort(key=calcular_peso_escolha)

        vagas_faltantes = qtd_necessaria - len(alocacao[nome_atv])
        
        if vagas_faltantes > 0 and disponiveis:
            # Pega os N primeiros que têm o MENOR score para ESTA atividade
            novos = disponiveis[:vagas_faltantes]
            
            for colaborador in novos:
                alocacao[nome_atv].append(colaborador)
                origem_do_colaborador[colaborador] = nome_atv
                
                # ATUALIZAÇÃO ÚNICA: Evita duplicar histórico
                if colaborador not in historico: historico[colaborador] = {}
                if nome_atv not in historico[colaborador]:
                    historico[colaborador][nome_atv] = {"qtd": 0, "ultima_data": ""}
                
                historico[colaborador][nome_atv]["qtd"] += 1
                historico[colaborador][nome_atv]["ultima_data"] = data_atual

                # Distribui para Links Diretos (Sem contar +1 no histórico de 'qtd' para não inflar)
                for nome_linkado in links_raiz:
                    if nome_linkado in alocacao:
                        atv_dest = next((a for a in lista_atividades if a['nome'] == nome_linkado), None)
                        if atv_dest:
                            v_dest = int(atv_dest['pessoas']) - len(alocacao[nome_linkado])
                            pode_exc = atv_dest.get('pode_exceder') == "S"
                            
                            if (v_dest > 0 or pode_exc) and colaborador not in alocacao[nome_linkado]:
                                alocacao[nome_linkado].append(colaborador)
            
            for n in novos:
                disponiveis.remove(n)

        else:
            # CASO A ATIVIDADE JÁ ESTEJA CHEIA (Verificação B -> C com base na Origem)
            for nome_linkado in links_raiz:
                if nome_linkado in alocacao:
                    atv_dest = next((a for a in lista_atividades if a['nome'] == nome_linkado), None)
                    if atv_dest:
                        links_de_b = mapa_links.get(nome_linkado, [])
                        
                        for colaborador in alocacao[nome_atv]:
                            atv_origem = origem_do_colaborador.get(colaborador)
                            links_autorizados_origem = mapa_links.get(atv_origem, [])

                            v_dest = int(atv_dest['pessoas']) - len(alocacao[nome_linkado])
                            pode_exc = atv_dest.get('pode_exceder') == "S"

                            # Tenta alocar em B
                            if (v_dest > 0 or pode_exc) and colaborador not in alocacao[nome_linkado]:
                                alocacao[nome_linkado].append(colaborador)
                                if colaborador not in historico: historico[colaborador] = {}
                                if nome_linkado not in historico[colaborador]:
                                    historico[colaborador][nome_linkado] = {"qtd": 0, "ultima_data": ""}
                                historico[colaborador][nome_linkado]["qtd"] += 1
                                historico[colaborador][nome_linkado]["ultima_data"] = data_atual

                                # Tenta alocar no link de B (C) apenas se a origem autorizar
                                for link_de_b in links_de_b:
                                    if link_de_b in links_autorizados_origem and link_de_b in alocacao:
                                        atv_final = next((a for a in lista_atividades if a['nome'] == link_de_b), None)
                                        if atv_final:
                                            v_f = int(atv_final['pessoas']) - len(alocacao[link_de_b])
                                            if (v_f > 0 or atv_final.get('pode_exceder') == "S") and colaborador not in alocacao[link_de_b]:
                                                alocacao[link_de_b].append(colaborador)
                                                if link_de_b not in historico[colaborador]:
                                                    historico[colaborador][link_de_b] = {"qtd": 0, "ultima_data": ""}
                                                historico[colaborador][link_de_b]["qtd"] += 1
                                                historico[colaborador][link_de_b]["ultima_data"] = data_atual
                            elif not pode_exc:
                                break

        # Limpeza FinaL
        alocacao[nome_atv] = list(set(alocacao[nome_atv]))

    # Distribui quem sobrou
    if disponiveis:
        for atividade in atividades_ordenadas:
            if atividade.get('pode_exceder') == "S":
                nome_atv_excede = atividade['nome']
                alocacao[nome_atv_excede].extend(disponiveis)
                alocacao[nome_atv_excede] = list(set(alocacao[nome_atv_excede]))
                break

    return alocacao, historico


def gerar_distribuicao():
    # Carrega os dados
    escala_json = ler_arquivos_json("dados_escala.json")
    atividades = ler_arquivos_json("atividades_pendentes.json")
    # Carrega o histórico existente ou cria um vazio
    historico_geral = ler_arquivos_json("historico.json") 

    resultado_final = {}

    for data_str, turnos in escala_json.items():    
        resultado_final[data_str] = {}
        
        for nome_do_turno, lista_colaboradores in turnos.items():
            # recebe a alocação e o histórico atualizado
            alocacao_turno, historico_geral = distribuir_por_turno(
                lista_colaboradores, 
                atividades, 
                historico_geral,
                data_str
            )
            resultado_final[data_str][nome_do_turno] = alocacao_turno

    # GERAÇÃO DO EXCEL
    caminho_salvar = filedialog.asksaveasfilename(
        defaultextension='.xlsx', 
        filetypes=[("Arquivo Excel", "*.xlsx")], 
        title='Salvar Distribuição', 
        initialfile=f'Distribuição.xlsx'
    )

    if caminho_salvar:
        try:
            # nome_arquivo_excel = os.path.basename(caminho_salvar)

            wb = Workbook()
            
            # ABA DISTRIBUIÇÃO (DADOS DO PANDAS)
            ws_dist = wb.active
            ws_dist.title = "Distribuição"
            
            # Criar o DataFrame df_final
            linhas_base = []
            lista_turnos = sorted(list(set(t for turnos in escala_json.values() for t in turnos.keys())))
            for atividade in atividades:
                nome_atv_fixo = atividade['nome'].strip().upper()
                for turno in lista_turnos:
                    linhas_base.append({"ATIVIDADE": nome_atv_fixo, "TURNO": turno})
            
            df_final = pd.DataFrame(linhas_base)
            for data, turnos_da_data in resultado_final.items():
                data_curta = data.split(" ")[0]
                coluna_colaboradores = []
                for index, row in df_final.iterrows():
                    atv_procurada = row["ATIVIDADE"]
                    turno_procurado = row["TURNO"]
                    alocacao_turno = turnos_da_data.get(turno_procurado, {})
                    equipe = next((lista for n, lista in alocacao_turno.items() if n.strip().upper() == atv_procurada), [])
                    coluna_colaboradores.append(" / ".join(equipe))
                df_final[data_curta] = coluna_colaboradores

            # Converter DataFrame do Pandas para linhas do Openpyxl
            for r in dataframe_to_rows(df_final, index=False, header=True):
                ws_dist.append(r)

            # ABA HISTÓRICO
            ws_historico = wb.create_sheet(title="Histórico")
            
            # Identifica os colaboradores únicos na escala
            todos_colaboradores = set()
            for data_str, turnos in escala_json.items():
                for nome_turno, lista in turnos.items():
                    for nome_colab in lista:
                        todos_colaboradores.add((nome_colab.strip().upper(), nome_turno.upper()))
            
            # Ordenar por turno e depois nome
            lista_colaboradores_ordenada = sorted(list(todos_colaboradores), key=lambda x: (x[1], x[0]))

            # Identifica as atividades para criar as colunas
            nomes_atividades = [atv['nome'].strip().upper() for atv in atividades]
            colunas_historico = ['COLABORADOR', 'TURNO'] + nomes_atividades
            ws_historico.append(colunas_historico)

            # Preenche os dados e as fórmulas
            for i, (nome_colab, turno_colab) in enumerate(lista_colaboradores_ordenada, start=2):
                # Escreve Nome e Turno
                ws_historico.cell(row=i, column=1).value = nome_colab
                ws_historico.cell(row=i, column=2).value = turno_colab
                
                # Para cada atividade (coluna), insere a fórmula de contagem
                for j, nome_atv in enumerate(nomes_atividades, start=3):                    
                    formula = (
                        f'=SUMPRODUCT((Distribuição!$A$2:$A$1000="{nome_atv}")*'
                        f'(Distribuição!$B$2:$B$1000="{turno_colab}")*'
                        f'(--ISNUMBER(SEARCH("{nome_colab}", Distribuição!$C$2:$AG$1000))))'
                    )
                    ws_historico.cell(row=i, column=j).value = formula

            # Formata como Tabela
            ultima_col_letra = chr(64 + len(colunas_historico)) # Converte índice para letra
            if len(colunas_historico) > 26: # se tiver muitas atividades
                ultima_col_letra = "Z" 
                
            ref_tabela = f'A1:{ultima_col_letra}{len(lista_colaboradores_ordenada) + 1}'
            tab = Table(displayName='Tabela_Historico', ref=ref_tabela)
            style = TableStyleInfo(name="TableStyleMedium9", showRowStripes=True)
            tab.tableStyleInfo = style
            ws_historico.add_table(tab)
            
            # Salvar o arquivo final
            wb.save(caminho_salvar)

            CTkMessagebox(title="Sucesso", message="Distribuição baixado com sucesso!", icon="check")
        except Exception as e:
            CTkMessagebox(title="Erro", message=f"Erro ao salvar: {e}", icon="cancel")
