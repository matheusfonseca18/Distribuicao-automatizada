import json
import os
import random
import pandas as pd

# lê os arquivos JSON
def ler_arquivos_json(nome_arquivo):
    if os.path.exists(nome_arquivo):
        with open(nome_arquivo, "r") as f:
            return json.load(f)
    return {}

def salvar_arquivo_json(nome_arquivo, dados):
    # if os.path.exists(nome_arquivo):
    with open(nome_arquivo, "w") as f:
        json.dump(dados, f, indent=4)

def distribuir_por_turno(lista_colaboradores, lista_atividades, historico, data_atual):
    # ordena por grupo
    atividades_ordenadas = sorted(lista_atividades, key=lambda x: int(x['grupo']))
    
    disponiveis = random.sample(lista_colaboradores, len(lista_colaboradores))
    
    alocacao = {atv['nome']: [] for atv in lista_atividades}
    
    # Rastreia onde o colaborador foi escalado primeiro para validar links
    origem_do_colaborador = {}

    # Mapa de links para consulta rápida
    mapa_links = {atv['nome']: [l.strip() for l in atv.get('link', "").split(",") if l.strip()] for atv in lista_atividades}

    for atividade in atividades_ordenadas:
        nome_atv = atividade['nome']
        qtd_necessaria = int(atividade['pessoas'])
        links_raiz = mapa_links.get(nome_atv, [])

        # equilíbrio
        def calcular_peso_escolha(colab):
            # Busca dados no histórico: {"qtd": 0, "ultima_data": ""}
            dados_h = historico.get(colab, {}).get(nome_atv, {"qtd": 0, "ultima_data": ""})
            
            score = dados_h["qtd"]
            
            # Penalidade por repetição: se trabalhou ontem, score sobe muito
            # Isso joga o colaborador para o fim da fila, mas o mantém como opção se não houver outros
            if dados_h["ultima_data"] != "" and dados_h["ultima_data"] < data_atual:
                score += 100 
            
            return score

        disponiveis.sort(key=calcular_peso_escolha)

        # alocação
        vagas_faltantes = qtd_necessaria - len(alocacao[nome_atv])
        
        if vagas_faltantes > 0:
            novos = disponiveis[:vagas_faltantes]
            for colaborador in novos:
                alocacao[nome_atv].append(colaborador)
                origem_do_colaborador[colaborador] = nome_atv
                
                # Atualiza Histórico da Atividade Principal
                if colaborador not in historico: historico[colaborador] = {}
                if nome_atv not in historico[colaborador]:
                    historico[colaborador][nome_atv] = {"qtd": 0, "ultima_data": ""}
                
                historico[colaborador][nome_atv]["qtd"] += 1
                historico[colaborador][nome_atv]["ultima_data"] = data_atual

                # Distribui para Links Diretos (B)
                for nome_linkado in links_raiz:
                    if nome_linkado in alocacao:
                        atv_dest = next((a for a in lista_atividades if a['nome'] == nome_linkado), None)
                        if atv_dest:
                            v_dest = int(atv_dest['pessoas']) - len(alocacao[nome_linkado])
                            pode_exc = atv_dest.get('pode_exceder') == "S"
                            
                            if (v_dest > 0 or pode_exc) and colaborador not in alocacao[nome_linkado]:
                                alocacao[nome_linkado].append(colaborador)
                                # Atualiza Histórico do Link
                                if nome_linkado not in historico[colaborador]:
                                    historico[colaborador][nome_linkado] = {"qtd": 0, "ultima_data": ""}
                                historico[colaborador][nome_linkado]["qtd"] += 1
                                historico[colaborador][nome_linkado]["ultima_data"] = data_atual
            
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

        # Limpezas Finais
        alocacao[nome_atv] = list(set(alocacao[nome_atv]))
        if not alocacao[nome_atv] and atividade.get('pode_vazio') == "N":
            alocacao[nome_atv] = ["SEM COLABORADOR"]

    # Distribui quem sobrou
    if disponiveis:
        for atividade in atividades_ordenadas:
            if atividade.get('pode_exceder') == "S":
                nome_atv_excede = atividade['nome']
                alocacao[nome_atv_excede].extend(disponiveis)
                alocacao[nome_atv_excede] = list(set(alocacao[nome_atv_excede]))
                break

    return alocacao, historico


# --- EXECUÇÃO PRINCIPAL ---
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

    # Salva o histórico para a próxima vez que rodar o código (funciona pro print, mas n cria o arquivo)
    salvar_arquivo_json("historico.json", historico_geral)

    # --- SAÍDA NO TERMINAL ---
    # for data, turnos in resultado_final.items():
    #     print(f"\n=== ESCALA DO DIA: {data} ===")

    #     for atividade in atividades:
    #         nome_atividade = atividade['nome']
    #         print('-----------------------------------------------------')
    #         print(f"Atividade: {nome_atividade.upper()}")

    #         for nome_turno, distribuicao_turno in turnos.items():
    #             equipe = distribuicao_turno.get(nome_atividade, [])
    #             print(f"  {nome_turno}: {' / '.join(equipe)}")

    # # --- RELATÓRIO DE HISTÓRICO ---
    # print("\n\n=== RELATÓRIO DE ACÚMULO (HISTÓRICO) ===")
    # for colab, tarefas in sorted(historico_geral.items()):
    #     tarefas_ordenadas = sorted(tarefas.items())
    #     resumo = ", ".join([f"{t}: {q}" for t, q in tarefas_ordenadas])

    # --- GERAÇÃO DO EXCEL ---
    nome_arquivo_excel = "Escala_Final.xlsx"

    with pd.ExcelWriter(nome_arquivo_excel) as writer:
        # ESQUELETO
        linhas_base = []
        lista_turnos = sorted(list(set(t for turnos in escala_json.values() for t in turnos.keys())))
        
        for atividade in atividades:
            nome_atv_fixo = atividade['nome'].strip().upper()
            for turno in lista_turnos:
                linhas_base.append({
                    "ATIVIDADE": nome_atv_fixo,
                    "TURNO": turno
                })

        df_final = pd.DataFrame(linhas_base)

        # PREENCHIMENTO DAS DATAS
        for data, turnos_da_data in resultado_final.items():
            data_curta = data.split(" ")[0]
            coluna_colaboradores = []

            for index, row in df_final.iterrows():
                atv_procurada = row["ATIVIDADE"]
                turno_procurado = row["TURNO"]
                
                # Acessa os dados alocados para aquele turno e dia
                alocacao_turno = turnos_da_data.get(turno_procurado, {})
                
                # Busca ignorando maiúsculas/minúsculas
                equipe = []
                for nome_atv_json, lista_nomes in alocacao_turno.items():
                    if nome_atv_json.strip().upper() == atv_procurada:
                        equipe = lista_nomes
                        break
                coluna_colaboradores.append(" / ".join(equipe))
            
            df_final[data_curta] = coluna_colaboradores

        df_final.to_excel(writer, sheet_name="Distribuição", index=False)
        
        # Aba Histórico
        dados_historico = [{"COLABORADOR": c, "ACÚMULO": ", ".join([f"{t}: {info['qtd']}" for t, info in sorted(tar.items())])} 
                        for c, tar in sorted(historico_geral.items())]
        pd.DataFrame(dados_historico).to_excel(writer, sheet_name="Histórico", index=False)

    print(f"Arquivo gerado: {nome_arquivo_excel}")
