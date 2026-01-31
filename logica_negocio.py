import json
# import random

# lê os arquivos JSON
with open("dados_escala.json", "r") as f:
    escala_json = json.load(f)

with open("atividades_pendentes.json", "r") as f:
    atividades_json = json.load(f)

# def distribuir_por_turno(lista_colaboradores, lista_atividades): #VERSÃO 1 (QUE RODA ATUALMENTE)
#     # Ordena por grupo para garantir a prioridade
#     atividades_ordenadas = sorted(lista_atividades, key=lambda x: int(x['grupo']))
#     disponiveis = lista_colaboradores.copy()
#     # disponiveis = random.sample(lista_colaboradores, len(lista_colaboradores)) # aleatório para teste posterior
    
#     # inicia o dicionário de alocação com listas vazias
#     alocacao = {atv['nome']: [] for atv in lista_atividades}

#     for atividade in atividades_ordenadas:
#         nome_atv = atividade['nome']
#         qtd_necessaria = int(atividade['pessoas'])
#         links = [l.strip() for l in atividade.get('link', "").split(",") if l.strip()]

#         # Verifica qnts vagas faltam (considerando o que já veio de link)
#         vagas_preenchidas = len(alocacao[nome_atv])
#         vagas_faltantes = qtd_necessaria - vagas_preenchidas

#         if vagas_faltantes > 0:
#             # Pega os próximos colaboradores disponíveis
#             novos = disponiveis[:vagas_faltantes]
            
#             for colaborador in novos:
#                 alocacao[nome_atv].append(colaborador)
                
#                 # RESERVA: Se essa atividade tem links, já coloca o colaborador neles também
#                 for nome_linkado in links:
#                     if nome_linkado in alocacao:
#                         alocacao[nome_linkado].append(colaborador)
            
#             # Remove quem foi usado da lista geral de disponíveis
#             for n in novos:
#                 disponiveis.remove(n)

#         # LIMPEZA 1: Remove duplicados que podem ter vindo dos links cruzados
#         alocacao[nome_atv] = list(set(alocacao[nome_atv]))

#         # Verifica regra de "Pode Vazio"
#         if not alocacao[nome_atv] and atividade.get('pode_vazio') == "N":
#             alocacao[nome_atv] = ["SEM COLABORADOR"]

#     # distribui quem sobrou, se tiver
#     if disponiveis:
#         for atividade in atividades_ordenadas:
#             if atividade.get('pode_exceder') == "S":
#                 nome_atv_excede = atividade['nome']
                
#                 # Adiciona o restante dos disponíveis
#                 alocacao[nome_atv_excede].extend(disponiveis)
                
#                 # LIMPEZA 2: Remove duplicado
#                 alocacao[nome_atv_excede] = list(set(alocacao[nome_atv_excede]))
                
#                 break

#     return alocacao


# def distribuir_por_turno(lista_colaboradores, lista_atividades): #VERSÃO 2 (COM RECURSIVIDADE, MSM PESSOA FICOU EM MTS ATIVIDADES)
#     # Ordena por grupo para garantir a prioridade
#     atividades_ordenadas = sorted(lista_atividades, key=lambda x: int(x['grupo']))
#     disponiveis = lista_colaboradores.copy()
#     # disponiveis = random.sample(lista_colaboradores, len(lista_colaboradores)) # aleatório para teste posterior
    
#     # inicia o dicionário de alocação com listas vazias
#     alocacao = {atv['nome']: [] for atv in lista_atividades}

#     for atividade in atividades_ordenadas:
#         nome_atv = atividade['nome']
#         qtd_necessaria = int(atividade['pessoas'])
#         links = [l.strip() for l in atividade.get('link', "").split(",") if l.strip()]

#         # Verifica qnts vagas faltam (considerando o que já veio de link)
#         vagas_preenchidas = len(alocacao[nome_atv])
#         vagas_faltantes = qtd_necessaria - vagas_preenchidas

#         if vagas_faltantes > 0:
#             # Pega os próximos colaboradores disponíveis
#             novos = disponiveis[:vagas_faltantes]
            
#             for colaborador in novos:
#                 alocacao[nome_atv].append(colaborador)
                
#                 # RECURSIVIDADE SIMPLIFICADA: 
#                 # Começamos com os links da atividade atual
#                 links_para_processar = [l.strip() for l in atividade.get('link', "").split(",") if l.strip()]
                
#                 while links_para_processar:
#                     proximo_nome = links_para_processar.pop(0) # Pega o primeiro link da fila
                    
#                     if proximo_nome in alocacao and colaborador not in alocacao[proximo_nome]:
#                         alocacao[proximo_nome].append(colaborador)
                        
#                         # AQUI ESTÁ O SEGREDO: 
#                         # Busca a definição da atividade linkada para ver se ELA tem links
#                         info_linkada = next((a for a in lista_atividades if a['nome'] == proximo_nome), None)
#                         if info_linkada:
#                             novos_links = [l.strip() for l in info_linkada.get('link', "").split(",") if l.strip()]
#                             links_para_processar.extend(novos_links) # Adiciona os links da atividade B para serem verificados
            
#             # Remove quem foi usado da lista geral de disponíveis
#             for n in novos:
#                 disponiveis.remove(n)

#         # LIMPEZA 1: Remove duplicados que podem ter vindo dos links cruzados
#         alocacao[nome_atv] = list(set(alocacao[nome_atv]))

#         # Verifica regra de "Pode Vazio"
#         if not alocacao[nome_atv] and atividade.get('pode_vazio') == "N":
#             alocacao[nome_atv] = ["SEM COLABORADOR"]

#     # distribui quem sobrou, se tiver
#     if disponiveis:
#         for atividade in atividades_ordenadas:
#             if atividade.get('pode_exceder') == "S":
#                 nome_atv_excede = atividade['nome']
                
#                 # Adiciona o restante dos disponíveis
#                 alocacao[nome_atv_excede].extend(disponiveis)
                
#                 # LIMPEZA 2: Remove duplicado
#                 alocacao[nome_atv_excede] = list(set(alocacao[nome_atv_excede]))
                
#                 break

#     return alocacao


def distribuir_por_turno(lista_colaboradores, lista_atividades): #VERSÃO 3 (PARECE Q DEU CERTO, TESTAR)
    # Ordena por grupo para garantir a prioridade
    atividades_ordenadas = sorted(lista_atividades, key=lambda x: int(x['grupo']))
    disponiveis = lista_colaboradores.copy()
    # disponiveis = random.sample(lista_colaboradores, len(lista_colaboradores)) # aleatório para teste posterior
    
    # inicia o dicionário de alocação com listas vazias
    alocacao = {atv['nome']: [] for atv in lista_atividades}

    for atividade in atividades_ordenadas:
        nome_atv = atividade['nome']
        qtd_necessaria = int(atividade['pessoas'])
        links = [l.strip() for l in atividade.get('link', "").split(",") if l.strip()]

        # Verifica qnts vagas faltam (considerando o que já veio de link)
        vagas_preenchidas = len(alocacao[nome_atv])
        vagas_faltantes = qtd_necessaria - vagas_preenchidas

        if vagas_faltantes > 0:
            novos = disponiveis[:vagas_faltantes]
            
            for colaborador in novos:
                alocacao[nome_atv].append(colaborador)
                
                # RESERVA: ver vagas na atividade com link antes de distribuir
                for nome_linkado in links:
                    if nome_linkado in alocacao:
                        atv_dest = next((a for a in lista_atividades if a['nome'] == nome_linkado), None)
                        if atv_dest:
                            # Só entra se tiver vaga ou se puder exceder
                            vagas_dest = int(atv_dest['pessoas']) - len(alocacao[nome_linkado])
                            pode_exc = atv_dest.get('pode_exceder') == "S"
                            
                            if (vagas_dest > 0 or pode_exc) and colaborador not in alocacao[nome_linkado]:
                                alocacao[nome_linkado].append(colaborador)
            
            for n in novos:
                disponiveis.remove(n)
        
        else:
            # CASO A ATIVIDADE JÁ ESTEJA CHEIA
            for nome_linkado in links:
                if nome_linkado in alocacao:
                    atv_dest = next((a for a in lista_atividades if a['nome'] == nome_linkado), None)
                    if atv_dest:
                        for colaborador in alocacao[nome_atv]:
                            vagas_dest = int(atv_dest['pessoas']) - len(alocacao[nome_linkado])
                            pode_exc = atv_dest.get('pode_exceder') == "S"
                            
                            if (vagas_dest > 0 or pode_exc) and colaborador not in alocacao[nome_linkado]:
                                alocacao[nome_linkado].append(colaborador)
                            elif not pode_exc:
                                break # Para de tentar colocar gente se encheu e não pode exceder

        # LIMPEZA 1: Remove duplicados que podem ter vindo dos links cruzados
        alocacao[nome_atv] = list(set(alocacao[nome_atv]))

        # Verifica regra de "Pode Vazio"
        if not alocacao[nome_atv] and atividade.get('pode_vazio') == "N":
            alocacao[nome_atv] = ["SEM COLABORADOR"]

    # distribui quem sobrou, se tiver
    if disponiveis:
        for atividade in atividades_ordenadas:
            if atividade.get('pode_exceder') == "S":
                nome_atv_excede = atividade['nome']
                
                # Adiciona o restante dos disponíveis
                alocacao[nome_atv_excede].extend(disponiveis)
                
                # LIMPEZA 2: Remove duplicado
                alocacao[nome_atv_excede] = list(set(alocacao[nome_atv_excede]))
                
                break

    return alocacao



# --- EXECUÇÃO PRINCIPAL ---

resultado_final = {}

for data_str, turnos in escala_json.items():    
    # Processa cada turno daquela data
    resultado_final[data_str] = {}
    for nome_do_turno, lista_colaboradores in turnos.items():
        resultado_final[data_str][nome_do_turno] = distribuir_por_turno(lista_colaboradores, atividades_json)

# Exemplo do resultado
for data, turnos in resultado_final.items():
    print(f"\n=== ESCALA DO DIA: {data} ===")

    for atividade in atividades_json:
        nome_atividade = atividade['nome']
        print('-----------------------------------------------------')
        print(f"Atividade: {nome_atividade.upper()}")

        for nome_turno, distribuicao_turno in turnos.items():
            equipe = distribuicao_turno.get(nome_atividade, [])

            print(f"  {nome_turno}: {' / '.join(equipe)}")
            

