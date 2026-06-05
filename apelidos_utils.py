import json
import os


ARQUIVO_APELIDOS = "apelidos.json"


def normalizar_nome(nome):
    return str(nome).strip()


def chave_nome(nome):
    return normalizar_nome(nome).upper()


def normalizar_apelidos(dados):
    if not isinstance(dados, dict):
        return {}

    apelidos = {}
    for nome, valor in dados.items():
        nome_limpo = normalizar_nome(nome)
        if not nome_limpo:
            continue

        if isinstance(valor, dict):
            apelido = valor.get("apelido", "")
        else:
            apelido = valor

        apelido_limpo = normalizar_nome(apelido)
        if apelido_limpo:
            apelidos[nome_limpo] = apelido_limpo

    return apelidos


def carregar_apelidos(nome_arquivo=ARQUIVO_APELIDOS):
    if not os.path.exists(nome_arquivo):
        return {}

    try:
        with open(nome_arquivo, "r", encoding="utf-8") as arquivo:
            conteudo = arquivo.read().strip().upper()
    except OSError:
        return {}

    if not conteudo:
        return {}

    try:
        dados = json.loads(conteudo)
    except json.JSONDecodeError:
        try:
            dados, _ = json.JSONDecoder().raw_decode(conteudo)
        except json.JSONDecodeError:
            return {}

    return normalizar_apelidos(dados)


def salvar_apelidos(apelidos, nome_arquivo=ARQUIVO_APELIDOS):
    dados_limpos = normalizar_apelidos(apelidos)

    with open(nome_arquivo, "w", encoding="utf-8") as arquivo:
        json.dump(dados_limpos, arquivo, ensure_ascii=False, indent=4)


def listar_colaboradores_unicos(dados_escala):
    if not isinstance(dados_escala, dict):
        return []

    colaboradores = []
    vistos = set()

    for turnos in dados_escala.values():
        if not isinstance(turnos, dict):
            continue

        for lista_colaboradores in turnos.values():
            if not isinstance(lista_colaboradores, list):
                continue

            for nome in lista_colaboradores:
                nome_limpo = normalizar_nome(nome)
                chave = chave_nome(nome_limpo)
                if nome_limpo and chave not in vistos:
                    colaboradores.append(nome_limpo)
                    vistos.add(chave)

    return sorted(colaboradores, key=chave_nome)


def montar_mapa_apelidos(apelidos):
    return {
        chave_nome(nome): normalizar_nome(apelido)
        for nome, apelido in normalizar_apelidos(apelidos).items()
        if normalizar_nome(apelido)
    }


def nome_exibicao(nome, mapa_apelidos):
    nome_limpo = normalizar_nome(nome)
    return mapa_apelidos.get(chave_nome(nome_limpo), nome_limpo)
