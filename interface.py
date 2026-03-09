import tkinter
from tkinter import StringVar
import customtkinter as ctk
from CTkMessagebox import CTkMessagebox
import os
from tkinter.filedialog import askopenfilename
from PIL import Image
from logica_interface import salvar_dados, carregar_dados, carregar_dados_arquivo, validar_numero, processar_excel_escala, baixar_escala_modelo
from logica_negocio import gerar_distribuicao

# Listas globais que armazena os dicionários de dados das atividades
NOME_ARQUIVO_DADOS_ATIVIDADES = "atividades_pendentes.json"
dados_atividades = []  
NOME_ARQUIVO_DADOS_ESCALA = "dados_escala.json"
dados_escala = {} 
NOME_ARQUIVO_DADOS_ARQUIVO = "dados_arquivo.json"
dados_arquivo = []

def importar_escala():
    caminho = askopenfilename(title="Seleciona um arquivo Excel", filetypes=[("Arquivo Excel", "*.xlsx *.xls")])
    if not caminho:
        return

    resultado = processar_excel_escala(caminho)
    
    if resultado:
        global dados_escala, dados_arquivo
        dados_escala = resultado
        
        nome_arquivo_importado = os.path.basename(caminho)
        
        # Garante que dados_arquivo seja uma lista antes do append
        if dados_arquivo is None:
            dados_arquivo = []
            
        dados_arquivo.append(nome_arquivo_importado)
        
        salvar_dados(NOME_ARQUIVO_DADOS_ESCALA, dados_escala)
        salvar_dados(NOME_ARQUIVO_DADOS_ARQUIVO, dados_arquivo)
        
        nome_arquivo.configure(text=str(nome_arquivo_importado))
        CTkMessagebox(title="Sucesso", message=f"Arquivo importado com sucesso: {nome_arquivo_importado}")


# FUNÇÕES de Interface

def alternar_edicao(frame, botao):
    # Verifica se está no modo de edição ou salvando
    if botao.cget('text') == 'Editar':
        # MODO EDIÇÃO
        botao.configure(text='Salvar', fg_color="green")
        
        # Nome (Label > Entry)
        texto_atual = frame.lbl_nome.cget('text')
        frame.lbl_nome.destroy()
        frame.entry_nome = ctk.CTkEntry(frame, width=100, justify="center")
        frame.entry_nome.insert(0, texto_atual)
        frame.entry_nome.grid(row=0, column=0, padx=10)
        
        # Pessoas (Label > Entry)
        texto_atual = frame.lbl_pessoas.cget('text')
        frame.lbl_pessoas.destroy()
        frame.entry_pessoas = ctk.CTkEntry(frame, width=100, justify="center")
        frame.entry_pessoas.insert(0, texto_atual)
        frame.entry_pessoas.grid(row=0, column=1, padx=10)

        # Grupo (Label > Entry)
        texto_atual = frame.lbl_grupo.cget('text')
        frame.lbl_grupo.destroy()
        frame.entry_grupo = ctk.CTkEntry(frame, width=100, justify="center")
        frame.entry_grupo.insert(0, texto_atual)
        frame.entry_grupo.grid(row=0, column=2, padx=10)

        # Vazio (Label > Container > Checkbox Centralizado)
        val_atual = frame.lbl_vazio.cget('text')
        frame.lbl_vazio.destroy()
        
        # Cria um container transparente para segurar a posição
        frame.container_vazio = ctk.CTkFrame(frame, width=100, height=28, fg_color="transparent")
        frame.container_vazio.grid(row=0, column=3, padx=10)
        
        frame.var_vazio_edit = tkinter.IntVar(value=1 if val_atual == "S" else 0)
        frame.chk_vazio = ctk.CTkCheckBox(frame.container_vazio, text="", variable=frame.var_vazio_edit, width=0)
        # Place centraliza dentro do container de 100px
        frame.chk_vazio.place(relx=0.5, rely=0.5, anchor="center") 

        # Excede (Label > Container > Checkbox Centralizado)
        val_atual = frame.lbl_excede.cget('text')
        frame.lbl_excede.destroy()
        
        # Cria um container transparente
        frame.container_excede = ctk.CTkFrame(frame, width=100, height=28, fg_color="transparent")
        frame.container_excede.grid(row=0, column=4, padx=10)
        
        frame.var_excede_edit = tkinter.IntVar(value=1 if val_atual == "S" else 0)
        frame.chk_excede = ctk.CTkCheckBox(frame.container_excede, text="", variable=frame.var_excede_edit, width=0)
        frame.chk_excede.place(relx=0.5, rely=0.5, anchor="center")

    else:
        # MODO SALVAR
        # Atualizar o dicionário de dados
        nome_antigo = frame.dados['nome']  # Guarda o nome antigo para localizar no JSON

        frame.dados['nome'] = frame.entry_nome.get()
        frame.dados['pessoas'] = frame.entry_pessoas.get()
        frame.dados['grupo'] = frame.entry_grupo.get()
        frame.dados['pode_vazio'] = "S" if frame.var_vazio_edit.get() == 1 else "N"
        frame.dados['pode_exceder'] = "S" if frame.var_excede_edit.get() == 1 else "N"

        nome_novo = frame.entry_nome.get()
        
        for atividade in dados_atividades:
            if "link" in atividade and atividade["link"]:
                # Remove a atividade excluída dos links das outras atividades
                links = [link.strip() for link in atividade["link"].split(", ")]
                if nome_antigo in links:
                    links = [nome_novo if link == nome_antigo else link for link in links]
                    atividade["link"] = ", ".join(links) if links else ""

        # Salva no arquivo JSON
        salvar_dados(NOME_ARQUIVO_DADOS_ATIVIDADES, dados_atividades)

        # volta a interface pro padrão (Widgets Editáveis > Labels)
        
        # Nome
        novo_texto = frame.dados['nome']
        frame.entry_nome.destroy()
        frame.lbl_nome = ctk.CTkLabel(frame, text=novo_texto, width=100, wraplength=100)
        frame.lbl_nome.grid(row=0, column=0, padx=10)

        # Pessoas
        novo_texto = frame.dados['pessoas']
        frame.entry_pessoas.destroy()
        frame.lbl_pessoas = ctk.CTkLabel(frame, text=novo_texto, width=100, wraplength=100)
        frame.lbl_pessoas.grid(row=0, column=1, padx=10)

        # Grupo
        novo_texto = frame.dados['grupo']
        frame.entry_grupo.destroy()
        frame.lbl_grupo = ctk.CTkLabel(frame, text=novo_texto, width=100, wraplength=100)
        frame.lbl_grupo.grid(row=0, column=2, padx=10)

        # Vazio - Destrói o container (o checkbox morre junto)
        novo_texto = frame.dados['pode_vazio']
        frame.container_vazio.destroy() 
        frame.lbl_vazio = ctk.CTkLabel(frame, text=novo_texto, width=100)
        frame.lbl_vazio.grid(row=0, column=3, padx=10)

        # Excede - Destrói o container
        novo_texto = frame.dados['pode_exceder']
        frame.container_excede.destroy()
        frame.lbl_excede = ctk.CTkLabel(frame, text=novo_texto, width=100)
        frame.lbl_excede.grid(row=0, column=4, padx=10)

        # Restaura botão
        botao.configure(text='Editar', fg_color=["#3B8ED0", "#1F6AA5"])

def criar_widget_atividade(atividade_dados):    
    novo_frame = ctk.CTkFrame(frame_atividades, fg_color='transparent')
    novo_frame.pack(pady=5, padx=10, fill="x")
    
    novo_frame.dados = atividade_dados 

    # Armazena referências dos Labels no novo_frame para acesso na edição
    novo_frame.lbl_nome = ctk.CTkLabel(novo_frame, text=atividade_dados['nome'], width=100, wraplength=100)
    novo_frame.lbl_nome.grid(row=0, column=0, padx=10)

    novo_frame.lbl_pessoas = ctk.CTkLabel(novo_frame, text=f"{atividade_dados['pessoas']}", width=100, wraplength=100)
    novo_frame.lbl_pessoas.grid(row=0, column=1, padx=10)

    novo_frame.lbl_grupo = ctk.CTkLabel(novo_frame, text=f"{atividade_dados['grupo']}", width=100, wraplength=100)
    novo_frame.lbl_grupo.grid(row=0, column=2, padx=10)

    novo_frame.lbl_vazio = ctk.CTkLabel(novo_frame, text=f"{atividade_dados['pode_vazio']}", width=100)
    novo_frame.lbl_vazio.grid(row=0, column=3, padx=10)

    novo_frame.lbl_excede = ctk.CTkLabel(novo_frame, text=f"{atividade_dados['pode_exceder']}", width=100)
    novo_frame.lbl_excede.grid(row=0, column=4, padx=10)

    my_image = ctk.CTkImage(light_image=Image.open("link.png"), dark_image=Image.open("link.png"), size=(18, 18))
    links = ctk.CTkButton(novo_frame, text="", image=my_image, command=lambda d=atividade_dados: abrir_janela_link(d), fg_color="transparent", width=50)
    links.grid(row=0, column=5, padx=10)

    botao_editar = ctk.CTkButton(novo_frame, text='Editar', width=60)
    # Atribui a função de alternar edição passando o frame e o próprio botão
    botao_editar.configure(command=lambda f=novo_frame, b=botao_editar: alternar_edicao(f, b))
    botao_editar.grid(row=0, column=6, padx=10)

    botao_apagar = ctk.CTkButton(novo_frame, text='Apagar', command=lambda f=novo_frame: apagar_atividade(f), width=60)
    botao_apagar.grid(row=0, column=7, padx=10)

def apagar_atividade(frame_para_apagar): # Remove o widget da tela e o dicionário da lista de dados
    if hasattr(frame_para_apagar, 'dados') and frame_para_apagar.dados in dados_atividades:
        nome_atividade_excluida = frame_para_apagar.dados['nome']
        for atividade in dados_atividades:
            if "link" in atividade and nome_atividade_excluida in atividade["link"]:
                # Remove a atividade excluída dos links das outras atividades
                links = atividade["link"].split(", ")
                links = [link.strip() for link in links if link.strip() != nome_atividade_excluida]
                atividade["link"] = ", ".join(links) if links else ""

        dados_atividades.remove(frame_para_apagar.dados)
        salvar_dados(NOME_ARQUIVO_DADOS_ATIVIDADES, dados_atividades) # Salva e atualizado
    
    frame_para_apagar.destroy()

def criar_nova_atividade():    
    nome_atividade = nome_var.get().strip()
    pessoas_atividade = pessoas_var.get().strip()
    grupo_atividade = grupo_var.get().strip()
    vazio_atividade = vazio_var.get()
    excede_atividade = excede_var.get()
    pode_vazio = "S" if vazio_atividade == 1 else "N"
    pode_exceder = "S" if excede_atividade == 1 else "N"

    if not nome_atividade or not pessoas_atividade or not grupo_atividade:
        CTkMessagebox(
            title="Alerta de Validação", 
            message="Por favor, preencha todos os campos obrigatórios.",
            icon="warning")
        return

    # estrutura o dicionário de dados
    nova_atividade_dados = {
        "nome": nome_atividade,
        "pessoas": pessoas_atividade,
        "grupo": grupo_atividade,
        "pode_vazio": pode_vazio,
        "pode_exceder": pode_exceder,
    }
    
    # adiciona os dados lista global e salva
    dados_atividades.append(nova_atividade_dados)
    salvar_dados(NOME_ARQUIVO_DADOS_ATIVIDADES, dados_atividades)
    
    # cria o widget
    criar_widget_atividade(nova_atividade_dados)

    # limpa os campos
    nome_var.set("")
    pessoas_var.set("")
    grupo_var.set("")
    vazio_var.set(0)
    excede_var.set(0)

# janela link

def abrir_janela_link(dados_atividade):
    janela_link = ctk.CTkToplevel(root)
    janela_link.geometry("500x500")
    janela_link.title(f"Links - {dados_atividade['nome']}")
    janela_link.grab_set()
    janela_link.focus_set()

    nome_atividade_atual = dados_atividade["nome"]
    ctk.CTkLabel(janela_link, text=f"Gerenciar links - {nome_atividade_atual}", font=ctk.CTkFont(weight="bold")).pack(pady=10)

    # Dicionário para guardar as variáveis de cada checkbox { "NomeAtividade": IntVar }
    vars_checkbox = {}

    # Frame rolável para os links
    scroll_links = ctk.CTkScrollableFrame(janela_link, height=300)
    scroll_links.pack(pady=10, padx=10, fill="both", expand=True)

    # Criar checkboxes para todas as outras atividades
    for atividade in dados_atividades:
        nome_outra = atividade["nome"]
        if nome_outra != nome_atividade_atual:
            var = tkinter.IntVar()
            
            # Se já tiver link no JSON já marca o checkbox
            links_atuais = dados_atividade.get("link", "").split(", ")
            if nome_outra in links_atuais:
                var.set(1)
            
            vars_checkbox[nome_outra] = var

            f = ctk.CTkFrame(scroll_links, fg_color="transparent")
            f.pack(fill="x", pady=2)
            ctk.CTkLabel(f, text=nome_outra).pack(side="left", padx=10)
            ctk.CTkCheckBox(f, text="Linkar", variable=var).pack(side="right", padx=10)

    def salvar_links():
        selecionados = [nome for nome, v in vars_checkbox.items() if v.get() == 1]
        
        # Atualizar a atividade ATUAL
        dados_atividade["link"] = ", ".join(selecionados)

        # Atualizar as outras atividades selecionadas
        for atividade in dados_atividades:
            nome_outra = atividade["nome"]
            
            # Se a outra atividade foi selecionada no checkbox
            if nome_outra in selecionados:
                links_da_outra = [l.strip() for l in atividade.get("link", "").split(",") if l.strip()]
                if nome_atividade_atual not in links_da_outra:
                    links_da_outra.append(nome_atividade_atual)
                atividade["link"] = ", ".join(links_da_outra)
            
            # Se ela NÃO está selecionada, mas o nome da atual estava lá, removemos (Desvincular)
            elif nome_outra in vars_checkbox: # Só mexe se for uma das que apareceu na lista
                links_da_outra = [l.strip() for l in atividade.get("link", "").split(",") if l.strip()]
                if nome_atividade_atual in links_da_outra:
                    links_da_outra.remove(nome_atividade_atual)
                    atividade["link"] = ", ".join(links_da_outra)

        # Salvar arquivo JSON
        salvar_dados(NOME_ARQUIVO_DADOS_ATIVIDADES, dados_atividades)
        CTkMessagebox(title="Sucesso", message="Links atualizados com sucesso!", icon="check")
        janela_link.destroy()

    # Botões
    btn_frame = ctk.CTkFrame(janela_link, fg_color="transparent")
    btn_frame.pack(fill="x", side="bottom", padx=10, pady=10)
    ctk.CTkButton(btn_frame, text="Salvar", command=salvar_links).pack(side="right", padx=5)
    ctk.CTkButton(btn_frame, text="Cancelar", command=janela_link.destroy).pack(side="right", padx=5)



# INTERFACE

root = ctk.CTk()
root.geometry("900x700")
root.title("Distribuição Automatizada")
validar_entrada_numero = root.register(validar_numero)

# Variáveis de entrada
nome_var = StringVar()
pessoas_var = StringVar()
grupo_var = StringVar()
vazio_var = tkinter.IntVar()
excede_var = tkinter.IntVar()
link_var = tkinter.IntVar()

# frame com campos do formulário
frame_entrada = ctk.CTkFrame(root, corner_radius=10)
frame_entrada.pack(pady=10, padx=20, fill="x")

# entradas
ctk.CTkLabel(frame_entrada, text="Nome da atividade:").pack(padx=10, pady=2, anchor="w")
ctk.CTkEntry(frame_entrada, textvariable=nome_var).pack(padx=10, pady=2, fill="x")

ctk.CTkLabel(frame_entrada, text="Quantidade de pessoas:").pack(padx=10, pady=2, anchor="w")
qntd_pessoas = ctk.CTkEntry(frame_entrada, textvariable=pessoas_var, validate="key", validatecommand=(validar_entrada_numero, "%d", "%P"))
qntd_pessoas.pack(padx=10, pady=2, fill="x")

ctk.CTkLabel(frame_entrada, text="Grupo de prioridade:").pack(padx=10, pady=2, anchor="w")
grupo_prioridade = ctk.CTkEntry(frame_entrada, textvariable=grupo_var, validate="key", validatecommand=(validar_entrada_numero, "%d", "%P"))
grupo_prioridade.pack(padx=10, pady=2, fill="x")


frame_checkboxes = ctk.CTkFrame(frame_entrada, fg_color="transparent")
frame_checkboxes.pack(fill="x", padx=10, pady=5)

ctk.CTkCheckBox(frame_checkboxes, text="Pode ficar vazio", variable=vazio_var).pack(side="left", padx=(0, 20))

ctk.CTkCheckBox(frame_checkboxes, text="Pode exceder", variable=excede_var).pack(side="left")

ctk.CTkButton(frame_checkboxes, text='Criar Atividade', command=criar_nova_atividade).pack(side="right", padx=20)

frame_escala = ctk.CTkFrame(frame_entrada, fg_color="transparent")
frame_escala.pack(padx=10, pady=5)

ctk.CTkButton(frame_escala, text='Importar escala', command=importar_escala).pack(side="left", padx=10)

ctk.CTkButton(frame_escala, text='Baixar escala modelo', command=baixar_escala_modelo).pack(side="left")

nome_arquivo = ctk.CTkLabel(frame_entrada, text="")
nome_arquivo.pack(anchor="center")


# Lista de atividades com cabeçalho fixo
frame_container_lista = ctk.CTkFrame(root, fg_color="transparent")
frame_container_lista.pack(pady=10, padx=20, fill="both", expand=True)

# Título
ctk.CTkLabel(
    frame_container_lista, 
    text="Lista de Atividades", 
    font=ctk.CTkFont(size=16, weight="bold")
).pack(pady=(0, 5), anchor="center")

# Cabeçalho Fixo
frame_cabecalho = ctk.CTkFrame(frame_container_lista, fg_color="#2b2b2b", corner_radius=5)
frame_cabecalho.pack(fill="x", pady=5)

subtitulos = ["NOME", "PESSOAS", "GRUPO", "VAZIO", "EXCEDE", "LINK"]
larguras = [130, 70, 130, 70, 130, 25]

for i, titulo in enumerate(subtitulos):
    label = ctk.CTkLabel(
        frame_cabecalho, 
        text=titulo, 
        font=ctk.CTkFont(size=11, weight="bold"),
        width=larguras[i],
        anchor="center"
    )
    label.grid(row=0, column=i, padx=10, pady=5)

# Frame de Atividades
frame_atividades = ctk.CTkScrollableFrame(frame_container_lista, height=200)
frame_atividades.pack(fill="both", expand=True)

gerar_dist = ctk.CTkButton(root, text="Gerar Distribuição", command=gerar_distribuicao)
gerar_dist.pack(pady=(0, 10))


# INICIALIZAÇÃO

# Carrega os dados do arquivo ao iniciar
dados_atividades = carregar_dados(NOME_ARQUIVO_DADOS_ATIVIDADES)
dados_escala = carregar_dados(NOME_ARQUIVO_DADOS_ESCALA)
dados_arquivo = carregar_dados_arquivo(NOME_ARQUIVO_DADOS_ARQUIVO, nome_arquivo)

# refaz a interface com os dados carregados
for dados in dados_atividades:
    criar_widget_atividade(dados)

root.mainloop()
