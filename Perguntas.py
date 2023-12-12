import tkinter as tk
from tkinter import simpledialog
import openpyxl
import pandas as pd

# Inicialização da janela principal
root = tk.Tk()
root.title("Questionário")

largura = root.winfo_screenwidth() // 2
altura = root.winfo_screenheight() // 2
pos_x = (root.winfo_screenwidth() - largura) // 2
pos_y = (root.winfo_screenheight() - altura) // 2

root.geometry(f"{largura}x{altura}+{pos_x}+{pos_y}")

frame = tk.Frame(root)
frame.pack(padx=10, pady=10)

# Criando um dicionário vazio para as respostas
respostas = {}

# Definindo as perguntas e opções de resposta
perguntas = {
    "1. Local onde a pesquisa foi aplicada (preenchimento interno)": [
        "(1) Aeroporto",
        "(2) Ponto turístico do ParNaMar",
        "(3) Pousada",
        "(4) Porto"
    ],
    "2. Qual sua nacionalidade?": [
        "(1) Brasileira",
        "(2) Francesa",
        "(3) Argentina",
        "(4) Mexicana",
        "(5) Chilena",
        "( ) Outros"
    ],
    "3. Qual seu gênero?": [
        "(1) Feminino",
        "(2) Masculino",
        "(3) Não binário",
        "(4) Prefiro não declarar",
        "( ) Outros"
    ],
    "4. Qual é sua faixa etária?": [
        "(1) Menor de 18",
        "(2) 18 a 25 anos",
        "(3) 26 a 35 anos",
        "(4) 36 a 45 anos",
        "(5) 46 a 55 anos",
        "(6) Acima de 56 anos"
    ],
    "5. Com quem você está viajando?": [
        "(1) Sozinho(a)",
        "(2) Com a família",
        "(3) Com meu companheiro(a)",
        "(4) Com amigos"
    ],
    "6. Qual o motivo da sua estadia?": [
        "(1) Turismo",
        "(2) Trabalho",
        "(3) Celebração",
        "( ) Outros"
    ],
    "7. Quanto tempo você tem de estadia na ilha?": [
        "(1) 1 a 3 dias",
        "(2) 4 a 6 dias",
        "(3) 7 a 10 dias",
        "(4) Mais de 10 dias"
    ],
    "8. Quantas vezes você já visitou Fernando de Noronha?": [
        "(1) É minha primeira vez",
        "(2) Duas vezes",
        "(3) Três vezes",
        "(4) Quatro vezes",
        "(5) Mais de 4 vezes"
    ],
    "9. Você conhece o Parque Nacional Marinho de Fernando de Noronha?": [
        "(1) Sim",
        "(2) Não"
    ],
    "10. Você sabe o que é TPA (Taxa de Preservação Ambiental)?": [
        "(1) Sim",
        "(2) Não"
    ],
    "11. Você entende a diferença entre a TPA e o ingresso do Parque Nacional Marinho Fernando de Noronha, cobrado pela Econoronha?": [
        "(1) Sim",
        "(2) Não"
    ],
    "12. Você comprou o seu ingresso de forma presencial ou online?": [
        "(1) Online",
        "(2) Presencial"
    ],
    "13. Como você avalia o valor cobrado pelo ingresso do Parque Nacional Marinho Fernando de Noronha?": [
        "(1) Muito caro",
        "(2) Caro",
        "(3) Justo",
        "(4) Barato",
        "(5) Muito barato"
    ],
    "14. Quais pontos turísticos você já visitou em Fernando de Noronha?": [
        "(1) PIC (Praia) do Sancho",
        "(2) Baía dos Porcos",
        "(3) PIC (Praia) do Leão",
        "(4) Trilha da Atalaia",
        "(5) Enseada dos Abreus",
        "(6) PIC (Baía) do Sueste",
        "(7) Baía dos Golfinhos",
        "(8) Capim Açu",
        "(9) Piscina Natural São José",
        "(10) Ponta das Caracas",
        "(11) Pontinha Pedra Alta",
        "(12) Caiera",
        "(13) Trilha São Joaquim",
        "( ) Outros"
    ],
    "15. Você sabia que esses pontos turísticos fazem parte do Parque Nacional Marinho?": [
        "(1) Sim",
        "(2) Não"
    ],
    "16. Como você avalia o serviço de agendamento de passeios?": [
        "(1) Muito bom",
        "(2) Bom",
        "(3) Regular",
        "(4) Ruim",
        "(5) Muito ruim"
    ],
    "17. Você agendou alguma trilha?": [
        "(1) Sim",
        "(2) Não"
    ],
    "18. Se sim, qual foi o meio de agendamento?": [
        "(1) Online",
        "(2) Presencial"
    ],
    "19. Como você avalia o serviço de agendamento de trilhas?": [
        "(1) Muito bom",
        "(2) Bom",
        "(3) Regular",
        "(4) Ruim",
        "(5) Muito ruim"
    ],
    "20. Como você avalia o Parque Nacional Marinho em relação à sinalização?": [
        "(1) Muito bom",
        "(2) Bom",
        "(3) Regular",
        "(4) Ruim",
        "(5) Muito ruim"
    ],
    "21. Como você avalia o Parque Nacional Marinho em relação à segurança?": [
        "(1) Muito bom",
        "(2) Bom",
        "(3) Regular",
        "(4) Ruim",
        "(5) Muito ruim"
    ],
    "22. Como você avalia o Parque Nacional Marinho em relação à preservação?": [
        "(1) Muito bom",
        "(2) Bom",
        "(3) Regular",
        "(4) Ruim",
        "(5) Muito ruim"
    ],
    "23. Quais desses pontos turísticos foram os melhores e os piores em relação a sinalização? MELHOR (Selecionar até 3 opções abaixo):": [
        "(1) PIC (Praia) do Sancho",
        "(2) Baía dos Porcos",
        "(3) PIC (Praia) do Leão",
        "(4) Trilha da Atalaia",
        "(5) Enseada dos Abreus",
        "(6) PIC (Baía) do Sueste",
        "(7) Baía dos Golfinhos",
        "(8) Capim Açu",
        "(9) Piscina Natural São José",
        "(10) Ponta das Caracas",
        "(11) Pontinha Pedra Alta",
        "(12) Caiera",
        "(13) Trilha São Joaquim"
    ],
    "PIOR (Selecionar até 3 opções abaixo):": [
        "(1) PIC (Praia) do Sancho",
        "(2) Baía dos Porcos",
        "(3) PIC (Praia) do Leão",
        "(4) Trilha da Atalaia",
        "(5) Enseada dos Abreus",
        "(6) PIC (Baía) do Sueste",
        "(7) Baía dos Golfinhos",
        "(8) Capim Açu",
        "(9) Piscina Natural São José",
        "(10) Ponta das Caracas",
        "(11) Pontinha Pedra Alta",
        "(12) Caiera",
        "(13) Trilha São Joaquim"
    ],
    "24. Quais desses pontos turísticos foram os melhores e os piores em relação à segurança? MELHOR (Selecionar até 3 opções abaixo):": [
        "(1) PIC (Praia) do Sancho",
        "(2) Baía dos Porcos",
        "(3) PIC (Praia) do Leão",
        "(4) Trilha da Atalaia",
        "(5) Enseada dos Abreus",
        "(6) PIC (Baía) do Sueste",
        "(7) Baía dos Golfinhos",
        "(8) Capim Açu",
        "(9) Piscina Natural São José",
        "(10) Ponta das Caracas",
        "(11) Pontinha Pedra Alta",
        "(12) Caiera",
        "(13) Trilha São Joaquim"
    ],
    "PIOR (Selecionar até 3 opções abaixo):": [
        "(1) PIC (Praia) do Sancho",
        "(2) Baía dos Porcos",
        "(3) PIC (Praia) do Leão",
        "(4) Trilha da Atalaia",
        "(5) Enseada dos Abreus",
        "(6) PIC (Baía) do Sueste",
        "(7) Baía dos Golfinhos",
        "(8) Capim Açu",
        "(9) Piscina Natural São José",
        "(10) Ponta das Caracas",
        "(11) Pontinha Pedra Alta",
        "(12) Caiera",
        "(13) Trilha São Joaquim"
    ],
    "25. Quais desses pontos turísticos foram os melhores e os piores em relação à manutenção? MELHOR (Selecionar até 3 opções abaixo):": [
        "(1) PIC (Praia) do Sancho",
        "(2) Baía dos Porcos",
        "(3) PIC (Praia) do Leão",
        "(4) Trilha da Atalaia",
        "(5) Enseada dos Abreus",
        "(6) PIC (Baía) do Sueste",
        "(7) Baía dos Golfinhos",
        "(8) Capim Açu",
        "(9) Piscina Natural São José",
        "(10) Ponta das Caracas",
        "(11) Pontinha Pedra Alta",
        "(12) Caiera",
        "(13) Trilha São Joaquim"
    ],
    "PIOR (Selecionar até 3 opções abaixo):": [
        "(1) PIC (Praia) do Sancho",
        "(2) Baía dos Porcos",
        "(3) PIC (Praia) do Leão",
        "(4) Trilha da Atalaia",
        "(5) Enseada dos Abreus",
        "(6) PIC (Baía) do Sueste",
        "(7) Baía dos Golfinhos",
        "(8) Capim Açu",
        "(9) Piscina Natural São José",
        "(10) Ponta das Caracas",
        "(11) Pontinha Pedra Alta",
        "(12) Caiera",
        "(13) Trilha São Joaquim"
    ],
    "26. Você contratou um guia turístico para os passeios?": [
        "(1) Sim",
        "(2) Não"
    ],
    "27. Se sim, sentiu que contribuiu para a sua viagem?": [
        "(1) Sim",
        "(2) Não"
    ],
    "28. Quais desses atendimentos você avalia como melhor?": [
        "(1) Monitores de trilha",
        "(2) Recepção no centro de visitantes",
        "(3) Lanchonetes de acesso às praias",
        "(4) Pontos de informação e controle (PIC)",
        "(5) Lojas de souvenir"
    ],
    "29. Teve algum atendimento que te deixou insatisfeito? Se sim, qual?": [
        "(1) Não",
        "( ) Outros"
    ],
    "30. Como você avalia o cuidado do parque com o meio ambiente?": [
        "(1) Muito bom",
        "(2) Bom",
        "(3) Regular",
        "(4) Ruim",
        "(5) Muito ruim"
    ],
    "31. Você sabe o que é o Pontos de Informação e Controle (PIC)?": [
        "(1) Sim",
        "(2) Não",
    ],
        "32. Você utilizou algum PIC durante sua viagem?": [
        "(1) Sim",
        "(2) Não",
    ],
        "33. Com qual finalidade você usou o PIC?": [
        "(1) Informação",
        "(2) Compras na loja de souvenir",
        "(3) Lanchonete",
        "(4) Tomar ducha, Banheiro",
        "(5) Refil da garrafa de água do Econoronha",
        "(6) Não utilizei",
        "( ) Outros",
    ],
        "34. Como você avalia a localização dos Pontos de Informação e Controle (PIC)?": [
        "(1) Muito bom",
        "(2) Bom",
        "(3) Regular",
        "(4) Ruim",
        "(5) Muito ruim",
        "(6) Não utilizei o PIC",
    ],
        "35. Se você foi ao PIC, como você avalia o atendimento no PIC?": [
        "(1) Muito bom",
        "(2) Bom",
        "(3) Regular",
        "(4) Ruim",
        "(5) Muito ruim",
        "(6) Não utilizei o PIC",
    ],
        "36. Se você foi ao PIC, como você avalia a infraestrutura no PIC?": [
        "(1) Muito bom",
        "(2) Bom",
        "(3) Regular",
        "(4) Ruim",
        "(5) Muito ruim",
        "(6) Não utilizei o PIC"
    ],
    "37. Como você avalia sua experiência no parque?": [
        "(1) Muito bom",
        "(2) Bom",
        "(3) Regular",
        "(4) Ruim",
        "(5) Muito ruim",
        "(6) Não fui ao Parque",
    ],
        "38. Como você avalia sua experiência na ilha?": [
        "(1) Muito bom",
        "(2) Bom",
        "(3) Regular",
        "(4) Ruim",
        "(5) Muito ruim",
    ],
        "39. Qual ponto turístico do Parque Nacional Marinho você mais gostou?": [
        "(1) PIC (Praia) do Sancho",
        "(2) Baía dos Porcos",
        "(3) PIC (Praia) do Leão",
        "(4) Trilha da Atalaia",
        "(5) Enseada dos Abreus",
        "(6) PIC (Baía) do Sueste",
        "(7) Baía dos Golfinhos",
        "(8) Capim Açu",
        "(9) Piscina Natural São José",
        "(10) Ponta das Caracas",
        "(11) Pontinha Pedra Alta",
        "(12) Caiera",
        "(13) Trilha São Joaquim",
    ],
        "40. Qual ponto turístico da ilha você mais gostou?": [
        "( ) Outros",
    ],
        "41. Você gostaria de adicionar alguma sugestão?": [
        "( ) Outros"
    ],
    "Sugestões": [
        "( ) Outros"
    ],
}

perguntas_list = list(perguntas.items())
num_pergunta_atual = 0
resposta_var = tk.StringVar(value="")

# Variável para armazenar a resposta "Outros"
outros_var = tk.StringVar()
outros_entry = tk.Entry(frame, textvariable=outros_var)

# Função para obter o número do formulário
def obter_numero_formulario():
    global numero_formulario
    numero_formulario = simpledialog.askstring("Número de formulário", "Insira o número de formulário:")
    if numero_formulario:
        root.withdraw()  # Esconde a janela principal
        iniciar_questionario()
    else:
        root.destroy()

# Função para iniciar o questionário
def iniciar_questionario():
    global respostas, num_pergunta_atual
    respostas = {}
    num_pergunta_atual = 0
    criar_pergunta(perguntas_list[num_pergunta_atual][0], perguntas_list[num_pergunta_atual][1])
    root.deiconify()  # Exibe a janela principal

# Função para salvar a resposta atual e passar para a próxima pergunta
def salvar_resposta(event=None):
    resposta = resposta_var.get()
    if resposta:
        pergunta = perguntas_list[num_pergunta_atual][0]
        respostas[pergunta] = resposta
        proxima_pergunta()

root.bind("<Return>", salvar_resposta)

# Função para avançar para a próxima pergunta
def proxima_pergunta():
    global num_pergunta_atual
    num_pergunta_atual += 1
    if num_pergunta_atual < len(perguntas_list):
        criar_pergunta(perguntas_list[num_pergunta_atual][0], perguntas_list[num_pergunta_atual][1])
    else:
        finalizar_questionario()

# Função para reiniciar o questionário
def reiniciar_questionario():
    global respostas, num_pergunta_atual
    respostas = {}
    num_pergunta_atual = 0
    criar_pergunta(perguntas_list[num_pergunta_atual][0], perguntas_list[num_pergunta_atual][1])

# Função para criar a interface de perguntas
def criar_pergunta(pergunta, opcoes):
    for widget in frame.winfo_children():
        widget.destroy()

    label = tk.Label(frame, text=pergunta, padx=20, pady=10)
    label.pack()

    resposta_var.set("")

    for i, opcao in enumerate(opcoes, start=1):
        if opcao.startswith("( ) Outros"):
            radio_btn = tk.Radiobutton(frame, text=opcao, variable=resposta_var, value="( ) Outros",
                                       command=exibir_caixa_outros)
            radio_btn.pack(anchor=tk.W)
            if resposta_var.get() == "( ) Outros":
                abrir_janela_texto_outros()
        else:
            radio_btn = tk.Radiobutton(frame, text=opcao, variable=resposta_var, value=opcao)
            radio_btn.pack(anchor=tk.W)

    btn_frame = tk.Frame(frame)
    btn_frame.pack()

    if num_pergunta_atual > 0:
        voltar_btn = tk.Button(btn_frame, text="Voltar", command=voltar_pergunta)
        voltar_btn.pack(side=tk.LEFT)

    if num_pergunta_atual < len(perguntas_list) - 1:
        proxima_pergunta_btn = tk.Button(btn_frame, text="Próxima Pergunta", command=salvar_resposta)
        proxima_pergunta_btn.pack(side=tk.LEFT)
    else:
        finalizar_btn = tk.Button(btn_frame, text="Finalizar", command=finalizar_questionario)
        finalizar_btn.pack(side=tk.LEFT)

# Função para exibir a caixa de texto para 'Outros'
def exibir_caixa_outros():
    if resposta_var.get() == "( ) Outros":
        abrir_janela_texto_outros()
    else:
        fechar_janela_texto_outros()

# Função para abrir a janela de texto para 'Outros'
def abrir_janela_texto_outros():
    global outros_entry_window
    outros_entry_window = tk.Toplevel(root)
    outros_entry_window.title("Resposta para 'Outros'")
    outros_var.set("")
    outros_entry = tk.Entry(outros_entry_window, textvariable=outros_var, width=50)
    outros_entry.pack(padx=20, pady=10)
    salvar_resposta_outros_btn = tk.Button(outros_entry_window, text="Salvar", command=salvar_resposta_outros)
    salvar_resposta_outros_btn.pack()
    outros_entry.bind("<Return>", lambda event: salvar_resposta_outros())
    outros_entry.focus_set()

    largura_janela_outros = 500
    altura_janela_outros = 200
    pos_x_outros = (outros_entry_window.winfo_screenwidth() - largura_janela_outros) // 2
    pos_y_outros = (outros_entry_window.winfo_screenheight() - altura_janela_outros) // 2
    outros_entry_window.geometry(f"{largura_janela_outros}x{altura_janela_outros}+{pos_x_outros}+{pos_y_outros}")

# Função para fechar a janela de texto para 'Outros'
def fechar_janela_texto_outros():
    global outros_entry_window
    if outros_entry_window is not None:
        outros_entry_window.destroy()

# Função para salvar a resposta de 'Outros'
def salvar_resposta_outros():
    resposta_outros = outros_var.get()
    if resposta_outros:
        respostas[perguntas_list[num_pergunta_atual][0]] = resposta_outros
        fechar_janela_texto_outros()
        proxima_pergunta()

# Função para voltar para a pergunta anterior
def voltar_pergunta():
    global num_pergunta_atual
    num_pergunta_atual -= 1
    criar_pergunta(perguntas_list[num_pergunta_atual][0], perguntas_list[num_pergunta_atual][1])

# Função para finalizar o questionário
def finalizar_questionario():
    nome_arquivo = f'respostas_questionario_{numero_formulario}.xlsx'
    perguntas = [pergunta[0] for pergunta in perguntas_list]
    respostas_list = [resposta for pergunta, resposta in respostas.items()]
    respostas_completas = respostas_list + [''] * (len(perguntas) - len(respostas_list))
    dados = {'Pergunta': perguntas, 'Resposta': respostas_completas}
    df = pd.DataFrame(dados)

    # Dividindo perguntas e respostas em colunas A e B do Excel
    df['Pergunta'] = df['Pergunta'].astype(str) + ":"
    df['Resposta'] = df['Resposta'].apply(lambda x: str(x).replace('[', '').replace(']', '').replace(',', '\n'))

    # Salvando em um arquivo do Excel
    df.to_excel(nome_arquivo, index=False)
    print(f"Os dados do questionário foram salvos em '{nome_arquivo}'")
    reiniciar_questionario()
    root.deiconify()

# Função para selecionar resposta usando teclas numéricas
def selecionar_resposta(event):
    key = event.char
    if key.isdigit():
        index = int(key) - 1
        if 0 <= index < len(perguntas_list[num_pergunta_atual][1]):
            resposta_var.set(perguntas_list[num_pergunta_atual][1][index])

root.bind("<Key>", selecionar_resposta)

criar_pergunta(perguntas_list[num_pergunta_atual][0], perguntas_list[num_pergunta_atual][1])

obter_numero_formulario()

root.mainloop()