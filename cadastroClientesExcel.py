import tkinter as tk
import openpyxl

# Função para salvar os dados na planilha do Excel
def salvar_dados():
    ## Abrir planilha do Excel e selecionar planilha ativa
    workbook = openpyxl.load_workbook('clientes.xlsx')
    sheet = workbook.active

    ### Inserir os dados dos clientes
    nome = nome_entry.get()
    cpf = cpf_entry.get()
    email = email_entry.get()
    telefone = telefone_entry.get()
    cidade = cidade_entry.get()
    estado = estado_entry.get()

    sheet.append([nome, cpf, email, telefone, cidade, estado])

    #### Salvar a planilha
    workbook.save('clientes.xlsx')

    ##### Limpar os campos de entrada
    nome_entry.delete(0, tk.END)
    cpf_entry.delete(0, tk.END)
    email_entry.delete(0, tk.END)
    telefone_entry.delete(0, tk.END)
    cidade_entry.delete(0, tk.END)
    estado_entry.delete(0, tk.END)

###### Criar a janela principal de cadastro
janela = tk.Tk()
janela.title('Cadastro de Clientes')
janela.geometry('400x300')

####### Campos de entrada dos dados
nome_label = tk.Label(janela, text='Nome:')
nome_label.grid(row=0, column=0, padx=10, pady=10)

nome_entry = tk.Entry(janela)
nome_entry.grid(row=0, column=1, padx=10, pady=10)

cpf_label = tk.Label(janela, text='CPF:')
cpf_label.grid(row=1, column=0, padx=10, pady=10)

cpf_entry = tk.Entry(janela)
cpf_entry.grid(row=1, column=1, padx=10, pady=10)

email_label = tk.Label(janela, text='Email:')
email_label.grid(row=2, column=0, padx=10, pady=10)

email_entry = tk.Entry(janela)
email_entry.grid(row=2, column=1, padx=10, pady=10)

telefone_label = tk.Label(janela, text='Telefone:')
telefone_label.grid(row=3, column=0, padx=10, pady=10)

telefone_entry = tk.Entry(janela)
telefone_entry.grid(row=3, column=1, padx=10, pady=10)

cidade_label = tk.Label(janela, text='Cidade:')
cidade_label.grid(row=4, column=0, padx=10, pady=10)

cidade_entry = tk.Entry(janela)
cidade_entry.grid(row=4, column=1, padx=10, pady=10)

estado_label = tk.Label(janela, text='Estado:')
estado_label.grid(row=5, column=0, padx=10, pady=10)

estado_entry = tk.Entry(janela)
estado_entry.grid(row=5, column=1, padx=10, pady=10)

######## Criar o botão de salvar
salvar_button = tk.Button(janela, text='Salvar', command=salvar_dados)
salvar_button.grid(row=6, column=1, padx=10, pady=10)

######## Iniciar a janela principal
janela.mainloop()