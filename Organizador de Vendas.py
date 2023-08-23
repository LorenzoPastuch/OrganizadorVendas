import pandas as pd
import tkinter as tk
from pathlib import Path
from tkinter.filedialog import askdirectory, askopenfilename


class pop_ups:

    def __init__(self):
        pass


    def concluido(self):
        tk.messagebox.showinfo(title="Concluído", message="Criação do(s) arquivo(s) concluida")


    def error(self):
        tk.messagebox.showerror(title="Erro", message="Arquivos selecionados inválidos")


    def error2(self):
        tk.messagebox.showerror(title="Erro", message="Lista completa não criada")


    def error3(self):
        tk.messagebox.showerror(title="Erro", message="Diretório de salvamento não selecionado")

class Organizar:

    def __init__(self):
        self.diretorio_vendas = tk.StringVar()
        self.fornecedores = tk.StringVar()
        self.diretorio_salvar = tk.StringVar()


    def selecionar_diretorio(self):
        caminho = askdirectory(title="Selecionar diretório")
        self.diretorio_vendas.set(caminho)
        self.diretorio_vendas = Path(self.diretorio_vendas.get())
        if caminho:
            text_selec["text"] = caminho


    def selecionar_fornecedores(self):
        arquivo = askopenfilename(title="Selecionar arquivo")
        self.fornecedores.set(arquivo)
        self.fornecedores = Path(self.fornecedores.get())
        if arquivo:
            text_selec_sup["text"] = arquivo


    def selecionar_save(self):
        save = askdirectory(title="Selecionar diretório")
        self.diretorio_salvar.set(save)
        self.diretorio_salvar = Path(self.diretorio_salvar.get())
        if save:
            dir_save["text"] = save

    def juntar_arquivos(self):  # Junta os arquivos contidos no diretorio de vendas por mes
        try:
            lista_venda_total = pd.DataFrame()  # Dataframe vazio
            for arquivo in (self.diretorio_vendas.iterdir()):  # Pega todos arquivos contido no diretorio
                lista_venda = pd.read_excel(Path(self.diretorio_vendas).joinpath(arquivo))  # Monta o caminho do arquivo
                lista_venda_total = pd.concat([lista_venda_total, lista_venda])  # Concatena as colunas de acordo com o nome

            lista_venda_total["Codigo"] = pd.to_numeric(lista_venda_total["Codigo"], errors="coerce", downcast="integer")  # Converte linhas não numericas para vazias
            lista_venda_total = lista_venda_total.dropna()  # remove linhas vazias
            lista_venda_total = lista_venda_total.reset_index(drop=True)  # renumera o índice
            lista_venda_total = lista_venda_total.rename(columns={"Codigo": "Código", "Nome": "Produto", "Quant.": "Quantidade","Vl. Total": "Valor Total"})  # renomeia as colunas

            # Cria lista completa de informações
            lista_venda_total["Fornecedor"] = None  # adiciona coluna
            lista_venda_total["Quantidade de Pedidos"] = 1
            self.lista_completa = lista_venda_total.groupby(by=["Código", "Produto", "Fornecedor"], as_index=False, dropna=False).sum()  # Agrupa conforme o código e soma as quantidades e soma as quantidades

            fornecedores = pd.read_excel(self.fornecedores)

            def extrair_fornecedor(produto_fornecedor):  # Extrai o fornecedor do produto e adiciona na coluna fornecedor
                produto_fornecedor_lower = produto_fornecedor.lower()
                for fornecedor in fornecedores["Nome"]:
                    if fornecedor.lower() in produto_fornecedor_lower:
                        return fornecedor
                return None

            self.lista_completa["Fornecedor"] = self.lista_completa["Produto"].apply(extrair_fornecedor)  # Aplica o metodo extrair_fornecedor para cada produto
            self.lista_completa["Fornecedor"] = self.lista_completa["Fornecedor"].fillna("Sem Fornecedor")  # Preenche as colunas vazias

            self.lista_completa = self.lista_completa.sort_values(by=["Fornecedor", "Quantidade"], ascending=[True, False])  # Organiza em ordem alfabética e decrescente


            self.lista_completa.to_excel(Path(self.diretorio_salvar).joinpath("Lista Completa.xlsx"), index=False)
            pop_ups.concluido(self)
        except:
            if not self.diretorio_salvar.get():
                pop_ups.error3(self)
            else:
                pop_ups.error(self)


    def filtrar_fornecedor(self):  # filtra cada fornecedor e salva em uma planilha separada
        try:
            caminho_salvar = Path(self.diretorio_salvar).joinpath("Listas por Fornecedor")
            if not Path(caminho_salvar).exists():  # cria a pasta de listas por fornecedor caso ela nao exista
                Path(caminho_salvar).mkdir()

            for fornecedor in self.lista_completa["Fornecedor"].unique():
                lista_fornecedor = self.lista_completa.loc[self.lista_completa["Fornecedor"] == fornecedor]
                soma = lista_fornecedor["Quantidade"].sum()
                lista_fornecedor.loc[lista_fornecedor.index[0], "Total Vendido"] = soma
                if r"/" in fornecedor:
                    fornecedor = fornecedor.replace("/", "-")
                nome_arquivo = f"{soma} {fornecedor}.xlsx"
                arquivo_salvar = Path(caminho_salvar).joinpath(nome_arquivo)
                lista_fornecedor.to_excel(arquivo_salvar, index=False)
            pop_ups.concluido(self)
        except:
            pop_ups.error2(self)


janela = tk.Tk()

organizador = Organizar()

janela.geometry("600x525+100+100")
janela.title("Organizador de Vendas")
janela.iconbitmap("icon.ico")

title = tk.Label(
    text="Organizador de Vendas",
    bg="#6EBAF8",
    width=30,
    height=3,
    font=("algerian",20),
    padx=10,
    pady=10,
    relief="ridge",
    borderwidth=20
)
title.grid(
    row=0,
    column=0,
    columnspan=2,
    sticky="nsew",
    padx=10,
    pady=10
)

text_sell = tk.Label(
    text="Selecione o diretório de arquivos contendo as vendas",
    anchor="w"
)
text_sell.grid(
    row=1,
    column=1,
    sticky="nsew",
    padx=10,
    pady=10
)

text_selec = tk.Label(
    text="Nenhum diretório selecionado",
    anchor="w"
)
text_selec.grid(
    row=2,
    column=0,
    columnspan=2,
    sticky="nsew",
    padx=10,
    pady=10
)
dir_sell_button = tk.Button(
    text="Selecionar",
    command=organizador.selecionar_diretorio,
)
dir_sell_button.grid(
    row=1,
    column=0,
    sticky="nsew",
    padx=10,
    pady=10
)

text_sup = tk.Label(
    text="Selecione o arquivo contendo os fornecedores",
    anchor="w"
)
text_sup.grid(
    row=3,
    column=1,
    sticky="nsew",
    padx=10,
    pady=10
)

text_selec_sup = tk.Label(
    text="Nenhum arquivo selecionado",
    anchor="w"
)
text_selec_sup.grid(
    row=4,
    column=0,
    columnspan=2,
    sticky="nsew",
    padx=10,
    pady=10
)

dir_sup_button = tk.Button(
    text="Selecionar",
    command=organizador.selecionar_fornecedores
)
dir_sup_button.grid(
    row=3,
    column=0,
    sticky="nsew",
    padx=10,
    pady=10
)

save = tk.Label(
    text="Selecione o diretório em que deseja salvar os arquivos",
    anchor="w"
)
save.grid(
    row=5,
    column=1,
    sticky="nsew",
    padx=10,
    pady=10
)

save_button = tk.Button(
    text="Selecionar",
    command=organizador.selecionar_save
)
save_button.grid(
    row=5,
    column=0,
    sticky="nsew",
    padx=10,
    pady=10
)

dir_save = tk.Label(
    text="Nenhum diretório selecionado",
    anchor="w"
)
dir_save.grid(
    row=6,
    column=0,
    columnspan=2,
    sticky="nsew",
    padx=10,
    pady=10
)

create_button = tk.Button(
    text="Criar lista completa de vendas",
    command=organizador.juntar_arquivos
)
create_button.grid(
    row=7,
    column=0,
    columnspan=2,
    padx=10,
    pady=10,
    sticky="nsew"
)

create_button_sup = tk.Button(
    text="Criar listas por fornecedor",
    command=organizador.filtrar_fornecedor
)
create_button_sup.grid(
    row=8,
    column=0,
    columnspan=2,
    padx=10,
    pady=10,
    sticky="nsew"
)

janela.mainloop()
