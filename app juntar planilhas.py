from tkinter import *
from pathlib import Path
import pandas as pd

class Aplicativo_juntar_excel:
    def __init__(self, master = None):
        #conteiner cabeçalho
        self.fontePadrao = ("Arial", "10")
        self.primeiroContainer = Frame(master)
        self.primeiroContainer["pady"] = 10
        self.primeiroContainer.pack()

        #conteiner label endereço
        self.segundoContainer = Frame(master)
        self.segundoContainer["padx"] = 20
        self.segundoContainer.pack()
        
        #conteiner input endereço
        self.terceiroContainer = Frame(master)
        self.terceiroContainer["padx"] = 20
        self.terceiroContainer.pack()
        
        #conteiner label nome do arquivo
        self.quintoConteiner = Frame(master)
        self.quintoConteiner["padx"] = 20
        self.quintoConteiner.pack()
        
        #conteiner input nome do arquivo
        self.sextoConteiner = Frame(master)
        self.sextoConteiner["padx"] = 20
        self.sextoConteiner.pack()

        #conteiner botão
        self.quartoContainer = Frame(master)
        self.quartoContainer["pady"] = 20
        self.quartoContainer.pack()

        #label que aparece no cabeçalho da pagina
        self.titulo = Label(self.primeiroContainer, text="Juntar tabelas excel!")
        self.titulo["font"] = ("Arial", "10", "bold")
        self.titulo.pack()
        
        #label que aparece antes da input endereço
        self.nomeLabel = Label(self.segundoContainer,text="Endereço da pasta com as planilhas", font=self.fontePadrao)
        self.nomeLabel.pack(side=LEFT)
        
        #input endereço
        self.endeco = Entry(self.terceiroContainer)
        self.endeco["width"] = 30
        self.endeco["font"] = self.fontePadrao
        self.endeco.pack(side=LEFT)
        
        #label que aparece antes da input nome do arquivo
        self.nomeLabel = Label(self.quintoConteiner,text="Nome do Arquivo", font=self.fontePadrao)
        self.nomeLabel.pack(side=LEFT)
        
        #input nome do arquivo
        self.nomeArquivo = Entry(self.sextoConteiner)
        self.nomeArquivo["width"] = 30
        self.nomeArquivo["font"] = self.fontePadrao
        self.nomeArquivo.pack(side=LEFT)

        #botão juntar
        self.autenticar = Button(self.quartoContainer)
        self.autenticar["text"] = "Agregar dados"
        self.autenticar["font"] = ("Calibri", "8")
        self.autenticar["width"] = 12
        self.autenticar["command"] = self.juntarPlanilhas
        self.autenticar.pack()

        self.mensagem = Label(self.quartoContainer, text="", font=self.fontePadrao)
        self.mensagem.pack()

    #Método verificar senha
    def juntarPlanilhas(self):
        
        #pegar o endereço informado pelo usuario
        endereco = self.endeco.get()
        #pegar nome do arquivo
        nomeArquivo = self.nomeArquivo.get()
        
        #print(endereco)
        #print(nomeArquivo)
        
        
        #script para juntar dados das planilhas
        tb_final = pd.DataFrame()
        #digite o caminho do arquivo
        caminho = Path(endereco)
        
        
        for c in caminho.glob('*'):
            if(c.suffix == '.xlsx'):
                #para ler todas as abas basta no parametro de planilha não passar o nome da planilha e sim none
                df=pd.read_excel(c, sheet_name=None)
                df = pd.concat(df)
                tb_final = pd.concat([tb_final, df])
       
        
        
            
            
        #digite o caminho onde salvar       
        tb_final.to_excel(endereco+"\\"+nomeArquivo+".xlsx", index=False)
                

           # self.mensagem["text"] = "Autenticado"
       




root = Tk()
Aplicativo_juntar_excel(root)
root.mainloop()
