import pandas as pd
import pywhatkit as wa
from tkinter import *
from tkinter.filedialog import askopenfilename
from pkg_resources import resource_filename

 
def escolher_arquivo():
    arquivo_escolher = askopenfilename(filetypes=[("Arquivos do Excel", ".xlsx;.xls")]) #determina o tipo de arquivo
    if arquivo_escolher:
        arquivo_escolhido.set(arquivo_escolher)
        


def enviar_mensagens():
    #Carrego a base de dados lendo uma Aba do Excel 
    df = pd.read_excel(arquivo_escolhido.get())
    df.head()
    
    #Conto quantas linhas possui aquela Aba
    qtd_linhas = df.shape[0]

    #Pego a mensagem escrita na caixa de texto
    mensagem = mensagem_texto.get("1.0", END)
    
    for x in range(qtd_linhas):
        celular = df.iloc[x, 2]
        celular = celular.replace('(', '').replace(')', '').replace('-', '').replace(' ', '')
        wa.sendwhatmsg_instantly(f"{celular}", f"{mensagem}", 20, True, 5)


#Configurações da janela do tkinter padrao   
window = Tk()
window.title("Mensagem Automática")
window.config(padx=20, pady=20)


#Aqui adiciono o logo do CRPRS a janela tkinter
arquivo_foto = resource_filename(__name__,"logo.png")
photo = PhotoImage(file = arquivo_foto) 
logo = Label( window, image = photo, padx= 10, pady= 10).place(x = 300, y = 100) 



#Labels da janela TKinter
Label(window, text="Arquivo:").grid(row=0, column=0, pady = 10)
Label(window, text="Mensagem:").grid(row=1, column=0, pady = 10)
Label(window, text="Erion Teixeira").grid(row=6, column=0, pady = 25)


#Aqui passamos o nome do arquivo escolhido para o StringVar que guarda o valor do Entry 
arquivo_escolhido = StringVar()


#Entradas do arquivo 
entrada_arquivo = Entry(window, textvariable=arquivo_escolhido, width=30).grid(row=0, column=1, pady=10)


#Caixa de texto para a mensagem
mensagem_texto = Text(window, height=6, width=30)
mensagem_texto.grid(row=1, column=1, pady=10)

#Botões adicionar e arquivo/enviar mensagem
botao_adicionar = Button(window, text="Adicionar Arquivo", width=15, command=escolher_arquivo).grid(row=0, column=2, pady=10)
botao_enviar = Button(window, text="Enviar Mensagens", command=enviar_mensagens).grid(row=2, column=1, pady=10)


#Configuração de tamanho minimo e maximo da janela TKinter 
window.maxsize(500,275)
window.minsize(500,275)
window.mainloop()