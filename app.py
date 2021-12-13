# EXECUÇÃO COMPLETA DO SOFTWARE
import os
import tkinter.filedialog
from PySimpleGUI import PySimpleGUI as sg
from classes import criar
from classes import editar

sg.theme("DarkAmber")

layout = [
    [sg.Text("SOFTGEST - Relatory Maker")],
    [sg.Radio("Criar Planilha", "opt", key=1),sg.Radio("Editar Planilha", "opt", key=2, default=True)],
    [sg.Button("Iniciar")]
]

wind = sg.Window("SoftGest - Report Maker", layout)

# Estrutura de escolhas entre a forma de uso do programa: criação ou edição de documento.
while True:
    eventos, valores = wind.read()
    if eventos == sg.WINDOW_CLOSED:
        break

    if eventos == "Iniciar":

        # Usuário escolher a opção de criar um novo documento.
        if valores[1] == True and valores[2] == False:
            wind.close()
            criar.Criacao()     # Importando a classe de criação.

        # Usuário escolher a opção de criar um novo documento.
        else:
            wind.close()
            file = tkinter.filedialog.askopenfilename()     #       Importando a classe tkinter para selecionar o arquivo para edição.
            nome = os.path.basename(file)
            editar.Edicao(file, nome)        #           # Importando a classe de edição.
