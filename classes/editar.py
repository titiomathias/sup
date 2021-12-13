# Classe de Edição (Interface Gráfica)
from datetime import datetime
from PySimpleGUI import PySimpleGUI as sg
from classes import makeredit
from classes import criar

now = datetime.now()
date = f"{now.day}/{now.month}/{now.year}"

# Criando e iniciando classe
class Edicao:
    def __init__(self, file, nome):

        # Criando Interface gráfica para input das informações

        sg.theme("DarkAmber")

        layout = [
            [sg.Text("Fazer edição de arquivo:")],
            [sg.Text(f"{nome}")],
            [sg.Text("Digite a data de hoje:", size=18), sg.Input(key="data", size=34, default_text=date)],
            [sg.Text("Razão Social:", size=18), sg.Input(key="client", size=34)],
            [sg.Text("CNPJ do cliente:", size=18), sg.Input(key="cnpj", size=34)],
            [sg.Text("Modelo da OLT:", size=18), sg.Input(key="olt", size=34)],
            [sg.Text("Modelo do Equipamento:", size=18), sg.Input(key="ont", size=34)],
            [sg.Text("Serial Number:", size=18), sg.Input(key="sn", size=34)],
            [sg.Text("Versão do Firmware:", size=18), sg.Input(key="v-firmware", size=34)],
            [sg.Text("Versão do Hardware:", size=18), sg.Input(key="v-hardware", size=34)],
            [sg.Text("Dúvida/Problema:", size=18), sg.Input(key="duvida", size=34)],
            [sg.Text("Atendimento:", size=18), sg.Radio("Bem-sucedido", "opt", key=1), sg.Radio("Mal-sucedido", "opt", key=2)],
            [sg.Text("Resolução:", size=18)],
            [sg.Multiline(size=(55, 4), key='resolucao', default_text="Digite a solução para o problema aqui...")],
            [sg.Button("Pronto", size=15), sg.Button("Sair", size=15), sg.Button("Criar", size=15) ]
        ]

        janela = sg.Window("SoftGest - Editar Report", layout)

        # Estrutura de escolhas entre a forma de uso do programa: enviar, modificar modo de uso ou sair do programa.
        while True:
                # Recebendo valores da janela.
                eventos, valores = janela.read()

                # Fechar janela.
                if eventos == sg.WINDOW_CLOSED:
                    break

                # Enviar dados.
                if eventos == "Pronto":
                    nome_planilha = file
                    data = date
                    client = valores["client"]
                    cnpj = valores["cnpj"]
                    olt = valores["olt"]
                    ont = valores["ont"]
                    sn = valores["sn"]
                    v_firmware = valores["v-firmware"]
                    v_hardware = valores["v-hardware"]
                    duvida = valores["duvida"]
                    resolucao = valores["resolucao"]
                    local = valores["local"]
                    if valores[1] == True and valores[2] == False:
                        atendimento = True
                    else:
                        atendimento = False

                    # Passando valores da interface gráfica para código de execução back-end.
                    makeredit.EditarPlanilha(nome_planilha, data, client, cnpj, olt, ont, sn, v_firmware, v_hardware,
                                             duvida, resolucao, atendimento)

                # Opção de uso de criação. Importando e executando classe de criação.
                if eventos == "Criar":
                    janela.close()
                    criar.Criacao()

                # Opção de saída do programa.
                if eventos == "Sair":
                    break