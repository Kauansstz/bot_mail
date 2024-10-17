import win32com.client as win32
import pandas as pd
import pyautogui
from time import sleep
import sys
import os

if getattr(sys, 'frozen', False):
    script_dir = os.path.dirname(sys.executable)
else:
    script_dir = os.path.dirname(os.path.abspath(__file__))

file_path = os.path.join(script_dir, 'xlsx', 'emprestimo.xlsx')
dock = pd.read_excel(file_path)

def soneca():
    return sleep(2)

# Excluir arquivos de excell antigos 
pyautogui.keyDown("win")
pyautogui.press("r")
pyautogui.keyUp("win")
soneca()
pyautogui.write("C:\\Users\\nsi\\bot_mail\\dist\\xlsx")
soneca()
pyautogui.press("enter")
soneca()
pyautogui.keyDown("ctrl")
soneca()
pyautogui.press("a")
pyautogui.keyUp("ctrl")
soneca()
pyautogui.keyDown("shift")
soneca
pyautogui.press("del")
pyautogui.keyUp("shift")
soneca()
pyautogui.press("enter")
soneca()
pyautogui.keyDown('alt')
soneca()
pyautogui.press("F4")
pyautogui.keyUp("alt")
soneca()
# Excluir arquivos de excell antigos 
# ////////////////////////////////////////////////
# Começo de abrir e fazer download do arquivo
pyautogui.press("win")
soneca()
pyautogui.write("Teams")
soneca()
pyautogui.press("enter")
sleep(30)
pyautogui.click(x=860, y=70)
sleep(30)
pyautogui.click(x=850, y=450)
soneca()
pyautogui.scroll(-450)
soneca()
pyautogui.rightClick(x=1000, y=598)
soneca()
pyautogui.click(x=1020, y=460)
sleep(10)
pyautogui.keyDown('alt')
soneca()
pyautogui.press("F4")
pyautogui.keyUp("alt")
# Fim do  abrir e fazer download do arquivo
# ////////////////////////////////////////////////
# Começo de Renomear o arquivo
soneca()
pyautogui.keyDown("win")
pyautogui.press("r")
pyautogui.keyUp("win")
soneca()
pyautogui.write("C:\\Users\\nsi\\bot_mail\\dist\\xlsx")
soneca()
pyautogui.press("enter")
soneca()
pyautogui.keyDown("ctrl")
pyautogui.press("a")
pyautogui.keyUp("ctrl")
soneca()
pyautogui.press("F2")
soneca()
pyautogui.write("emprestimo")
soneca()
pyautogui.press("enter")
soneca()
pyautogui.keyDown('alt')
soneca()
pyautogui.press("F4")
pyautogui.keyUp("alt")
# Fim do Renomear o arquivo
# ////////////////////////////////////////////////
 
 
if 'Status' in dock.columns:
    dock_read = dock[dock['Status'] == "Não Entregue"]
    if len(dock_read) >= 1:
        mensagem = [f""" <p>
            <span>
            <strong>Equipamento:</strong> {row['Equipamento']}
            </span>, 
            <span>
            <strong>Sala:</strong> {row['Sala']},
            </span>
            <span>
            <strong>Técnico:</strong> {row['Tecnico']}
            </span>,
            <span>
            <strong>Responsavel:</strong> {row['Responsavel']}
            </span>,
            <span>
            <strong>Data:</strong> {row['Data']}
            </span>,
            <span>
            <strong>Status:</strong> {row['Status']}
            </span>, 
            </p>\n""" for index, row in dock_read.iterrows()]
        mensagem_email = "\n".join(mensagem)
        outlook = win32.Dispatch("outlook.application")
        email = outlook.CreateItem(0)
        email.To = "kauan.souza@animaeducacao.com.br; victor.bittencourt@animaeducacao.com.br; marcelo.fraiberg@animaeducacao.com.br"
        email.Subject = "Relatório de Emprestimo"
        email.HTMLBody = f"""
            <h2>Relátorio de Salas</h2>
            {mensagem_email}
        """

        email.Send()
        print("Email Enviado")
    else:
        print("Não foi encontrado nenhum equipamento 'não entregue'.")
else:
    print("A coluna 'STATUS' não foi encontrada.")

    

