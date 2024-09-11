import win32com.client as win32
import pandas as pd
import pyautogui
from time import sleep


def soneca():
    return sleep(1)

# Excluir arquivos de excell antigos 
pyautogui.doubleClick(x=30, y=10)
soneca()
pyautogui.doubleClick(x=450, y=310)
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
pyautogui.click(x=690, y=520)
soneca()
pyautogui.keyDown('alt')
soneca()
pyautogui.press("F4")
pyautogui.keyUp("alt")
soneca()
# Excluir arquivos de excell antigos 
# ////////////////////////////////////////////////
# abrir e fazer download do arquivo
pyautogui.press("win")
soneca()
pyautogui.write("Teams")
soneca()
pyautogui.press("enter")
sleep(8)
pyautogui.click(x=830, y=65)
soneca()
pyautogui.click(x=830, y=200)
pyautogui.scroll(-300)
soneca()
pyautogui.rightClick(x=997, y=487)
soneca()
pyautogui.click(x=1025, y=580)
soneca()
pyautogui.keyDown('alt')
soneca()
pyautogui.press("F4")
pyautogui.keyUp("alt")
# abrir e fazer download do arquivo
# ////////////////////////////////////////////////
 

dock = pd.read_excel("xlsx/emprestimo.xlsx")
if 'Status' in dock.columns:
    dock_read = dock[dock['Status'] == "Não Entregue"]
    if len(dock_read) >= 1:
        mensagem = [f""" <p>
            <span>
            <strong>Equipamento:</strong> {row['Equipamento']}
            </span>, 
            <span>
            <strong>status:</strong> {row['Status']}</span>, 
            <span>
            <strong>Sala:</strong> {row['Sala']},
            </span>
            <span><strong>Data:</strong> {row['Data']}</span>
            </p>\n""" for index, row in dock_read.iterrows()]
        mensagem_email = "\n".join(mensagem)

        print(mensagem_email)

        outlook = win32.Dispatch("outlook.application")
        email = outlook.CreateItem(0)
        email.To = "kauansantosdesouza45@gmail.com"
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

    

