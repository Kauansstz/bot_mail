import win32com.client as win32
import pandas as pd

# Ler o arquivo Excel local
teste = pd.read_excel("xlsx/CHECKLIST_SALAS_2024 (Salvo automaticamente).xlsx")




if 'STATUS' in teste.columns:
    test = teste[teste['STATUS'] == "CONCLUÍDA"]
else:
    print("A coluna 'STATUS' não foi encontrada.")


mensagem = [f""" <p>
            <span>
            <strong>Nome do espaço:</strong> {row['NOME DO ESPAÇO']}
            </span>, 
            <span>
            <strong>status:</strong> {row['STATUS']}</span>, 
            <span>
            <strong>Capacidade:</strong> {row['CAPACIDADE']},
            </span>, 
            <span><strong>Modelo do Computador:</strong> {row['COMPUTADO MODELO']}</span>
            </p>\n""" for index, row in test.iterrows()]
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