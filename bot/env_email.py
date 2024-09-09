import win32com.client as win32
import pandas as pd

teste = pd.read_excel("xlsx/emprestimo.xlsx")

if 'Status' in teste.columns:
    test = teste[teste['Status'] == "Não Entregue"]
    if len(test) >= 1:
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
    else:
        print("Não foi encontrado nenhum equipamento não entregue")
else:
    print("A coluna 'STATUS' não foi encontrada.")

    

