import win32com.client as win32
import time
from selenium.webdriver.common.by import By
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from datetime import date, datetime
import tkinter as tk
from tkinter import simpledialog, messagebox

# Criacao da janela oculta para entrada do email e senha do aluno
root = tk.Tk()
root.withdraw()

# Coleta dos dados de login usando caixas de entrada
email_login = simpledialog.askstring("Login", "Digite seu e-mail Mackenzie:")
senha_login = simpledialog.askstring("Senha", "Digite sua senha:", show='*')

# Inicializa o Chrome com WebDriver Manager para facilitar a compatibilidade
servico = Service(ChromeDriverManager().install())
navegador = webdriver.Chrome(service=servico)
navegador.maximize_window()
navegador.get("https://ava.mackenzie.br/")  # Abre o site do Moodle Mackenzie

# Funcoes auxiliares

def existe_elemento(elemento):
    return len(elemento) > 0

def num_mes(x):  # Converte nome do mês para número
    months = {'jan': 1, 'fev': 2, 'mar': 3, 'abr': 4, 'mai': 5, 'jun': 6,
              'jul': 7, 'ago': 8, 'set': 9, 'out': 10, 'nov': 11, 'dez': 12}
    return months[x.strip()[:3].lower()]

data_atual = date.today()
time.sleep(4)  # Espera para a página carregar

# Realiza login no Moodle usando os dados do aluno
navegador.find_element("xpath", '//*[@id="i0116"]').send_keys(email_login)
navegador.find_element("xpath", '//*[@id="idSIButton9"]').click()
time.sleep(2)
navegador.find_element("xpath", '//*[@id="i0118"]').send_keys(senha_login)
navegador.find_element("xpath", '//*[@id="idSIButton9"]').click()
time.sleep(2)
navegador.find_element("xpath", '//*[@id="idSIButton9"]').click()
time.sleep(2)

# Acessa a página de calendário para encontrar atividades
navegador.find_element("xpath", '//*[@id="global_nav_calendar_link"]/div[1]').click()
time.sleep(3)

# Busca pelos elementos de tarefas e inicializa listas para os dados
tarefas = navegador.find_elements("class name", "fc-end")
lista_materias, lista_nome_atividades, lista_data = [], [], []

# Para cada tarefa encontrada, extrai os dados (matéria, nome, data)
for i in tarefas:
    try:
        i.click()
        elements = navegador.find_elements("class name", 'view_event_link')
        verifica1 = existe_elemento(elements)
        elements = navegador.find_elements("xpath", '//*[@id="event-details-trap-focus"]/div[2]/div[1]/time')
        verifica_data1 = existe_elemento(elements)

        if verifica1:
            texto = navegador.find_element("xpath", ' //*[@id="event-details-trap-focus"]/div[2]/table/tbody/tr[1]/td/a').text
            texto1 = texto.split("-")
            lista_materias.append(texto1[0])
            nom_atividade = navegador.find_element("class name", 'view_event_link').text
            lista_nome_atividades.append(nom_atividade)
            if verifica_data1:
                data = navegador.find_element("xpath", '//*[@id="event-details-trap-focus"]/div[2]/div[1]/time').text
            else:
                data = navegador.find_element("xpath", '//*[@id="event-details-trap-focus"]/div[2]/div[1]/span/time[1]').text
            lista_data.append(data)

        navegador.find_element("xpath", '//*[@id="event-details-trap-focus"]/a').click()
    except:
        continue

# Converte as datas extraídas para o formato correto (tipo date)
lista_data2 = []
for i in lista_data:
    if "," in i:
        data_separada = i.split(",")[0].split(" ")
    else:
        data_separada = i.split(" ")
    data_completa = f"{data_atual.year}-{num_mes(data_separada[1])}-{data_separada[0]}"
    data_final = datetime.strptime(data_completa, "%Y-%m-%d").date()
    lista_data2.append(data_final)

# Cria uma tabela HTML com os dados das atividades
corpo_email = """
<table border="1" cellpadding="8" cellspacing="0" style="border-collapse: collapse; width: 100%; font-family: Arial, sans-serif; font-size: 14px;">
    <thead style="background-color: #f2f2f2;">
        <tr>
            <th>#</th><th>Matéria</th><th>Atividade</th><th>Data</th><th>Status</th>
        </tr>
    </thead><tbody>
"""

# Preenche a tabela com os dados de cada atividade
for i in range(len(lista_data2)):
    materia = lista_materias[i]
    atividade = lista_nome_atividades[i]
    data_entrega = lista_data2[i].strftime("%d/%m/%Y")
    status = "<span style='color: red;'>⚠ Atrasada</span>" if lista_data2[i] < data_atual else "<span style='color: green;'>✅ Dentro do prazo</span>"
    corpo_email += f"<tr><td>{i+1}</td><td>{materia}</td><td>{atividade}</td><td>{data_entrega}</td><td>{status}</td></tr>"

corpo_email += "</tbody></table>"

# Monta e envia o e-mail pelo Outlook com o resumo das pendências
outlook = win32.Dispatch("Outlook.Application")
email = outlook.CreateItem(0)
email.To = email_login
hoje = datetime.now().strftime("%d/%m/%Y")
email.Subject = f"Relatório Geral de Tarefas - {hoje}"
email.HTMLBody = f""" 
<p style="font-family: Arial, sans-serif; font-size: 15px;">Olá Aluno! Aqui está o <strong>resumo de tarefas encontradas</strong>:</p>
{corpo_email}
<p style="font-family: Arial, sans-serif; font-size: 15px;">Faça suas atividades em dia para não ficar de recuperação.</p>
<p style="font-family: Arial, sans-serif; font-size: 15px;">Bons estudos!<br><strong>Taskbot :)</strong></p>
"""
email.Send()

# Mensagem final informando que o processo foi concluído
messagebox.showinfo("Finalizado", "Relatório enviado com sucesso!")
