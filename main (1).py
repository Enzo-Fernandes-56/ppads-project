# Abre o chrome
import win32com.client as win32
import time
import requests
from bs4 import BeautifulSoup
from selenium.webdriver.common.by import By
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from datetime import date
from datetime import datetime


servico = Service(ChromeDriverManager().install())

navegador = webdriver.Chrome(service=servico)
navegador.maximize_window()


def existe_elemento(elemento):
    if len(elemento) == 0:
        return False
    else:
        return True


def num_mes(x):
    months = {
        'jan': 1,
        'fev': 2,
        'mar': 3,
        'abr': 4,
        'mai': 5,
        'jun': 6,
        'jul': 7,
        'ago': 8,
        'set': 9,
        'out': 10,
        'nov': 11,
        'dez': 12
    }
    a = x.strip()[:3].lower()
    ez = months[a]
    return ez


data_atual = date.today()

# Login e Senha do aluno
# Login
navegador.find_element(
    "xpath", '//*[@id="i0116"]').send_keys("10415482@mackenzista.com.br")
navegador.find_element("xpath", '//*[@id="idSIButton9"]').click()

# Senha
navegador.find_element("xpath", '//*[@id="i0118"]').send_keys("Nutela22+")
time.sleep(2)
navegador.find_element("xpath", '//*[@id="idSIButton9"]').click()
time.sleep(2)
# Botão de continuar conectado
navegador.find_element("xpath", '//*[@id="idSIButton9"]').click()
time.sleep(2)
# Botão do Calendário
navegador.find_element(
    "xpath", '//*[@id="global_nav_calendar_link"]/div[1]').click()
time.sleep(3)

# Le Calendario
lista_tarefa = []
tarefas = navegador.find_elements("class name", "fc-end")
# Quantidade de Eventos
contador = 0

# lista de materias
lista_materias = []

# lista do nome da atividade
lista_nome_atividades = []
# Quantidade Tarefas Concluídas
qtd_concluidas = 0
# lista data atividades
lista_data = []
# try:
for i in tarefas:
    i.click()
    elements = navegador.find_elements("class name", 'view_event_link')
    verifica1 = existe_elemento(elements)
    elements = navegador.find_elements(
        "xpath", '//*[@id="event-details-trap-focus"]/div[2]/div[1]/time')
    verifica_data1 = existe_elemento(elements)
    # elements = navegador.find_elements("xpath", '//*[@id="calendar-app"]/div[4]/div/div/table/tbody/tr/td/div/div/div[3]/div[2]/table/tbody/tr[2]/td[3]/a/div/span[3]')
    # verifica_tarefas_concluidas = existe_elemento(elements)
    # if verifica_tarefas_concluidas:
    #     print("Achamos uma tarefa concluída")
    #     qtd_concluidas += 1
    if verifica1:
        texto = navegador.find_element(
            "xpath", ' //*[@id="event-details-trap-focus"]/div[2]/table/tbody/tr[1]/td/a').text
        texto1 = texto.split("-")
        print(texto1[0])
        lista_materias.append(texto1[0])
        print(navegador.find_element("class name", 'view_event_link').text)
        nom_atividade = navegador.find_element(
            "class name", 'view_event_link').text
        lista_nome_atividades.append(nom_atividade)
        if verifica_data1:
            print(navegador.find_element(
                "xpath", '//*[@id="event-details-trap-focus"]/div[2]/div[1]/time').text)
            data = navegador.find_element(
                "xpath", '//*[@id="event-details-trap-focus"]/div[2]/div[1]/time').text
            lista_data.append(data)
        else:
            print(navegador.find_element(
                "xpath", '//*[@id="event-details-trap-focus"]/div[2]/div[1]/span/time[1]').text)
            data = navegador.find_element(
                "xpath", '//*[@id="event-details-trap-focus"]/div[2]/div[1]/span/time[1]').text
            lista_data.append(data)
    elements = navegador.find_elements(
        By.CSS_SELECTOR, '.details_title .title')
    verifica2 = existe_elemento(elements)
    if verifica2:
        print(navegador.find_element("class name", 'view_event_link').text)
    navegador.find_element(
        "xpath", '//*[@id="event-details-trap-focus"]/a').click()
    contador += 1

    # Formatação Data
num = 0
lista_data2 = []
for i in lista_data:
    if "," in i:
        sem_virgula = i.split(",")
        data_separada = sem_virgula[0].split(" ")
        lista_data[num] = data_separada[0] + " " + data_separada[1]
        data_completa = str(data_atual.year) + "-" + \
            str(num_mes(data_separada[1])) + "-" + str(data_separada[0])
        data_final = datetime.strptime(data_completa, "%Y-%m-%d").date()
        lista_data2.append(data_final)
    else:
        data_separada = i.split(" ")
        # print(data_separada[0] + " " + data_separada[1])
        lista_data[num] = data_separada[0] + " " + data_separada[1]
        data_completa = str(data_atual.year) + "-" + \
            str(num_mes(data_separada[1])) + "-" + str(data_separada[0])
        data_final = datetime.strptime(data_completa, "%Y-%m-%d").date()
        lista_data2.append(data_final)
    num += 1

    contador = 0
corpo_email = ""  # Variável que vai guardar o texto do e-mail

for i in lista_data2:
    if i >= data_atual:
        corpo_email += f"Matéria: {lista_materias[contador]}\n"
        corpo_email += f"Atividade: {lista_nome_atividades[contador]}\n"
        corpo_email += f"Data: {lista_data2[contador]}\n"
        corpo_email += "-" * 30 + "\n"  # linha separadora
    else:
        corpo_email += f"A tarefa da matéria {lista_materias[contador]} passou do prazo.\n"
        corpo_email += "-" * 30 + "\n"
    contador += 1

    import win32com


outlook = win32com.client.Dispatch("Outlook.Application")


email = outlook.CreateItem(0)

materia = "Tópicos de Engenharia de Software"
data = "20-04"

# Email info
email.To = "<enzo56goveia@gmail.com>"
email.Subject = " Lembrete de Entrega da Atividade de {materia} ".format(
    materia=materia)
email.HTMLBody = f""" 
<p>Olá Aluno!  Não esqueça que o prazo para envio da tarefa é até {data}.</p>
{corpo_email}
<p>Faça suas ativiades em dia para não ficar de recuperação</p>

<p>Bons estudos!</p>
<p>Taskbot :)</p>
"""


email.Send()
print(email.To)
