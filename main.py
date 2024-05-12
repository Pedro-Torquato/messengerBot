import pandas as pd
import datetime
from selenium import webdriver
import urllib
import time
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By


def main():
    navegador = webdriver.Chrome()
    navegador.get('https://web.whatsapp.com/')
    while len(navegador.find_elements(By.ID, 'pane-side')) < 1:
        time.sleep(1)

    clientes_df = clientes_amanha()
    for i, nome in enumerate(clientes_df['Nome']):
        telefone = clientes_df.loc[i, "Telefone"]
        date = clientes_df.loc[i, "Data"]
        time_hour = clientes_df.loc[i, "Hora"]
        message = create_message(nome, date, time_hour)
        link = create_link(telefone, message)
        send_message(link, navegador)


def clientes_amanha():
    agenda_df = pd.read_excel("agenda.xlsx")
    clientes = agenda_df.loc[agenda_df['Data'] == str(datetime.date.today() + datetime.timedelta(days=1))]
    clientes.reset_index(drop=True, inplace=True)
    return clientes


def create_message(nome, data, hora):
    data = data.strftime('%d/%m/%Y')
    hora = hora.strftime('%H:%M')
    with open('mensagem.txt', 'r') as document:
        mensagem = document.read().replace('(cliente)', nome).replace('(data)', data).replace(
            '(horario)', str(hora))
    return mensagem


def create_link(telefone, message):
    message = urllib.parse.quote(message)
    link = f"https://wa.me/55{telefone}?text={message}"
    return link


def send_message(link, navegador):
    navegador.get(link)
    while len(navegador.find_elements(
            By.XPATH, '//*[@id="action-button"]')) < 1:
        time.sleep(1)
    navegador.find_element(By.XPATH, '//*[@id="action-button"]').send_keys(Keys.ENTER)
    time.sleep(3)
    navegador.find_element(By.XPATH, '//*[@id="fallback_block"]/div/div/h4[2]/a').send_keys(Keys.ENTER)
    while len(navegador.find_elements(
            By.XPATH, '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[2]/button')) < 1:
        time.sleep(1)
    navegador.find_element(
        By.XPATH, '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[2]/button').send_keys(Keys.ENTER)
    time.sleep(5)


main()
