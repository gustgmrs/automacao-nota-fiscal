from selenium import webdriver
import os
import pandas as pd

from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select
from time import sleep

# Inicializando o serviço do Chrome
service = Service('/home/gustgmrs/Downloads/chromedriver')
driver = webdriver.Chrome(service=service)

# Pegando caminho da página
caminho = os.getcwd()
file = 'file://' + caminho + '/files/login.html'
driver.get(file)


# Declaração de funções
def create_dataframe(caminho_xlsx):
    '''
    Função para criar um dataframe com os dados do arquivo .xlsx

    Parâmetros:
        caminho_xlsx (string): Caminho do arquivo .xlsx
    Returns:
        df (DataFrame): Dataframe com os dados do arquivo .xlsx
    '''
    df = pd.read_excel(caminho_xlsx)
    return df


def login(name, passwd):
    '''
    Função para fazer login no sistema

    Parâmetros:
        name (string): Nome do usuário
        passwd (string): Senha do usuário
    '''

    # Preenchendo nome
    user = driver.find_element(By.XPATH, '/html/body/div/form/input[1]')
    user.send_keys(name)
    sleep(0.5)

    # Preenchendo senha
    password = driver.find_element(By.XPATH, '/html/body/div/form/input[2]')
    password.send_keys(passwd)
    sleep(0.5)

    # Clicando no botão
    button = driver.find_element(By.XPATH, '/html/body/div/form/button')
    button.click()


def main():
    # Logando no sistema
    login('gustavo', '123')

    # Criando dataframe
    df_notas = create_dataframe('files/NotasEmitir.xlsx')


if __name__ == '__main__':
    main()
