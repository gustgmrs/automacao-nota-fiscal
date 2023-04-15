from selenium import webdriver
import os
import pandas as pd

from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select

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

    # Preenchendo senha
    password = driver.find_element(By.XPATH, '/html/body/div/form/input[2]')
    password.send_keys(passwd)

    # Clicando no botão
    button = driver.find_element(By.XPATH, '/html/body/div/form/button')
    button.click()


def main():
    '''
    Função principal que iterará sobre o dataframe e preencherá os campos

    Parâmetros:
        None
    Returns:
        None
    '''

    # Logando no sistema
    login('gustavo', '123')

    # Criando dataframe
    df_notes = create_dataframe('files/NotasEmitir.xlsx')

    for line in df_notes.index:
        # Coletando e preenchendo dados
        # Preenchendo nome
        name = driver.find_element(By.NAME, 'nome')
        name.send_keys(df_notes.loc[line, 'Cliente'])

        # Preenchendo endereço
        endereco = driver.find_element(By.NAME, 'endereco')
        endereco.send_keys(df_notes.loc[line, 'Endereço'])

        # Preenchendo bairro
        bairro = driver.find_element(By.NAME, 'bairro')
        bairro.send_keys(df_notes.loc[line, 'Bairro'])

        # Preenchendo município
        municipio = driver.find_element(By.NAME, 'municipio')
        municipio.send_keys(df_notes.loc[line, 'Municipio'])

        # Preenchendo CEP
        cep = driver.find_element(By.NAME, 'cep')
        cep.send_keys(str(df_notes.loc[line, 'CEP']))

        # Preenchendo UF
        uf = Select(driver.find_element(By.NAME, 'uf'))
        uf.select_by_visible_text(df_notes.loc[line, 'UF'])

        # Preenchendo CPF/CNPJ
        cpf_cnpj = driver.find_element(By.NAME, 'cnpj')
        cpf_cnpj.send_keys(str(df_notes.loc[line, 'CPF/CNPJ']))

        # Preenchendo inscrição estadual
        descricao = driver.find_element(By.NAME, 'inscricao')
        descricao.send_keys(str(df_notes.loc[line, 'Inscricao Estadual']))

        # Preenchendo descrição
        descricao = driver.find_element(By.NAME, 'descricao')
        descricao.send_keys(df_notes.loc[line, 'Descrição'])

        # Preenchendo quantidade
        quantidade = driver.find_element(By.NAME, 'quantidade')
        quantidade.send_keys(str(df_notes.loc[line, 'Quantidade']))

        # Preenchendo valor unitário
        valor_unitario = driver.find_element(By.NAME, 'valor_unitario')
        valor_unitario.send_keys(str(df_notes.loc[line, 'Valor Unitario']))

        # Preenchendo valor total
        valor_total = driver.find_element(By.NAME, 'total')
        valor_total.send_keys(str(df_notes.loc[line, 'Valor Total']))

        # Recarregando página para limpar os campos
        driver.refresh()


if __name__ == '__main__':
    main()
