from selenium import webdriver
from selenium.webdriver.chrome.service import Service
import pandas as pd
from selenium.webdriver.common.keys import Keys
import time
from selenium.webdriver.common.by import By

# Atribuir variaveis
url = "https://rpachallenge.com/?lang=EN"
arquivo = "C:/Users/iago.barbosa/Desktop/Python/Arquivo/challenge.xlsx"
# Lendo o arquivo Excel com Pandas
dados = pd.read_excel(arquivo)
print(dados)

# Iniciando automação
print('Iniciando Automação\n')
print("Acessando o site RPA Challenge")

browser= webdriver.Chrome()
browser.get(url)
# Esperando o site carregar
time.sleep(2)

# Maximizar janela
browser.maximize_window()

# Dando Start no Site para iniciar o preenchimento dos dados nos campos.#
botao_Start = browser.find_element(By.XPATH, '/html/body/app-root/div[2]/app-rpa1/div/div[1]/div[6]/button')
botao_Start.click()

# Iterar sobre a tabela
for i, row in dados.iterrows():
    primeiroNome = row["First Name"]
    sobrenome = row["Last Name"]
    nomeEmpresa = row["Company Name"]
    profissao = row["Role in Company"]
    endereco = row["Address"]
    email = row["Email"]
    telefone = row["Phone Number"]

    # Preencher campos com os dados da pagina usando o selenium
    # Primeiro nome
    browser.find_element(By.CSS_SELECTOR, '[ng-reflect-name="labelFirstName"]').send_keys(primeiroNome)
    # Sobrenome
    browser.find_element(By.CSS_SELECTOR, '[ng-reflect-name="labelLastName"]').send_keys(sobrenome)
    # Numero de telefone
    browser.find_element(By.CSS_SELECTOR, '[ng-reflect-name="labelPhone"]').send_keys(str(telefone))
    # E-mail
    browser.find_element(By.CSS_SELECTOR, '[ng-reflect-name="labelEmail"]').send_keys(email)
    # Endereço
    browser.find_element(By.CSS_SELECTOR, '[ng-reflect-name="labelAddress"]').send_keys(endereco)
    # Nome da empresa
    browser.find_element(By.CSS_SELECTOR, '[ng-reflect-name="labelCompanyName"]').send_keys(nomeEmpresa)
    # Profissão
    browser.find_element(By.CSS_SELECTOR, '[ng-reflect-name="labelRole"]').send_keys(profissao)
    # Botaão para proxima linha
    browser.find_element(By.XPATH, '/html/body/app-root/div[2]/app-rpa1/div/div[2]/form/input').click()

print('Fim da automação')