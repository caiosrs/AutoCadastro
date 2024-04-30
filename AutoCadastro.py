from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.ui import Select  # Adicione esta linha
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import time

driver = webdriver.Chrome()

url = "http://127.0.0.1:5000/"

driver.get(url)

time.sleep(2)

# Preenche os campos do formulário
email_element = driver.find_element(By.ID, 'email')
email_element.send_keys('caio.silva@informatecservicos.com.br')
time.sleep(2)
senha_element = driver.find_element(By.ID, 'senha')
senha_element.send_keys('1234567890')
time.sleep(2)

cadastro_button = driver.find_element(By.XPATH, '//button[@type="submit"]')
cadastro_button.click()
time.sleep(2)

def preencher_formulario(
        id, nome, sexo, endereco, numero, bairro, cidade, estado, cep, telefone, estadocivil, datanascimento, cpf,
        salario, funcao, tituloeleitor, zona, secao, idcentro, nomecentro, nomepai, nomemae, observacoes,
):
    id_element = driver.find_element(By.ID, 'id')
    id_element.send_keys(id)
    time.sleep(1)
    nome_element = driver.find_element(By.ID, 'nome')
    nome_element.send_keys(nome)
    time.sleep(1)
    sexo_element = driver.find_element(By.ID, 'sexo')
    if sexo.lower() == 'masculino':
        Select(sexo_element).select_by_index(0)
    elif sexo.lower() == 'feminino':
        Select(sexo_element).select_by_index(1)
    else:
        Select(sexo_element).select_by_index(2)
    time.sleep(1)
    endereco_element = driver.find_element(By.ID, 'endereco')
    endereco_element.send_keys(endereco)
    time.sleep(1)
    numero_element = driver.find_element(By.ID, 'numero')
    numero_element.send_keys(numero)
    time.sleep(1)
    bairro_element = driver.find_element(By.ID, 'bairro')
    bairro_element.send_keys(bairro)
    time.sleep(1)
    cidade_element = driver.find_element(By.ID, 'cidade')
    cidade_element.send_keys(cidade)
    time.sleep(1)
    estado_element = driver.find_element(By.ID, 'estado')
    if estado_element.is_displayed():
        # Obtém o valor visível do dropdownlist na planilha
        estado_planilha = row['Estado']
        
        # Seleciona a opção correspondente no dropdownlistc
        Select(estado_element).select_by_visible_text(estado_planilha)
    time.sleep(1)
    cep_element = driver.find_element(By.ID, 'CEP')
    cep_element.send_keys(cep)
    time.sleep(1)
    telefone_element = driver.find_element(By.ID, 'telefone')
    telefone_element.send_keys(telefone)
    time.sleep(1)    
    estadocivil_element = driver.find_element(By.ID, 'estado-civil')
    estadocivil_element.send_keys(estadocivil)
    time.sleep(1)
    datanascimento_element = driver.find_element(By.ID, 'data-nascimento')
    datanascimento_element.send_keys(datanascimento)
    time.sleep(1)
    cpf_element = driver.find_element(By.ID, 'cpf')
    cpf_element.send_keys(cpf)
    time.sleep(1)
    salario_element = driver.find_element(By.ID, 'salario')
    salario_element.send_keys(salario)
    time.sleep(1)
    funcao_element = driver.find_element(By.ID, 'funcao')
    funcao_element.send_keys(funcao)
    time.sleep(1)
    tituloeleitor_element = driver.find_element(By.ID, 'titulo-eleitor')
    tituloeleitor_element.send_keys(tituloeleitor)
    time.sleep(1)
    zona_element = driver.find_element(By.ID, 'zona')
    zona_element.send_keys(zona)
    time.sleep(1)
    secao_element = driver.find_element(By.ID, 'secao')
    secao_element.send_keys(secao)
    time.sleep(1)
    idcentro_element = driver.find_element(By.ID, 'id-centro-resultado')
    idcentro_element.send_keys(idcentro)
    time.sleep(1)
    nomecentro_element = driver.find_element(By.ID, 'nome-centro-resultado')
    nomecentro_element.send_keys(nomecentro)
    time.sleep(1)
    nomepai_element = driver.find_element(By.ID, 'nome-pai')
    nomepai_element.send_keys(nomepai)
    time.sleep(1)
    nomemae_element = driver.find_element(By.ID, 'nome-mae')
    nomemae_element.send_keys(nomemae)
    time.sleep(1)
    observacoes_element = driver.find_element(By.ID, 'observacoes')
    observacoes_element.send_keys(observacoes)
    time.sleep(1)
    cadastro_button = driver.find_element(By.XPATH, '//button[@type="button"]')
    cadastro_button.click()
    time.sleep(3)
    webdriver.ActionChains(driver).send_keys(Keys.ESCAPE).perform()
    time.sleep(2)
    #'sucesso' pelo ID ou outro identificador melhor
    try:
        success_message = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, 'sucesso'))
        )
        print("Cadastro realizado com sucesso!")
    except:
        print("Tempo limite excedido. Falha no cadastro.")

    id_element.clear()
    nome_element.clear()
    #sexo_element.clear()
    endereco_element.clear()
    numero_element.clear()
    bairro_element.clear()
    cidade_element.clear()
    #estado_element.clear()
    cep_element.clear()
    telefone_element.clear()
    estadocivil_element.clear()
    datanascimento_element.clear()
    cpf_element.clear()
    salario_element.clear()
    funcao_element.clear()
    tituloeleitor_element.clear()
    zona_element.clear()
    secao_element.clear()
    idcentro_element.clear()
    nomecentro_element.clear()
    nomepai_element.clear()
    nomemae_element.clear()
    observacoes_element.clear()

df = pd.read_excel('Saída Geral.xlsx')

# Iterar sobre as linhas da planilha
for index, row in df.iterrows():
    id = row['ID']
    nome = row['NOME']
    sexo = row['Sexo']
    endereco = row['Endereço']
    numero = row['Numero']
    bairro = row['Bairro']
    telefone = row['Telefone']
    cep = row['CEP']
    cidade = row['Cidade']
    estado = row['Estado']
    estadocivil = row['Estado Civil']
    datanascimento = row ['Data de Nascimento']
    cpf = row['CPF']
    salario = row['Salário']
    funcao = row['Função']
    tituloeleitor = row['Título Eleitor']
    zona = row['Zona']
    secao = row['Seção']
    idcentro = row['Centro de Resultado ID']
    nomecentro = row['Centro de Resultado']
    nomepai = row['Nome do Pai']
    nomemae = row['Nome da Mãe']
    observacoes = row['Observações']

    preencher_formulario(
        id, nome, sexo, endereco, numero, bairro, cidade, estado, cep, telefone, estadocivil, datanascimento, cpf,
        salario, funcao, tituloeleitor, zona, secao, idcentro, nomecentro, nomepai, nomemae, observacoes,
)

driver.quit()