from selenium import webdriver
from selenium.webdriver.common.by import By
from time import sleep
from docx import Document
import os

driver = webdriver.Chrome()
driver.get('https://contabil-devaprender.netlify.app/')

campo_email = driver.find_element(By.XPATH,"//input[@type='email']")
sleep(1)
campo_email.click()
campo_email.send_keys('teste@gmail.com')


campo_senha = driver.find_element(By.XPATH,"//input[@type='password']")
sleep(1)
campo_senha.click()
campo_senha.send_keys('123456')
sleep(1)

campo_btn_entrar = driver.find_element(By.XPATH,"//button[@type='submit']")
sleep(1)
campo_btn_entrar.click()

sleep(2)
botoes_sistemas = driver.find_elements(By.XPATH,"//a[@class='btn btn-primary mt-auto']")
sleep(2)
botoes_sistemas[0].click()

def inserir_valores_documento(caminho):
    caminho_arquivo_word = r'C:\Users\leonardo\Desktop\projContabil\relatorios\Relatorio_Contabilidade_Delícias_Foodie.docx'
    caminho_arquivo_word = caminho
    arquivo_word = Document(caminho_arquivo_word)

    ativo_circulante = ''
    caixa_equivalentes = ''
    contas_receber =''
    estoques = ''
    ativo_nao_circulante = ''
    imobilizado = ''
    intangivel = ''
    total_ativo = ''

    for tabela in arquivo_word.tables:
        for linha in tabela.rows:
            if 'Ativo Circulante' in linha.cells[0].text.strip():
                ativo_circulante = linha.cells[1].text.strip()
                ativo_circulante += '00'
            elif 'Caixa e Equivalentes' in linha.cells[0].text.strip():
                caixa_equivalentes = linha.cells[1].text.strip()
                caixa_equivalentes += '00'
            elif 'Contas a Receber' in linha.cells[0].text.strip():
                contas_receber = linha.cells[1].text.strip()
                contas_receber += '00'
            elif 'Estoques' in linha.cells[0].text.strip():
                estoques = linha.cells[1].text.strip()
                estoques += '00'
            elif 'Ativo Não Circulante' in linha.cells[0].text.strip():
                ativo_nao_circulante = linha.cells[1].text.strip()
                ativo_nao_circulante += '00'
            elif 'Imobilizado' in linha.cells[0].text.strip():
                imobilizado = linha.cells[1].text.strip()
                imobilizado += '00'
            elif 'Intangível' in linha.cells[0].text.strip():
                intangivel = linha.cells[1].text.strip()
                intangivel += '00'
            elif 'Total do Ativo' in linha.cells[0].text.strip():
                total_ativo = linha.cells[1].text.strip() 
                total_ativo += '00'


    campo_ativo_circulante = driver.find_element(By.XPATH,"//input[@id='ativo_circulante']")
    sleep(1)
    campo_ativo_circulante.click()
    campo_ativo_circulante.send_keys(ativo_circulante)

    # Preenchendo o campo "Caixa e Equivalentes de Caixa"
    campo_caixa_equivalentes = driver.find_element(By.XPATH, "//input[@id='caixa_equivalentes']")
    sleep(1)
    campo_caixa_equivalentes.click()
    campo_caixa_equivalentes.send_keys(caixa_equivalentes)

    # Preenchendo o campo "Contas a Receber"
    campo_contas_receber = driver.find_element(By.XPATH, "//input[@id='contas_receber']")
    sleep(1)
    campo_contas_receber.click()
    campo_contas_receber.send_keys(contas_receber)

    # Preenchendo o campo "Estoques"
    campo_estoques = driver.find_element(By.XPATH, "//input[@id='estoques']")
    sleep(1)
    campo_estoques.click()
    campo_estoques.send_keys(estoques)

    # Preenchendo o campo "Ativo Não Circulante"
    campo_ativo_nao_circulante = driver.find_element(By.XPATH, "//input[@id='ativo_nao_circulante']")
    sleep(1)
    campo_ativo_nao_circulante.click()
    campo_ativo_nao_circulante.send_keys(ativo_nao_circulante)

    # Preenchendo o campo "Imobilizado"
    campo_imobilizado = driver.find_element(By.XPATH, "//input[@id='imobilizado']")
    sleep(1)
    campo_imobilizado.click()
    campo_imobilizado.send_keys(imobilizado)

    # Preenchendo o campo "Intangível"
    campo_intangivel = driver.find_element(By.XPATH, "//input[@id='intangivel']")
    sleep(1)
    campo_intangivel.click()
    campo_intangivel.send_keys(intangivel)

    # Preenchendo o campo "Total do Ativo"
    campo_total_ativo = driver.find_element(By.XPATH, "//input[@id='total_ativo']")
    sleep(1)
    campo_total_ativo.click()
    campo_total_ativo.send_keys(total_ativo)

    button_cadastrar = driver.find_element(By.XPATH,"//button[@class='btn btn-primary']")
    sleep(1)
    button_cadastrar.click()

pasta_relatorios = r'C:\Users\leonardo\Desktop\projContabil\relatorios'
for nome_arquivo in os.listdir(pasta_relatorios):
    if nome_arquivo.endswith('.docx'):
        caminho_arquivo = os.path.join(pasta_relatorios,nome_arquivo)
        inserir_valores_documento(caminho_arquivo)

input("Aperte enter para fechar")