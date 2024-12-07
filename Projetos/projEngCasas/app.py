from selenium import webdriver
from selenium.webdriver.common.by import By
from time import sleep
from datetime import datetime
import openpyxl

# Configuração do driver
driver = webdriver.Chrome()
driver.get('https://www.imoveismartinelli.com.br/pesquisa-de-imoveis/?locacao_venda=V&id_cidade%5B%5D=21&finalidade=&dormitorio=&garagem=&vmi=&vma=&ordem=4')

sleep(3)

# Capturando elementos
precos = driver.find_elements(By.XPATH, "//div[@class='card-valores']/div")
links = driver.find_elements(By.XPATH, "//a[@class='carousel-cell is-selected']")

# Validando os elementos encontrados
print(f"Preços encontrados: {len(precos)}")
print(f"Links encontrados: {len(links)}")

if not precos or not links:
    print("Nenhum dado encontrado. Verifique os XPaths.")
else:
    # Abrindo a planilha
    workbok = openpyxl.load_workbook('C:/Users/leonardo/Desktop/Projetos/projEngCasas/imoveis.xlsx')
    print(workbok.sheetnames)  # Certifique-se de que "precos" está aqui
    pagina_imoveis = workbok['precos']

    # Adicionando os dados
    for preco, link in zip(precos, links):
        preco_formatado = preco.text.split(' ')[1] if ' ' in preco.text else preco.text
        link_pronto = link.get_attribute('href')
        data_atual = datetime.now().strftime('%d/%m/%Y')
        pagina_imoveis.append([preco_formatado, link_pronto, data_atual])
        print(f"Adicionado: {preco_formatado}, {link_pronto}, {data_atual}")

    # Salvando a planilha
    workbok.save('C:/Users/leonardo/Desktop/Projetos/projEngCasas/imoveis.xlsx')
    print("Planilha salva com sucesso!")

# Fechando o navegador
driver.quit()
