from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import NoSuchElementException
import pandas as pd
from time import sleep
import random

# Variáveis dinâmicas
from variaveis_dinamicas import estado, login, senha, caixa_valor_inicial, caixa_valor_final, local_download_excel

# Função para iniciar o driver
def iniciar_driver():
    chrome_options = Options()
    arguments = ['--lang=pt-BR', '--window-size=1920,1080', '--incognito']
    for argument in arguments:
        chrome_options.add_argument(argument)

    chrome_options.add_experimental_option('prefs', {
        'download.prompt_for_download': False,
        'profile.default_content_setting_values.notifications': 2,
        'profile.default_content_setting_values.automatic_downloads': 1,
    })
    driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=chrome_options)
    return driver

driver = iniciar_driver()

# Abrir o Klooks e fazer login
driver.get('https://app.klooks.com.br/login')
sleep(random.uniform(2, 4))

campo_email = driver.find_element(By.XPATH, '//*[@id="username"]')
campo_email.send_keys(login)
sleep(random.uniform(2, 4))

campo_senha = driver.find_element(By.XPATH, '//*[@id="password"]')
campo_senha.send_keys(senha)
sleep(random.uniform(2, 4))

botao_login = driver.find_element(By.XPATH, '//*[text()="Login"]')
sleep(random.uniform(2, 4))
botao_login.click()

# Clicar na Busca Avançada e definir critérios de pesquisa
busca_avancada = driver.find_element(By.XPATH, '//*[@id="searchForm"]/div[2]/a')
busca_avancada.click()
sleep(random.uniform(2, 4))

opcao_estado = driver.find_element(By.XPATH, '//*[@id="advanced-search"]/div[3]/div[2]/div[1]/div/span[2]/div/button')
opcao_estado.click()
sleep(random.uniform(2, 4))

label_pesquisar = driver.find_element(By.XPATH, '//*[@id="advanced-search"]/div[3]/div[2]/div[1]/div/span[2]/div/ul/li[1]/div/input')
label_pesquisar.click()
label_pesquisar.send_keys(estado)
sleep(random.uniform(2, 4))

caixa_estado = driver.find_element(By.XPATH, '//*[@id="advanced-search"]/div[3]/div[2]/div[1]/div/span[2]/div/ul/li[17]/a/label/input')
caixa_estado.click()
sleep(random.uniform(2, 4))

placeholder_caixa_inicial = driver.find_element(By.XPATH, '//*[@id="caixaEquivalentesMin"]')
driver.execute_script("arguments[0].scrollIntoView(true);", placeholder_caixa_inicial)
placeholder_caixa_inicial.click()
placeholder_caixa_inicial.send_keys(caixa_valor_inicial)
sleep(random.uniform(2, 4))

placeholder_caixa_final = driver.find_element(By.XPATH, '//*[@id="caixaEquivalentesMax"]')
placeholder_caixa_final.click()
placeholder_caixa_final.send_keys(caixa_valor_final)
sleep(random.uniform(2, 4))

placeholder_ordenar = driver.find_element(By.ID, 'order')
driver.execute_script("arguments[0].scrollIntoView(true);", placeholder_ordenar)
sleep(random.uniform(2, 4))

select = Select(placeholder_ordenar)
select.select_by_visible_text('Caixa')
sleep(random.uniform(2, 4))

botao_pesquisar = driver.find_element(By.XPATH, '//*[@id="advanced-search"]/div[5]/div[2]/button')
sleep(random.uniform(2, 4))
botao_pesquisar.click()

# Coletar informações das empresas
dados_empresas = []
divs_empresas = driver.find_elements(By.XPATH, '//div[@class="col-xs-11"]')
sleep(random.uniform(2, 4))


while True:
    divs_empresas = driver.find_elements(By.XPATH, '//div[@class="col-xs-11"]')
    sleep(random.uniform(2, 4))
    for div_empresa in divs_empresas:
        link_empresa = div_empresa.find_element(By.XPATH, './a').get_attribute('href')
        driver.execute_script("window.open(arguments[0])", link_empresa)
        sleep(random.uniform(2, 4))

        driver.switch_to.window(driver.window_handles[-1])
        sleep(random.uniform(2, 4))
        
        try:
            botao_mais_informacoes = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, '//*[@id="moreInfo"]'))
            )
            botao_mais_informacoes.click()
            sleep(random.uniform(2, 4))

            nome_cnpj_element = driver.find_element(By.XPATH, '/html/body/div[1]/div[2]/div[2]/div/div[2]/div[2]/div/div[1]/div[1]/div/div[1]/h1')
            nome_empresa = nome_cnpj_element.text.split('\n')[0]
            cnpj = nome_cnpj_element.text.split('\n')[1].split(': ')[1]
            
            telefone_element = driver.find_element(By.XPATH, '//*[@id="business-details"]/p[6]')
            telefone_completo = telefone_element.text
            indice_dois_pontos = telefone_completo.index(':') + 2  # Encontra o índice após os dois pontos
            telefone = telefone_completo[indice_dois_pontos:]  # Pega o texto após os dois pontos

            email_element = driver.find_element(By.XPATH, '//*[@id="business-details"]/p[7]')
            email_completo = email_element.text
            if "@" in email_completo:  # Verifica se há "@" no texto
                indice_dois_pontos = email_completo.index(':') + 2
                email = email_completo[indice_dois_pontos:]
            else:
                email = "vazio"  # Define o email como vazio se não houver "@"


            cnae_element = driver.find_element(By.XPATH, '//*[@id="primaryActivity"]')
            cnae_completo = cnae_element.text
            indice_hifen = cnae_completo.index('-') + 2  # Encontra o índice após o hífen
            cnae_principal = cnae_completo[indice_hifen:]

            distrito_element = driver.find_element(By.XPATH, '/html/body/div[1]/div[2]/div[2]/div/div[2]/div[2]/div/div[1]/div[1]/div/h2/small[1]')
            distrito_completo = distrito_element.text
            indice_dois_pontos = distrito_completo.index(':') + 2
            distrito = distrito_completo[indice_dois_pontos:]

            dados_empresas.append({
            'Nome da empresa': nome_empresa,
            'CNPJ': cnpj,
            'Telefone': telefone,
            'cnae_principal': cnae_principal,
            'dono': dono,
            'email': email,
            'distrito': distrito,
            # Adicione aqui os outros dados que você quer coletar
        })
            
            driver.close()
            sleep(random.uniform(4, 6))
            driver.switch_to.window(driver.window_handles[0])

        except NoSuchElementException as e:
            dono = "vazio"  
            telefone = "vazio"
            email = "vazio"
            cnae_principal = "vazio"
            distrito = "vazio"
            print(f"Erro ao encontrar elemento: {e}")

    try:
    botao_proximo = driver.find_element(By.XPATH, xpath_botao_proximo)
    botao_proximo.click()
    sleep(random.uniform(2, 4))

    except NoSuchElementException:
    break  # Se não encontrar o botão, encerra a coleta de dados



df = pd.DataFrame(dados_empresas)
file_path = local_download_excel
df.to_excel(file_path, index=False)
