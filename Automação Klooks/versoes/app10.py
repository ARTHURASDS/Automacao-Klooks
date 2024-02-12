from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import *
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.support import expected_conditions as condicao_esperada 
import pandas as pd
from time import sleep
import random

# Variáveis dinâmicas
from variaveis_dinamicas import *

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
    
    wait = WebDriverWait(
        driver,
        10,
        poll_frequency=1,
        ignored_exceptions=[
            NoSuchElementException,
            ElementNotVisibleException,
            ElementNotSelectableException,
        ]
    )
    return driver, wait

driver, wait = iniciar_driver()

# Abrir o Klooks e fazer login
driver.get('https://app.klooks.com.br/login')
sleep(random.uniform(3, 7))

campo_email = wait.until(condicao_esperada.visibility_of_element_located(
    (By.XPATH, '//*[@id="username"]'))) # esse aq é melhor pq espera todos elementos ficarem visiveis.
campo_email.send_keys(login)
sleep(random.uniform(3, 7))

campo_senha = wait.until(condicao_esperada.visibility_of_element_located(
    (By.XPATH, '//*[@id="password"]')))
#campo_senha = driver.find_element(By.XPATH, '//*[@id="password"]')
campo_senha.send_keys(senha)
sleep(random.uniform(3, 7))

botao_login = wait.until(condicao_esperada.visibility_of_element_located(
    (By.XPATH, '//*[text()="Login"]')))
#botao_login = driver.find_element(By.XPATH, '//*[text()="Login"]')
sleep(random.uniform(3, 7))
botao_login.click()

# Clicar na Busca Avançada e definir critérios de pesquisa
busca_avancada = wait.until(condicao_esperada.visibility_of_element_located(
    (By.XPATH, '//*[@id="searchForm"]/div[2]/a')))
#busca_avancada = driver.find_element(By.XPATH, '//*[@id="searchForm"]/div[2]/a')
busca_avancada.click()
sleep(random.uniform(3, 7))

opcao_estado = wait.until(condicao_esperada.visibility_of_element_located(
    (By.XPATH, '//*[@id="advanced-search"]/div[3]/div[2]/div[1]/div/span[2]/div/button')))
#opcao_estado = driver.find_element(By.XPATH, '//*[@id="advanced-search"]/div[3]/div[2]/div[1]/div/span[2]/div/button')
opcao_estado.click()
sleep(random.uniform(3, 7))

for estado, xpath_numero in estados_dic.items():

    label_pesquisar = wait.until(condicao_esperada.visibility_of_element_located(
        (By.XPATH, '//*[@id="advanced-search"]/div[3]/div[2]/div[1]/div/span[2]/div/ul/li[1]/div/input')))
    #label_pesquisar = driver.find_element(By.XPATH, '//*[@id="advanced-search"]/div[3]/div[2]/div[1]/div/span[2]/div/ul/li[1]/div/input')
    label_pesquisar.click()
    label_pesquisar.send_keys(estado)
    sleep(random.uniform(3, 7))

    caixa_estado = wait.until(condicao_esperada.visibility_of_element_located(
        (By.XPATH, f'//*[@id="advanced-search"]/div[3]/div[2]/div[1]/div/span[2]/div/ul/li[{xpath_numero}]/a/label')))
    #caixa_estado = driver.find_element(By.XPATH, '//*[@id="advanced-search"]/div[3]/div[2]/div[1]/div/span[2]/div/ul/li[17]/a/label/input')
    caixa_estado.click()
    sleep(random.uniform(3, 7))

    caixa_fechar = wait.until(condicao_esperada.visibility_of_element_located(
        (By.XPATH, '//*[@id="advanced-search"]/div[3]/div[2]/div[1]/div/span[2]/div/ul/li[1]/div/span[2]/button/i')))
    label_pesquisar.clear()

placeholder_caixa_inicial = wait.until(condicao_esperada.visibility_of_element_located(
    (By.XPATH, '//*[@id="caixaEquivalentesMin"]')))
#placeholder_caixa_inicial = driver.find_element(By.XPATH, '//*[@id="caixaEquivalentesMin"]')
driver.execute_script("arguments[0].scrollIntoView(true);", placeholder_caixa_inicial)
placeholder_caixa_inicial.click()
placeholder_caixa_inicial.send_keys(caixa_valor_inicial)
sleep(random.uniform(3, 7))

placeholder_caixa_final = wait.until(condicao_esperada.visibility_of_element_located(
    (By.XPATH, '//*[@id="caixaEquivalentesMax"]')))
#placeholder_caixa_final = driver.find_element(By.XPATH, '//*[@id="caixaEquivalentesMax"]')
placeholder_caixa_final.click()
placeholder_caixa_final.send_keys(caixa_valor_final)
sleep(random.uniform(3, 7))

placeholder_ordenar = wait.until(condicao_esperada.visibility_of_element_located(
    (By.ID, 'order')))
#placeholder_ordenar = driver.find_element(By.ID, 'order')
driver.execute_script("arguments[0].scrollIntoView(true);", placeholder_ordenar)
sleep(random.uniform(3, 7))

select = Select(placeholder_ordenar)
select.select_by_visible_text('Caixa')
sleep(random.uniform(3, 7))

botao_pesquisar = wait.until(condicao_esperada.visibility_of_element_located(
    (By.XPATH, '//*[@id="advanced-search"]/div[5]/div[2]/button')))
#botao_pesquisar = driver.find_element(By.XPATH, '//*[@id="advanced-search"]/div[5]/div[2]/button')
sleep(random.uniform(3, 7))
botao_pesquisar.click()

# Coletar informações das empresas
dados_empresas = []

while True:
    # divs_empresas = wait.until(condicao_esperada.visibility_of_element_located(
    #(By.XPATH, '//div[@class="col-xs-11"]')))
    divs_empresas = driver.find_elements(By.XPATH, '//div[@class="col-xs-11"]')
    if not divs_empresas:
        break
    
    for div_empresa in divs_empresas:
        # Captura de informações das empresas...
        try:
            link_empresa = div_empresa.find_element(By.XPATH, './a').get_attribute('href')
            driver.execute_script("window.open(arguments[0])", link_empresa)
            sleep(random.uniform(3, 7))

            driver.switch_to.window(driver.window_handles[-1])
            sleep(random.uniform(3, 7))

            botao_mais_informacoes = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, '//*[@id="moreInfo"]'))
            )
            botao_mais_informacoes.click()
            sleep(random.uniform(3, 7))

            # Verifica a situação cadastral
            situacao_cadastral_element = wait.until(condicao_esperada.visibility_of_element_located(
    (By.XPATH, '//*[@id="business-details"]/p[5]')))
            #situacao_cadastral_element = driver.find_element(By.XPATH, '//*[@id="business-details"]/p[5]')
            situacao_cadastral = situacao_cadastral_element.text
            
            if 'Situação Cadastral: ATIVA' == situacao_cadastral:
                # Captura dos dados da empresa...
                nome_cnpj_element = wait.until(condicao_esperada.visibility_of_element_located(
    (By.XPATH, '/html/body/div[1]/div[2]/div[2]/div/div[2]/div[2]/div/div[1]/div[1]/div/div[1]/h1')))
                #nome_cnpj_element = driver.find_element(By.XPATH, '/html/body/div[1]/div[2]/div[2]/div/div[2]/div[2]/div/div[1]/div[1]/div/div[1]/h1')
                nome_empresa = nome_cnpj_element.text.split('\n')[0]
                cnpj = nome_cnpj_element.text.split('\n')[1].split(': ')[1]

                telefone_element = wait.until(condicao_esperada.visibility_of_element_located(
    (By.XPATH, '//*[@id="business-details"]/p[6]')))
                #telefone_element = driver.find_element(By.XPATH, '//*[@id="business-details"]/p[6]')
                telefone_completo = telefone_element.text
                indice_dois_pontos = telefone_completo.index(':') + 2  # Encontra o índice após os dois pontos
                telefone = telefone_completo[indice_dois_pontos:]  # Pega o texto após os dois pontos

                email_element = wait.until(condicao_esperada.visibility_of_element_located(
    (By.XPATH, '//*[@id="business-details"]/p[7]')))
                #email_element = driver.find_element(By.XPATH, '//*[@id="business-details"]/p[7]')
                email_completo = email_element.text
                if "@" in email_completo:  # Verifica se há "@" no texto
                    indice_dois_pontos = email_completo.index(':') + 2
                    email = email_completo[indice_dois_pontos:]
                else:
                    email = "vazio"  # Define o email como vazio se não houver "@"


                cnae_element = wait.until(condicao_esperada.visibility_of_element_located(
    (By.XPATH, '//*[@id="primaryActivity"]')))
                #cnae_element = driver.find_element(By.XPATH, '//*[@id="primaryActivity"]')
                cnae_completo = cnae_element.text
                indice_hifen = cnae_completo.index('-') + 2  # Encontra o índice após o hífen
                cnae_principal = cnae_completo[indice_hifen:]

                distrito_element = wait.until(condicao_esperada.visibility_of_element_located(
    (By.XPATH, '/html/body/div[1]/div[2]/div[2]/div/div[2]/div[2]/div/div[1]/div[1]/div/h2/small[1]')))
                #distrito_element = driver.find_element(By.XPATH, '/html/body/div[1]/div[2]/div[2]/div/div[2]/div[2]/div/div[1]/div[1]/div/h2/small[1]')
                distrito_completo = distrito_element.text
                indice_dois_pontos = distrito_completo.index(':') + 2
                distrito = distrito_completo[indice_dois_pontos:]

                dono = wait.until(condicao_esperada.visibility_of_element_located(
    (By.XPATH, '/html/body/div[1]/div[2]/div[2]/div/div[2]/div[2]/div/div[2]/div/div[2]/p[1]/strong')))
                #dono = driver.find_element(By.XPATH, '/html/body/div[1]/div[2]/div[2]/div/div[2]/div[2]/div/div[2]/div/div[2]/p[1]/strong').text

            # Captura de outros dados...

            dados_empresas.append({
                'Nome da empresa': nome_empresa,
                'CNPJ': cnpj,
                'Telefone': telefone,
                'cnae_principal': cnae_principal,
                'dono': dono,
                'email': email,
                'distrito': distrito,
                # Outros campos...
            })
        
            driver.close()
            sleep(random.uniform(3, 7))
            driver.switch_to.window(driver.window_handles[0])

        except NoSuchElementException as e:
            print(f"Erro ao encontrar elemento: {e}")
            continue

    try:
        botao_proximo = driver.find_element(By.XPATH, '//*[@id="search"]/div[1]/div[2]/div[42]/a[3]')
        botao_proximo.click()
        sleep(random.uniform(3, 7))

    except NoSuchElementException:
        break  # Se não encontrar o botão, encerra a coleta de dados

# Criação do DataFrame e salvamento no Excel
df = pd.DataFrame(dados_empresas)
file_path = local_download_excel
df.to_excel(file_path, index=False)
