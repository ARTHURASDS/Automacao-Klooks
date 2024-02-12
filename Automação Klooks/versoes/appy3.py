from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from time import sleep
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import NoSuchElementException
import pandas as pd

# Variáveis dinâmicas
from variaveis_dinamicas import estado, login, senha, caixa_valor_inicial, caixa_valor_final

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

# Abrir o Klooks
driver.get('https://app.klooks.com.br/login')
sleep(2)

# Preencher credenciais e fazer login
campo_email = driver.find_element(By.XPATH, '//*[@id="username"]')
campo_email.send_keys(login)
sleep(2)

campo_senha = driver.find_element(By.XPATH, '//*[@id="password"]')
campo_senha.send_keys(senha)
sleep(2)

botao_login = driver.find_element(By.XPATH, '//*[text()="Login"]')
sleep(2)
botao_login.click()

# Clicar na Busca Avançada e definir critérios de pesquisa
busca_avancada = driver.find_element(By.XPATH, '//*[@id="searchForm"]/div[2]/a')
busca_avancada.click()
sleep(2)

opcao_estado = driver.find_element(By.XPATH, '//*[@id="advanced-search"]/div[3]/div[2]/div[1]/div/span[2]/div/button')
opcao_estado.click()
sleep(2)

label_pesquisar = driver.find_element(By.XPATH, '//*[@id="advanced-search"]/div[3]/div[2]/div[1]/div/span[2]/div/ul/li[1]/div/input')
label_pesquisar.click()
label_pesquisar.send_keys(estado)
sleep(2)

caixa_estado = driver.find_element(By.XPATH, '//*[@id="advanced-search"]/div[3]/div[2]/div[1]/div/span[2]/div/ul/li[17]/a/label/input')
caixa_estado.click()
sleep(2)

placeholder_caixa_inicial = driver.find_element(By.XPATH, '//*[@id="caixaEquivalentesMin"]')
driver.execute_script("arguments[0].scrollIntoView(true);", placeholder_caixa_inicial)
placeholder_caixa_inicial.click()
placeholder_caixa_inicial.send_keys(caixa_valor_inicial)
sleep(2)

placeholder_caixa_final = driver.find_element(By.XPATH, '//*[@id="caixaEquivalentesMax"]')
placeholder_caixa_final.click()
placeholder_caixa_final.send_keys(caixa_valor_final)
sleep(2)

placeholder_ordenar = driver.find_element(By.ID, 'order')
driver.execute_script("arguments[0].scrollIntoView(true);", placeholder_ordenar)
sleep(2)

select = Select(placeholder_ordenar)
select.select_by_visible_text('Caixa')
sleep(2)

botao_pesquisar = driver.find_element(By.XPATH, '//*[@id="advanced-search"]/div[5]/div[2]/button')
sleep(2)
botao_pesquisar.click()

# Coletar informações das empresas
dados_empresas = []

# Coletar os elementos das empresas
divs_empresas = driver.find_elements(By.XPATH, '//div[@class="col-xs-11"]')
sleep(2)

# Loop para coletar informações de cada empresa
for div_empresa in divs_empresas:
    link_empresa = div_empresa.find_element(By.XPATH, './a').get_attribute('href')
    driver.execute_script("window.open(arguments[0])", link_empresa)
    sleep(2)

    driver.switch_to.window(driver.window_handles[-1])
    sleep(2)
    
    botao_mais_informacoes = driver.find_element(By.XPATH, '//*[@id="moreInfo"]')
    botao_mais_informacoes.click()
    sleep(2)
    
    nome_cnpj_element = driver.find_element(By.XPATH, '/html/body/div[1]/div[2]/div[2]/div/div[2]/div[2]/div/div[1]/div[1]/div/div[1]/h1')
    nome_empresa = nome_cnpj_element.text.split('\n')[0]
    cnpj = nome_cnpj_element.text.split('\n')[1].split(': ')[1]
    
    # Preenche com "vazio" se o XPath não for encontrado
    try:
        telefone = driver.find_element(By.XPATH, '//*[@id="business-details"]/p[6]').text
        email = driver.find_element(By.XPATH, '//*[@id="business-details"]/p[7]').text
        cnae_principal = driver.find_element(By.XPATH, '//*[@id="primaryActivity"]').text
        distrito = driver.find_element(By.XPATH, '/html/body/div[1]/div[2]/div[2]/div/div[2]/div[2]/div/div[1]/div[1]/div/h2/small[2]').text.split('\n')[1]
        dono = driver.find_element(By.XPATH, '/html/body/div[1]/div[2]/div[2]/div/div[2]/div[2]/div/div[2]/div/div[2]/p[1]/strong').text
    except NoSuchElementException:
        dono = "vazio"  
        telefone = "vazio"
        email = "vazio"
        cnae_principal = "vazio"
        distrito = "vazio"

    # Adiciona os dados a uma estrutura de dicionário
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

    # Fecha a guia/página atual e volta para a janela anterior
    driver.close()
    sleep(4)
    driver.switch_to.window(driver.window_handles[0])

# Converte os dados coletados em um DataFrame pandas
df = pd.DataFrame(dados_empresas)

# Caminho completo para o arquivo Excel
file_path = r"C:\\Users\\arthu\\Desktop\\Automação Klooks\\informacoes_empresas.xlsx"

# Salva o DataFrame em um arquivo Excel no caminho especificado
df.to_excel(file_path, index=False)
