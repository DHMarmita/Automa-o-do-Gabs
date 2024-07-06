from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
import re


def buscar_informacoes_google(cliente, estado, cidade):
    # Configuração do webdriver do Chrome
    options = Options()

    # Configuração do serviço do webdriver
    service = Service(ChromeDriverManager().install())

    # Inicialização do webdriver
    driver = webdriver.Chrome(service=service, options=options)

    try:
        # Montando a query de busca no Google
        query = f"{cliente} - {estado} - {cidade} - CNPJ site:cnpj.biz"

        # URL do Google Search
        google_url = f"https://www.google.com/search?q={query}"

        # Abrindo o Google Search
        driver.get(google_url)

        # Esperando um momento para carregar os resultados
        driver.implicitly_wait(5)

        # Pegando os links dos resultados do Google
        search_results = driver.find_elements(By.CSS_SELECTOR, "div.g")

        link_encontrado = None

        for result in search_results:
            link_element = result.find_element(By.TAG_NAME, 'a')
            link = link_element.get_attribute('href')
            if 'cnpj.biz' in link:
                link_encontrado = link
                break

        # Verifica se encontrou um link válido do cnpj.biz
        if link_encontrado:
            # Indo para o link encontrado
            driver.get(link_encontrado)

            # Esperando um momento para carregar a página do cnpj.biz
            driver.implicitly_wait(5)

            # Extraindo as informações desejadas usando XPath absoluto
            informacoes = {}

            # CNPJ
            try:
                informacoes['CNPJ'] = driver.find_element(
                    By.XPATH, "/html/body/div[1]/div/div/div/div[1]/p[1]/span[1]/b").text.strip()
            except Exception as e:
                informacoes['CNPJ'] = 'N/A'
            return informacoes
        else:
            # Se não encontrou nenhum link válido
            informacoes = {
                'CNPJ': 'N/A'
            }
            return informacoes

    finally:
        # Fechando o navegador ao finalizar
        driver.quit()


def buscar_informacoes_econodata(cnpj):
    # Configuração do ChromeDriver
    chrome_options = Options()
    # Instancia o driver
    driver = webdriver.Chrome(service=Service(
        ChromeDriverManager().install()), options=chrome_options)
    # Define o tempo de espera implícito de 10 segundos
    driver.implicitly_wait(10)

    try:
        # Realiza a busca no Google
        driver.get("https://www.google.com")
        search_box = driver.find_element(By.NAME, "q")
        search_box.send_keys(f"econodata - {cnpj}")
        search_box.send_keys(Keys.RETURN)


        # Verifica se há resultados na página
        try:
            first_result = driver.find_element(By.XPATH, "(//h3)[1]/../..")
            first_result.click()
        except:
            return {"nome_fantasia": "N/A", "logradouro": "N/A", "bairro": "N/A", "cep": "N/A"}

        # Tenta extrair o nome fantasia usando o XPath fornecido
        try:
            p_nome_fantasia = driver.find_element(
                By.XPATH, "/html/body/div[2]/div[2]/div[2]/div[2]/div[1]/div[5]/div[2]/div[2]/div[3]/div/div[2]/p")
            nome_fantasia = p_nome_fantasia.text.strip()
        except:
            nome_fantasia = "N/A"

        # Tenta extrair o logradouro usando itemprop='streetAddress'
        try:
            div_logradouro = driver.find_element(
                By.XPATH, "//div[@itemprop='streetAddress']")
            logradouro = div_logradouro.text.strip()
        except:
            logradouro = "N/A"

        # Tenta extrair o bairro usando o XPath fornecido
        try:
            span_bairro = driver.find_element(
                By.XPATH, "/html/body/div[2]/div[2]/div[2]/div[2]/div[1]/div[2]/div[1]/div[2]/div[2]/div[2]/span")
            bairro = span_bairro.text.strip()
        except:
            bairro = "N/A"

        # Tenta extrair o CEP usando itemprop='postalCode'
        try:
            div_cep = driver.find_element(
                By.XPATH, "//div[@itemprop='postalCode']")
            cep_element = div_cep.find_element(By.XPATH, "./span")
            cep = cep_element.text.strip()
        except:
            cep = "N/A"

        return {"nome_fantasia": nome_fantasia, "logradouro": logradouro, "bairro": bairro, "cep": cep}

    except Exception as e:
        print(f"Erro ao buscar informações: {e}")

    finally:
        driver.quit()


# Exemplo de uso
cnpj = "33.435.231/0001-87"
resultado = buscar_informacoes_econodata(cnpj)
print(resultado)
