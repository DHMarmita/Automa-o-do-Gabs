from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager


def buscar_informacoes_google(cliente, estado, cidade):
    # Configuração do webdriver do Chrome
    options = Options()
    # Rodar em modo headless (sem abrir janelas do navegador)
    options.headless = True

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
