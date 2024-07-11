import pandas as pd
from concurrent.futures import ThreadPoolExecutor, as_completed
from buscar_google import buscar_informacoes_econodata, buscar_informacoes_google

# Função para normalizar string para UPPERCASE e sem espaços em branco no final


def normalizar_valor(valor):
    if isinstance(valor, str):
        # Remove espaços em branco no início e fim, e converte para UPPERCASE
        return valor.strip().upper()
    return 'N/A'  # Retorna 'N/A' se o valor não for uma string (caso de NaN)

# Função para buscar informações utilizando threads


def buscar_informacoes_thread(registro):
    cliente = registro['Cliente']
    estado = registro['Estado']
    cidade = registro['Cidade']

    # Realiza a busca de informações do cnpj.biz
    retorno_cnpj = buscar_informacoes_google(cliente, estado, cidade)
    # Realiza a busca de informações do econodata.com.br
    retorno_enderco = buscar_informacoes_econodata(retorno_cnpj['CNPJ'])
    #retorno = buscar_informacoes_econodata(cnpj)

    # Atualiza o registro com as informações encontradas no cnpj
    registro.update(retorno_cnpj)
    registro.update(retorno_enderco)

    # Atualiza o registro com as informações encontradas

# Caminho para o arquivo Excel
caminho_arquivo = 'arquivo.xlsx'

# Carregando o arquivo Excel
xls = pd.ExcelFile(caminho_arquivo)

# Lista para armazenar os registros de todas as planilhas
registros = []

# Iterando sobre todas as planilhas no arquivo Excel
for nome_planilha in xls.sheet_names:
    # Carregando a planilha atual
    df = pd.read_excel(xls, sheet_name=nome_planilha)

    # Iterando pelas linhas da planilha atual
    for indice, linha in df.iterrows():
        # Lendo os campos e aplicando a normalização
        equipamento = normalizar_valor(linha.get('Equipamento', ''))
        cliente = normalizar_valor(linha.get('Cliente', ''))
        estado = normalizar_valor(linha.get('Estado', ''))
        cidade = normalizar_valor(linha.get('Cidade', ''))
        marca_icp = normalizar_valor(linha.get('Marca ICP', ''))
        modelo = normalizar_valor(linha.get('Modelo', ''))

        # Criando um dicionário para representar a linha atual
        registro = {
            'Equipamento': equipamento,
            'Cliente': cliente,
            'Estado': estado,
            'Cidade': cidade,
            'Marca ICP': marca_icp,
            'Modelo': modelo,
            'Planilha': nome_planilha  # Adicionando o nome da planilha
        }

        # Adicionando o registro à lista de registros
        registros.append(registro)

# Definindo o número máximo de threads
max_threads = 15

# Criando um ThreadPoolExecutor para limitar a quantidade de threads
with ThreadPoolExecutor(max_workers=max_threads) as executor:
    # Criando uma lista para armazenar os futuros
    futures = [executor.submit(buscar_informacoes_thread, registro)
               for registro in registros]

    # Aguardando todas as threads terminarem
    for future in as_completed(futures):
        future.result()

# Convertendo a lista de registros em um DataFrame do pandas
df_registros = pd.DataFrame(registros)

# Escrevendo o DataFrame em um arquivo Excel
caminho_saida_excel = 'registros.xlsx'
df_registros.to_excel(caminho_saida_excel, index=False)

print(f"Arquivo Excel gerado com sucesso: {caminho_saida_excel}  Total: {len(registros)}")
