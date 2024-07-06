import pandas as pd
import json
import threading

from buscar_google import buscar_informacoes_google

# Script para ler um arquivo Excel contendo múltiplas planilhas,
# normalizar os valores para UPPERCASE e remover espaços em branco no final,
# e escrever os registros em um arquivo JSON no final do processamento.

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

    # Realiza a busca de informações do Google
    retorno = buscar_informacoes_google(cliente, estado, cidade)

    # Atualiza o registro com as informações encontradas
    registro.update(retorno)


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

# Criando uma lista de threads para executar a busca de informações
threads = []
for registro in registros:
    thread = threading.Thread(
        target=buscar_informacoes_thread, args=(registro,))
    threads.append(thread)
    thread.start()

# Esperando todas as threads terminarem
for thread in threads:
    thread.join()

# Escrevendo a lista de registros em um arquivo JSON
caminho_saida_json = 'registros.json'
with open(caminho_saida_json, 'w', encoding='utf-8') as f:
    json.dump(registros, f, ensure_ascii=False, indent=4)

print(f"Arquivo JSON gerado com sucesso: {caminho_saida_json}  Total: {len(registros)}")
