import requests
import re


def clean_cnpj(cnpj):
    return re.sub(r'\D', '', cnpj)


def get_company_info(cnpj):
    if cnpj == 'N/A':
        return {
            "nome_fantasia": 'N/A',
            "cep": 'N/A',
            "tipo_logradouro": 'N/A',
            "logradouro": 'N/A',
            "numero": 'N/A',
            "complemento": 'N/A',
            "bairro": 'N/A'

        }

    url = "https://cnpj.biz/api/v2/empresas/cnpj"
    headers = {
        'Content-Type': "application/json",
        'authorization': "Bearer FlCrASrWb2w4O6cB3ZY88fosbMB04Ld1xSX76Ka0thS30UHYz8gg3UCRZPO5",
        'Accept': "application/json"
    }
    body = {"cnpj": clean_cnpj(cnpj)}

    response = requests.post(url, json=body, headers=headers)

    if response.status_code == 200:
        result = response.json()
        return {
            "nome_fantasia": result['nome_fantasia'],
            "cep": result['endereco']['cep'],
            "tipo_logradouro": result['endereco']['tipo_logradouro'],
            "logradouro": result['endereco']['logradouro'],
            "numero": result['endereco']['numero'],
            "complemento": result['endereco']['complemento'],
            "bairro": result['endereco']['bairro']
        }
    else:
        return {"error": "Failed to fetch data", "status_code": response.status_code, "response": response.text}