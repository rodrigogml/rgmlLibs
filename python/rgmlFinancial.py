# rgmlFinancial.py

import requests
from datetime import datetime, timedelta
import urllib.parse

def downloadBcbSgsData(seriesId, startDate, endDate):
    """
    Baixa os dados de uma série específica do SGS do Banco Central do Brasil.

    Args:
        seriesId (int): O identificador da série a ser baixada.
        startDate (str): Data de início no formato 'dd/mm/aaaa'.
        endDate (str): Data de término no formato 'dd/mm/aaaa'.

    Returns:
        list: Lista de dicionários contendo os dados da série.

    Raises:
        Exception: Se a requisição para a API do SGS não retornar um status code 200.
    """
    baseUrl = "https://api.bcb.gov.br/dados/serie/bcdata.sgs.{}/dados"
    url = baseUrl.format(seriesId)
    params = {
        'formato': 'json',
        'dataInicial': startDate,
        'dataFinal': endDate
    }

    # print(f"Recuperando URL: {url}, Data Inicial: {startDate}, Data Final: {endDate}")
    response = requests.get(url, params=params)
    if response.status_code == 200:
        # print("Resposta:")
        # print(response.json())
        return response.json()
    else:
        raise Exception(f"Erro ao baixar os dados da série {seriesId}: HTTP {response.status_code}")
