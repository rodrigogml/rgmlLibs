import os
import mysql.connector
from mysql.connector import Error

def getMySQLConnection(config):
    """
    Estabelece uma conexão com um banco de dados MySQL utilizando os parâmetros de configuração fornecidos.

    Este método espera receber um dicionário 'config' contendo as configurações necessárias para a conexão.
    Estas configurações podem ser provenientes de um arquivo de propriedades ou podem ser definidas diretamente no sistema.

    Args:
    config (dict): Um dicionário contendo as configurações de conexão. Os seguintes parâmetros são esperados:
        - mysql_host (str): O endereço do host do banco de dados MySQL.
        - mysql_user (str): O nome de usuário para acessar o banco de dados.
        - mysql_password (str): A senha para acessar o banco de dados.
        - mysql_database (str): O nome do banco de dados específico a ser acessado.

    Retorna:
    mysql.connector.connection.MySQLConnection: Uma conexão ativa com o banco de dados MySQL.

    Levanta:
    ValueError: Se algum dos parâmetros obrigatórios não estiver presente no dicionário 'config'.
    Exception: Se ocorrer um erro ao tentar estabelecer a conexão com o banco de dados.

    Exemplo:
    config = {
        'mysql_host': 'localhost',
        'mysql_user': 'usuario',
        'mysql_password': 'senha',
        'mysql_database': 'nome_do_banco'
    }
    connection = getMySQLConnection(config)
    """    
    required_keys = ['mysql_host', 'mysql_user', 'mysql_password', 'mysql_database']
    for key in required_keys:
        if key not in config:
            raise ValueError(f"Atributo necessário '{key}' não encontrado no arquivo de propriedades.")

    # Estabelecendo a conexão
    try:
        connection = mysql.connector.connect(
            host=config['mysql_host'],
            user=config['mysql_user'],
            passwd=config['mysql_password'],
            database=config['mysql_database']
        )
        return connection
    except Error as err:
        raise Exception(f"Erro ao conectar ao banco de dados: {err}")
