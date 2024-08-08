# rgmlLib.py

import subprocess
import configparser
import os
import time

def runCmdLine(cmd):
    """
    Executa um comando da linha de comando e retorna a saída.

    Args:
    cmd (list): Comando a ser executado e seus argumentos em forma de lista.

    Returns:
    str: Saída do comando executado.

    Raises:
    Exception: Se ocorrer um erro durante a execução do comando.
    """
    try:
        return subprocess.check_output(cmd, stderr=subprocess.STDOUT).decode()
    except subprocess.CalledProcessError as e:
        raise Exception("Falha ao executar o comando: {}".format(e.output.decode())) from e


def setProperty(filepath, section, property_name, value):
    """
    Define o valor de uma propriedade em um arquivo de configuração.

    Args:
        filepath (str): Caminho para o arquivo de configuração.
        section (str): Seção no arquivo de configuração.
        property_name (str): Nome da propriedade a ser definida.
        value (str): Valor a ser atribuído à propriedade.
    """
    config = configparser.ConfigParser()
    config.read(filepath)

    if not config.has_section(section):
        config.add_section(section)
    
    config.set(section, property_name, value)

    with open(filepath, 'w') as configfile:
        config.write(configfile)

def getProperty(filepath, section, property_name):
    """
    Obtém o valor de uma propriedade em um arquivo de configuração.

    Args:
        filepath (str): Caminho para o arquivo de configuração.
        section (str): Seção no arquivo de configuração.
        property_name (str): Nome da propriedade a ser obtida.

    Returns:
        str: Valor da propriedade. Retorna None se a propriedade ou seção não existir.
    """
    config = configparser.ConfigParser()
    config.read(filepath)

    if config.has_section(section) and config.has_option(section, property_name):
        return config.get(section, property_name)
    else:
        return None
    
