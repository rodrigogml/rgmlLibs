import unicodedata
import re

def removeEspecialCaracteres(inputData):
    """
    Remove caracteres especiais de uma string ou de cada string em um array, convertendo letras acentuadas para não acentuadas.
    
    :param inputData: Uma string ou um array de strings para processamento.
    :return: A string ou o array de strings com caracteres especiais removidos.
    """
    def removeCaracteres(text):
        # Normalizar para remover acentuação
        nfkd_form = unicodedata.normalize('NFKD', text)
        only_ascii = nfkd_form.encode('ASCII', 'ignore').decode('ASCII')
        
        # Remover caracteres especiais mantendo espaços
        cleanText = re.sub(r'[^\w\s]', '', only_ascii)
        return cleanText

    if isinstance(inputData, str):
        return removeCaracteres(inputData)
    elif isinstance(inputData, list):
        return [removeCaracteres(item) for item in inputData]
    else:
        raise ValueError("O input deve ser uma string ou uma lista de strings.")

def toLowerCase(inputData):
    """
    Converte o conteúdo de uma string ou de cada string em um array para minúsculas.
    
    :param inputData: Uma string ou um array de strings para processamento.
    :return: A string ou o array de strings convertido(s) para minúsculas.
    """
    if isinstance(inputData, str):
        return inputData.lower()
    elif isinstance(inputData, list):
        return [item.lower() for item in inputData]
    else:
        raise ValueError("O input deve ser uma string ou uma lista de strings.")

def removeStopwordsFromFile(inputData, stopwordsFilePath):
    """
    Remove stopwords de uma string ou de cada string em um array, baseando-se em um arquivo de stopwords.
    
    :param inputData: Uma string ou um array de strings para processamento.
    :param stopwordsFilePath: Caminho para o arquivo de stopwords, onde cada linha contém uma stopword,
                              ignorando linhas que começam com # por serem consideradas comentários.
    :return: A string ou o array de strings com stopwords removidas.
    """
    # Ler as stopwords do arquivo, ignorando linhas de comentários
    with open(stopwordsFilePath, 'r', encoding='utf-8') as file:
        stopwords = [line.strip() for line in file if not line.startswith('#') and line.strip()]
    
    def removeWords(text):
        # Dividir o texto em palavras
        words = text.split()
        # Remover stopwords
        filteredWords = [word for word in words if word not in stopwords]
        # Reconstruir o texto
        return " ".join(filteredWords)

    if isinstance(inputData, str):
        return removeWords(inputData)
    elif isinstance(inputData, list):
        return [removeWords(item) for item in inputData]
    else:
        raise ValueError("O input deve ser uma string ou uma lista de strings.")
