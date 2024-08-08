def extractPDFTextContent(pdfPath):
    """
    Extrai os blocos de texto de um arquivo PDF e retorna o conteúdo em um array.
    
    :param pdfPath: Caminho para o arquivo PDF.
    :return: Array com blocos de texto extraídos do PDF.
    """
    import fitz  # PyMuPDF

    # Abrir o documento PDF
    doc = fitz.open(pdfPath)
    textBlocks = []

    for page in doc:
        # Extrair blocos de texto da página
        blocks = page.get_text("blocks")
        # Cada bloco é uma tupla, onde o índice 4 contém o texto
        for block in blocks:
            textBlocks.append(block[4])

    # Fechar o documento
    doc.close()

    return textBlocks