# rgmlADB.py

from rgmlLib import runCmdLine
import subprocess
import os
import time

def adbSwipe(x1, y1, x2, y2, delay=None):
    """
    Executa um comando swipe no adb.

    Args:
    x1, y1, x2, y2 (int): Coordenadas para o comando swipe.
    delay (int, optional): Delay após o comando swipe.

    Raises:
    Exception: Se ocorrer um erro durante a execução do comando adbSwipe.
    """
    adb_path = os.path.join(os.path.expanduser("~"), "AppData", "Local", "Android", "Sdk", "platform-tools", "adb")
    cmd = [adb_path, "shell", "input", "swipe", str(x1), str(y1), str(x2), str(y2)]

    if delay:
        cmd.append(str(delay))
        time.sleep(1)

    try:
        result = runCmdLine(cmd)
        return result
    except Exception as e:
        raise Exception("Falha ao executar adbSwipe: {}".format(str(e))) from e


def adbTap(x, y):
    """
    Executa um comando tap no adb.

    Args:
    x, y (int): Coordenadas para o comando tap.

    Raises:
    Exception: Se ocorrer um erro durante a execução do comando adbTap.
    """
    adb_path = os.path.join(os.path.expanduser("~"), "AppData", "Local", "Android", "Sdk", "platform-tools", "adb")
    cmd = [adb_path, "shell", "input", "tap", str(x), str(y)]

    try:
        result = runCmdLine(cmd)
        time.sleep(1)
        return result
    except Exception as e:
        raise Exception("Falha ao executar adbTap: {}".format(str(e))) from e

def adbPasteClipboard():
    """
    Cola o conteúdo do clipboard no campo de texto ativo no dispositivo Android usando adb.
    """
    adb_path = os.path.join(os.path.expanduser("~"), "AppData", "Local", "Android", "Sdk", "platform-tools", "adb")

    # Comando para colar o texto do clipboard
    cmd_paste = [adb_path, "shell", "input", "keyevent", "KEYCODE_PASTE"]

    try:
        result = runCmdLine(cmd_paste)
        time.sleep(1)
        return "Conteúdo do clipboard colado com sucesso."
    except Exception as e:
        raise Exception("Falha ao executar adbPasteClipboard: {}".format(str(e))) from e


def adbSendToClipboard(text):
    """
    Envia um texto para o clipboard do dispositivo Android usando adb.

    Args:
    text (str): Texto a ser enviado para o clipboard.
    """
    adb_path = os.path.join(os.path.expanduser("~"), "AppData", "Local", "Android", "Sdk", "platform-tools", "adb")

    # Preparando o texto para ser enviado através do comando (escapando caracteres necessários)
    text_escaped = text.replace("'", "\\'")

    # Comando para enviar o texto para o clipboard
    cmd_clipboard = [adb_path, "shell", "am", "broadcast", "-a", "clipper.set", "-e", "text", f"'{text_escaped}'"]

    try:
        result = runCmdLine(cmd_clipboard)
        time.sleep(1)
        return "Texto enviado para o clipboard com sucesso."
    except Exception as e:
        raise Exception("Falha ao executar adbSendToClipboard: {}".format(str(e))) from e


def adbText(text):
    """
    Envia um texto para o campo de texto ativo no dispositivo Android usando adb.
    ATENÇÃO: ESTE MÉTODO NÃO ACEITA CARACTERES QUE NÃO SEJAM UNICODE (LIMITAÇÃO DO ADB!)

    Args:
    text (str): Texto a ser enviado.

    Raises:
    Exception: Se ocorrer um erro durante a execução do comando adbText.
    """
    adb_path = os.path.join(os.path.expanduser("~"), "AppData", "Local", "Android", "Sdk", "platform-tools", "adb")

    # Tratando o texto para ser compatível com o comando adb
    # Substituindo espaços por '%s' e outros caracteres especiais conforme necessário
    treated_text = text.replace(" ", "%s")

    cmd = [adb_path, "shell", "input", "text", treated_text]

    try:
        result = runCmdLine(cmd)
        time.sleep(1)
        return result
    except Exception as e:
        raise Exception("Falha ao executar adbText: {}".format(str(e))) from e


def adbDoubleTap(x, y, interval=0.05):
    """
    Executa um duplo toque na tela do dispositivo Android usando adb.

    Args:
    x, y (int): Coordenadas X e Y na tela para o duplo toque.
    interval (float, opcional): Intervalo entre os toques, em segundos.
    """
    adb_path = os.path.join(os.path.expanduser("~"), "AppData", "Local", "Android", "Sdk", "platform-tools", "adb")
    tap_cmd = ["shell", "input", "tap", str(x), str(y)]

    subprocess.run([adb_path] + tap_cmd, shell=True)
    subprocess.run([adb_path] + tap_cmd, shell=True)

