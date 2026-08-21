# -*- coding: utf-8 -*-
"""Configuracao compartilhada do OCR (Tesseract) para os scripts de Instrumentos.

POR QUE ESTE ARQUIVO EXISTE
---------------------------
Cada etapa deste pipeline roda como um processo Python separado (a interface chama
analise_instrumentos.py, depois enhance.py, depois build_final.py). O pacote Python
`pytesseract` NAO descobre sozinho onde o Tesseract esta instalado no Windows: ele
simplesmente tenta executar o comando "tesseract" e conta com o PATH do sistema.

O instalador padrao do Tesseract no Windows (UB Mannheim) NAO adiciona o programa ao
PATH por default. Resultado: se so uma das etapas souber o caminho do executavel, as
outras quebram com TesseractNotFoundError -- mesmo com o Tesseract instalado e
funcionando. Foi exatamente isso que aconteceu no pipeline original.

Este modulo centraliza essa configuracao para que TODAS as etapas enxerguem o Tesseract
da mesma forma, e -- importante -- VERIFICA que ele realmente executa, em vez de apenas
supor que existe.
"""
import os
import shutil
import subprocess

# Caminhos onde o instalador padrao do Windows costuma colocar o executavel.
CAMINHOS_PADRAO = [
    r"C:\Program Files\Tesseract-OCR\tesseract.exe",
    r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
]


def localizar_tesseract():
    """Devolve o caminho do executavel do Tesseract, ou '' se nao encontrar.

    Ordem de busca: variavel de ambiente TESSERACT_CMD (preenchida pela interface a
    partir do caminho escolhido na Tela 3) -> caminhos padrao do Windows -> PATH.
    """
    do_ambiente = (os.environ.get("TESSERACT_CMD") or "").strip().strip('"')
    if do_ambiente and os.path.exists(do_ambiente):
        return do_ambiente
    for c in CAMINHOS_PADRAO:
        if os.path.exists(c):
            return c
    return shutil.which("tesseract") or ""


def configurar_ocr():
    """Configura o pytesseract e confirma que o Tesseract executa de verdade.

    Devolve (disponivel: bool, detalhe: str). O `detalhe` e' uma frase em portugues
    pronta para ser mostrada no log ou numa caixa de aviso.
    """
    try:
        import pytesseract
    except Exception as e:
        return False, ("o pacote Python 'pytesseract' nao esta instalado "
                       f"({e}). Instale com: pip install pytesseract")

    caminho = localizar_tesseract()
    if caminho:
        pytesseract.pytesseract.tesseract_cmd = caminho

    alvo = pytesseract.pytesseract.tesseract_cmd
    try:
        proc = subprocess.run([alvo, "--version"], capture_output=True, text=True, timeout=60)
    except FileNotFoundError:
        return False, ("o programa Tesseract nao foi encontrado. Ele e' um programa "
                       "separado do Python: instale o Tesseract OCR (build da UB Mannheim) "
                       "e, se ja estiver instalado, informe o caminho do tesseract.exe "
                       "na tela de parametros da interface.")
    except Exception as e:
        return False, f"o Tesseract foi encontrado em '{alvo}' mas nao pode ser executado ({e})."

    if proc.returncode != 0:
        return False, (f"o Tesseract em '{alvo}' respondeu com erro "
                       f"(codigo {proc.returncode}).")

    saida = (proc.stdout or proc.stderr or "").strip().splitlines()
    versao = saida[0] if saida else "versao desconhecida"
    return True, f"{alvo} ({versao})"


def idiomas_instalados():
    """Lista os idiomas de OCR instalados (ex.: ['eng', 'por', 'osd']).

    Devolve [] se nao conseguir consultar. Util para avisar a usuaria antes de rodar,
    caso ela escolha OCR em portugues sem ter baixado o pacote de idioma 'por'.
    """
    try:
        import pytesseract
    except Exception:
        return []
    alvo = pytesseract.pytesseract.tesseract_cmd or localizar_tesseract()
    if not alvo:
        return []
    try:
        proc = subprocess.run([alvo, "--list-langs"], capture_output=True, text=True, timeout=60)
    except Exception:
        return []
    linhas = (proc.stdout or "").strip().splitlines()
    return [l.strip() for l in linhas[1:] if l.strip()]
