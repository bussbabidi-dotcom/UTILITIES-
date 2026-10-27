import json
import os
import re
from pathlib import Path

import pythoncom
import win32com.client

# ==========================================================
# CONFIGURAÇÃO
# ==========================================================

CONFIG = Path(__file__).resolve().parent / "photo_config.json"

PASTA_PADRAO = r"\\192.168.0.29\hd-62\ÁLBUNS COLAÇÃO 2025.2\CAPAS\FOTO"

# Qualidade máxima do JPG no Photoshop
QUALIDADE_JPG = 12

# Tecla de atalho padrão
TECLA = "F6"


def carregar_config():
    """Lê photo_config.json (pasta, tecla, posição do widget...)."""

    try:
        dados = json.loads(CONFIG.read_text(encoding="utf-8"))
    except Exception:
        dados = {}

    dados.setdefault("pasta_saida", PASTA_PADRAO)
    dados.setdefault("tecla", TECLA)
    return dados


def salvar_config(dados):
    CONFIG.write_text(json.dumps(dados, indent=2, ensure_ascii=False), encoding="utf-8")


def limpar_nome(nome):
    """
    Remove caracteres que não podem ser usados em nomes
    de arquivos no Windows.
    """

    nome = re.sub(r'[<>:"/\\|?*]', '', nome)
    nome = nome.strip()

    if not nome:
        nome = "SEM_NOME"

    return nome


def exportar_jpg(pasta_saida):
    """
    Salva o documento ativo do Photoshop como JPG, com o nome
    da camada selecionada. Retorna (ok, mensagem, caminho).
    """

    pythoncom.CoInitialize()
    try:
        os.makedirs(pasta_saida, exist_ok=True)

        # Conecta ao Photoshop que já está aberto
        ps = win32com.client.Dispatch("Photoshop.Application")

        if ps.Documents.Count == 0:
            return False, "Nenhum PSD está aberto.", None

        documento = ps.ActiveDocument

        nome_camada = limpar_nome(documento.ActiveLayer.Name)

        caminho = os.path.join(pasta_saida, nome_camada + ".jpg")

        # Se já existir, substitui
        if os.path.exists(caminho):
            os.remove(caminho)

        jpeg_options = win32com.client.Dispatch("Photoshop.JPEGSaveOptions")
        jpeg_options.Quality = QUALIDADE_JPG
        jpeg_options.FormatOptions = 1
        jpeg_options.Matte = 1

        documento.SaveAs(caminho, jpeg_options, True)

        if not os.path.isfile(caminho):
            return False, "O Photoshop não criou o arquivo JPG.", None

        return True, f"{nome_camada}.jpg", caminho

    except Exception as erro:
        return False, f"{type(erro).__name__}: {erro}", None

    finally:
        pythoncom.CoUninitialize()


# ==========================================================
# USO PELO TERMINAL (F6 exporta, Esc encerra)
# ==========================================================

if __name__ == "__main__":

    import keyboard

    config = carregar_config()
    pasta = config["pasta_saida"]
    tecla = config["tecla"]

    def ao_apertar():
        ok, mensagem, caminho = exportar_jpg(pasta)
        if ok:
            print("====================================")
            print("JPG SALVO COM SUCESSO!")
            print(f"Arquivo: {caminho}")
            print("====================================")
        else:
            print("ERRO:")
            print(mensagem)

    keyboard.add_hotkey(tecla, ao_apertar, suppress=False)

    print("====================================")
    print("EXPORTADOR JPG INICIADO")
    print("====================================")
    print(f"Pasta: {pasta}")
    print(f"Pressione {tecla} para exportar o JPG ou Esc para encerrar.")

    try:
        keyboard.wait("esc")
    finally:
        keyboard.unhook_all()
