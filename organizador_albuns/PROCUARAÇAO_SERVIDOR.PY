"""
====================================================================
              ORGANIZADOR AUTOMÁTICO DE ÁLBUNS
====================================================================

REGRAS:

20X20
    -> 10 lâminas

30X20
    -> 15 lâminas

24X34
    -> 20 lâminas
    -> TAMBÉM procura o 20X20 da mesma pessoa
    -> 20X20 = 10 lâminas
    -> 24X34 = 20 lâminas


SERVIDOR:

\\\\192.168.0.29\\hd-61\\ÁLBUNS COLAÇÃO


ESTRUTURA:

ÁLBUNS COLAÇÃO
│
├── CAPAS
│   ├── capa 20x20
│   ├── capas 24X34 20X20
│   └── CAPAS 30X20
│
├── 1° REMESSA
├── 2° REMESSA
├── 3° REMESSA
├── ...
└── 9° REMESSA
    │
    └── #ÁLBUNS PRONTOS
        │
        └── CLIENTE
            │
            └── PESSOA
                │
                ├── Álbum 20x20
                │   └── JPEG
                │
                └── Álbum 34x24
                    └── JPEG


PLANILHA:

COLUNA A = NOME
COLUNA B = CÓDIGO
COLUNA C = PRODUTO


MODO TESTE:

True
    -> não copia
    -> não move
    -> não apaga

False
    -> copia os arquivos


====================================================================
"""

from pathlib import Path
from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string

import argparse
import csv
import os
import re
import shutil
import struct
import sys
from importlib.machinery import SourceFileLoader
import unicodedata
import urllib.request


def _configurar_saida():

    for canal in (sys.stdout, sys.stderr):

        if hasattr(canal, "reconfigure"):

            try:

                canal.reconfigure(encoding="utf-8", errors="replace")

            except Exception:

                pass


_configurar_saida()


# ====================================================================
# CONFIGURAÇÕES
# ====================================================================


# --------------------------------------------------------------------
# SERVIDOR
# --------------------------------------------------------------------

SERVIDOR = Path(
    r"D:\ÁLBUNS COLAÇÃO 2025.2"
)


# --------------------------------------------------------------------
# PASTA PRINCIPAL
# --------------------------------------------------------------------

PASTA_PRINCIPAL = "ÁLBUNS COLAÇÃO"


# --------------------------------------------------------------------
# PASTA DOS ÁLBUNS PRONTOS
# --------------------------------------------------------------------

PASTA_PRONTOS = "#ÁLBUNS PRONTOS"


# --------------------------------------------------------------------
# PASTA GERAL DE CAPAS
# --------------------------------------------------------------------

PASTA_CAPAS = "CAPAS"


# --------------------------------------------------------------------
# PLANILHA
#
# ALTERE AQUI PARA O LOCAL DA SUA PLANILHA
# --------------------------------------------------------------------

PLANILHA = Path(
    r"C:\Users\REIS\Documents\PAPA 2\LOTE 13\CREDENCIAIS.xlsx"
)


# --------------------------------------------------------------------
# DESTINO
#
# ALTERE AQUI PARA ONDE QUER COPIAR OS ÁLBUNS
# --------------------------------------------------------------------

DESTINO = Path(
    r"C:\Users\REIS\Documents\PAPA 2\LOTE 13"
)


# --------------------------------------------------------------------
# MODO TESTE
# --------------------------------------------------------------------
#
# True  = somente testa
# False = copia os arquivos
#
# --------------------------------------------------------------------

MODO_TESTE = True

SEM_PAUSA = False

PASTA_SCRIPT = Path(__file__).resolve().parent

PLANILHA_BAIXADA = PASTA_SCRIPT / "planilha_organizador.xlsx"

RELATORIO = PASTA_SCRIPT / "relatorio_organizador.csv"


# --------------------------------------------------------------------
# COLUNAS DA PLANILHA
# --------------------------------------------------------------------

COLUNA_NOME = 1

COLUNA_CODIGO = 2

COLUNA_PRODUTO = 3


# ====================================================================
# QUANTIDADE DE LÂMINAS
# ====================================================================

QUANTIDADE_LAMINAS = {

    "20X20": 10,

    "24X34": 20,

    "30X20": 15

}


# ====================================================================
# NORMALIZAR TEXTO
# ====================================================================

def normalizar(texto):

    if texto is None:

        return ""

    texto = str(texto).strip()

    # --------------------------------------------------------------
    # REMOVER ACENTOS
    # --------------------------------------------------------------

    texto = unicodedata.normalize(
        "NFD",
        texto
    )

    texto = "".join(
        caractere
        for caractere in texto
        if unicodedata.category(
            caractere
        ) != "Mn"
    )

    # --------------------------------------------------------------
    # MINÚSCULO
    # --------------------------------------------------------------

    texto = texto.lower()

    # --------------------------------------------------------------
    # CARACTERES ESPECIAIS
    # --------------------------------------------------------------

    texto = re.sub(
        r"[^a-z0-9]+",
        " ",
        texto
    )

    # --------------------------------------------------------------
    # ESPAÇOS DUPLICADOS
    # --------------------------------------------------------------

    texto = re.sub(
        r"\s+",
        " ",
        texto
    )

    return texto.strip()


# ====================================================================
# NORMALIZAR TAMANHO
# ====================================================================

def normalizar_tamanho(texto):

    if texto is None:

        return None

    texto = str(texto).upper()

    # --------------------------------------------------------------
    # VÁRIOS TIPOS DE X
    # --------------------------------------------------------------

    texto = texto.replace("×", "X")

    texto = texto.replace("*", "X")

    texto = texto.replace("/", "X")

    # --------------------------------------------------------------
    # REMOVER ESPAÇOS
    # --------------------------------------------------------------

    texto = re.sub(
        r"\s+",
        "",
        texto
    )

    # --------------------------------------------------------------
    # PROCURAR TAMANHO
    # --------------------------------------------------------------

    encontrados = re.findall(
        r"(20X20|24X34|34X24|30X20)",
        texto
    )

    if not encontrados:

        return None

    tamanho = encontrados[-1]

    # --------------------------------------------------------------
    # 34X24 E 24X34 SÃO O MESMO ÁLBUM
    # --------------------------------------------------------------

    if tamanho == "34X24":

        tamanho = "24X34"

    if tamanho not in QUANTIDADE_LAMINAS:

        return None

    return tamanho


# ====================================================================
# IDENTIFICAR TAMANHO DO PRODUTO
# ====================================================================

def extrair_tamanho(produto):

    return normalizar_tamanho(
        produto
    )


# ====================================================================
# PLANO -> TAMANHO DA PASTA NO SERVIDOR
# ====================================================================

ATALHOS_PLANO = (
    (("essencial", "essential"), "20X20"),
    (("prime",), "30X20"),
    (("moscou", "moskow", "moscovo"), "30X20"),
    (("berlim", "berlin", "berlinense"), "24X34"),
    (("diamond", "diamante"), "24X34"),
)


def tamanho_do_plano(plano):

    tamanho = extrair_tamanho(plano)

    if tamanho:

        return tamanho

    palavras = set(normalizar(plano).split())

    for chaves, tamanho in ATALHOS_PLANO:

        if any(chave in palavras for chave in chaves):

            return tamanho

    return None


def tamanhos_mencionados(texto):

    if texto is None:

        return set()

    texto = str(texto).upper()

    texto = texto.replace("×", "X").replace("*", "X").replace("/", "X")

    texto = re.sub(r"\s+", "", texto)

    achados = re.findall(r"(20X20|24X34|34X24|30X20|60X20)", texto)

    # 60x20 é a lâmina aberta do álbum 30x20.
    equivalentes = {
        "34X24": "24X34",
        "60X20": "30X20",
    }

    return {
        equivalentes.get(item, item)
        for item in achados
    }


def pontuacao_nome(nome_alvo, nome_candidato):

    nome_alvo = normalizar(nome_alvo)

    nome_candidato = normalizar(nome_candidato)

    palavras_alvo = set(nome_alvo.split())

    palavras_candidato = set(nome_candidato.split())

    comuns = palavras_alvo & palavras_candidato

    if (
        not nome_alvo
        or not comuns
        or (
            nome_alvo not in nome_candidato
            and len(comuns) < 2
        )
    ):

        return -1

    pontuacao = 1

    if nome_candidato == nome_alvo:

        pontuacao += 4

    if nome_alvo in nome_candidato:

        pontuacao += 3

    pontuacao += len(comuns)

    return pontuacao


def texto_celula(valor):

    if valor is None:

        return ""

    if isinstance(valor, float) and valor.is_integer():

        return str(int(valor))

    return str(valor).strip()


def link_exportacao(link):

    achado = re.search(r"/spreadsheets/d/([\w-]+)", link or "")

    if not achado:

        raise ValueError(f"Link da planilha inválido: {link}")

    gid = re.search(r"[#&?]gid=(\d+)", link)

    url = (
        "https://docs.google.com/spreadsheets/d/"
        + achado.group(1)
        + "/export?format=xlsx"
    )

    if gid:

        url += "&gid=" + gid.group(1)

    return url


def baixar_planilha_link(link, destino):

    destino = Path(destino)

    temporario = destino.with_suffix(".baixando")

    try:

        urllib.request.urlretrieve(link_exportacao(link), temporario)

        if temporario.read_bytes()[:2] != b"PK":

            raise ValueError(
                "o Google não devolveu a planilha "
                "(o link está compartilhado como 'qualquer pessoa com o link'?)"
            )

        temporario.replace(destino)

    except Exception as erro:

        temporario.unlink(missing_ok=True)

        if not destino.exists():

            raise ValueError(f"Não consegui baixar a planilha: {erro}")

        print(f"(não consegui baixar a planilha agora — usando a última cópia: {erro})")

    return destino


def pausa(mensagem="\nPressione ENTER para fechar..."):

    if SEM_PAUSA:

        return

    try:

        if not sys.stdin.isatty():

            return

    except Exception:

        return

    input(mensagem)


# ====================================================================
# COMPARAR NOMES
# ====================================================================

def nomes_correspondem(
    nome1,
    nome2
):

    n1 = normalizar(nome1)

    n2 = normalizar(nome2)

    if not n1 or not n2:

        return False

    # --------------------------------------------------------------
    # IGUAIS
    # --------------------------------------------------------------

    if n1 == n2:

        return True

    # --------------------------------------------------------------
    # UM CONTÉM O OUTRO
    # --------------------------------------------------------------

    if n1 in n2:

        return True

    if n2 in n1:

        return True

    return False


# ====================================================================
# LER PLANILHA
# ====================================================================

def coluna_para_numero(texto, rotulo="plano"):

    texto = str(texto or "").strip()

    if not texto:

        return None

    if texto.isdigit():

        numero = int(texto)

        if numero < 1:

            raise ValueError(f"Coluna do {rotulo} inválida: {texto}")

        return numero

    try:

        return column_index_from_string(texto.upper())

    except Exception:

        raise ValueError(f"Coluna do {rotulo} inválida: {texto}")


def cabecalhos_planilha(sheet, coluna_plano=None):

    nomes = {
        "nome",
        "cliente",
        "nome do cliente",
        "aluno",
        "nome do aluno",
        "nome do formando",
        "nome completo",
        "formando",
    }

    codigos = {
        "codigo",
        "codigo sge",
        "sge",
        "cod",
    }

    planos = {
        "plano",
        "produto",
        "album",
        "tipo",
        "tipo de album",
    }

    linhas_amostra = list(sheet.iter_rows(max_row=60, values_only=True))

    for numero, linha in enumerate(
        linhas_amostra[:20],
        start=1,
    ):

        colunas = {}
        candidatos_nome = []

        for indice, valor in enumerate(linha, start=1):

            texto = normalizar(valor)

            if texto in nomes:
                candidatos_nome.append((0 if texto == "nome" else 1, indice))

            if texto in codigos and "codigo" not in colunas:

                colunas["codigo"] = indice

            if texto in planos and "plano" not in colunas:

                colunas["plano"] = indice

        if candidatos_nome:
            melhor_ind = None
            melhor_chave = None
            for prio, ind in candidatos_nome:
                preenchidas = sum(
                    1
                    for l_amostra in linhas_amostra[numero:]
                    if ind <= len(l_amostra) and normalizar(l_amostra[ind - 1])
                )
                chave = (preenchidas > 0, -prio, preenchidas, -ind)
                if melhor_chave is None or chave > melhor_chave:
                    melhor_chave = chave
                    melhor_ind = ind
            colunas["nome"] = melhor_ind

        if "nome" in colunas and ("plano" in colunas or coluna_plano):

            if coluna_plano:

                colunas["plano"] = coluna_plano

            return numero, colunas

    return 0, {
        "nome": COLUNA_NOME,
        "codigo": COLUNA_CODIGO,
        "plano": coluna_plano or COLUNA_PRODUTO,
    }


def plano_no_filtro(filtro, plano, tamanho):

    if not filtro:

        return True

    alvo = normalizar(filtro)

    if not alvo:

        return True

    if tamanho and alvo == normalizar(tamanho):

        return True

    return alvo in normalizar(plano)


def carregar_planilha(
    caminho=None,
    aba=None,
    filtro_plano=None,
    coluna_plano=None,
    coluna_nome=None,
    coluna_codigo=None,
):

    caminho = Path(caminho) if caminho else PLANILHA

    print()
    print("=" * 80)
    print("📊 CARREGANDO PLANILHA")
    print("=" * 80)
    print(caminho)

    if aba:

        print(f"Aba: {aba}")

    if filtro_plano:

        print(f"Filtro de plano: {filtro_plano}")

    if not caminho.exists():

        print()
        print("❌ PLANILHA NÃO ENCONTRADA")
        print(caminho)
        return []

    try:

        workbook = load_workbook(caminho, data_only=True)

        if aba:

            if aba not in workbook.sheetnames:

                print()
                print(f"❌ ABA NÃO ENCONTRADA: {aba}")
                print("Abas: " + ", ".join(workbook.sheetnames))
                return []

            sheet = workbook[aba]

        else:

            sheet = workbook.active

    except Exception as erro:

        print()
        print("❌ ERRO AO ABRIR PLANILHA")
        print(erro)
        return []

    linha_cabecalho, colunas = cabecalhos_planilha(sheet, coluna_plano)

    if coluna_nome:

        colunas["nome"] = coluna_nome

    if coluna_codigo:

        colunas["codigo"] = coluna_codigo

    print(
        "Colunas: nome {nome}, código {codigo}, plano {plano}".format(
            **{chave: colunas.get(chave, "-") for chave in ("nome", "codigo", "plano")}
        )
    )

    clientes = {}

    for numero_linha, linha in enumerate(
        sheet.iter_rows(values_only=True),
        start=1,
    ):

        if numero_linha <= linha_cabecalho:

            continue

        def pegar(chave):

            indice = colunas.get(chave)

            if not indice or indice > len(linha):

                return ""

            return texto_celula(linha[indice - 1])

        nome = pegar("nome")
        codigo = pegar("codigo")
        plano = pegar("plano")

        if not nome:

            continue

        if normalizar(nome) in {"nome", "cliente", "nome do cliente", "aluno"}:

            continue

        tamanho = tamanho_do_plano(plano)

        if not plano_no_filtro(filtro_plano, plano, tamanho):

            continue

        chave = normalizar(nome)

        if chave not in clientes:

            clientes[chave] = {
                "nome": nome,
                "codigo": codigo,
                "produtos": [],
                "tamanhos": set(),
                "linhas": [],
                "itens": [],
            }

        cliente = clientes[chave]

        if codigo and not cliente["codigo"]:

            cliente["codigo"] = codigo

        if tamanho == "20X20":

            cliente["itens"] = [
                item_antigo
                for item_antigo in cliente["itens"]
                if not (item_antigo.get("sintetico") and item_antigo.get("tamanho") == "20X20")
            ]

        item = {
            "linha": numero_linha,
            "plano": plano,
            "tamanho": tamanho,
            "sintetico": False,
        }

        cliente["itens"].append(item)
        cliente["linhas"].append(numero_linha)

        if plano:

            cliente["produtos"].append(plano)

        if tamanho:

            cliente["tamanhos"].add(tamanho)

            ja_tem_20 = any(
                item_antigo.get("tamanho") == "20X20"
                for item_antigo in cliente["itens"]
            )

            if tamanho == "24X34" and not ja_tem_20:

                cliente["tamanhos"].add("20X20")
                cliente["itens"].append({
                    "linha": numero_linha,
                    "plano": plano,
                    "tamanho": "20X20",
                    "sintetico": True,
                })

    lista_clientes = list(clientes.values())

    print()

    for cliente in lista_clientes:

        print("👤 CLIENTE:")
        print(f"   {cliente['nome']}")
        print(f"   Código: {cliente['codigo']}")
        print(f"   Produtos: {', '.join(cliente['produtos'])}")
        print(
            "   Álbuns que serão procurados: "
            + ", ".join(sorted(tamanho for tamanho in cliente["tamanhos"]))
        )
        print()

    print("=" * 80)
    print(f"✅ CLIENTES ÚNICOS: {len(lista_clientes)}")
    print("=" * 80)

    return lista_clientes


# ====================================================================
# ENCONTRAR ÁLBUNS COLAÇÃO
# ====================================================================

def encontrar_albuns_colacao(caminho=None):

    # Usa exatamente a pasta indicada, sem acrescentar outra por cima.
    caminho = Path(caminho) if caminho else SERVIDOR

    print()
    print("=" * 80)

    print(
        "🌐 VERIFICANDO SERVIDOR"
    )

    print("=" * 80)

    print(
        caminho
    )

    if not caminho.exists():

        print()
        print(
            "❌ PASTA NÃO ENCONTRADA"
        )

        return None

    print()
    print(
        "✅ SERVIDOR ACESSÍVEL"
    )

    return caminho


# ====================================================================
# ENCONTRAR TODAS AS REMESSAS
# ====================================================================

def encontrar_remessas(
    albuns_colacao
):

    remessas = []

    print()
    print("=" * 80)

    print(
        "📁 PROCURANDO REMESSAS E PASTAS PRONTOS"
    )

    print("=" * 80)

    try:

        for pasta in albuns_colacao.iterdir():

            if not pasta.is_dir():

                continue

            nome = normalizar(
                pasta.name
            )

            if "remessa" in nome or "prontos" in nome:

                remessas.append(
                    pasta
                )

    except Exception as erro:

        print(
            "❌ ERRO:"
        )

        print(
            erro
        )

        return []

    # --------------------------------------------------------------
    # ORDENAR REMESSAS
    # --------------------------------------------------------------

    def numero_remessa(pasta):

        numeros = re.findall(
            r"\d+",
            pasta.name
        )

        if numeros:

            return int(
                numeros[0]
            )

        return 999999

    if not remessas:

        print()
        print("Nenhuma pasta de remessa ou de prontos.")
        print("Procurando direto na pasta indicada:")
        print(albuns_colacao)
        remessas.append(albuns_colacao)

    remessas.sort(
        key=numero_remessa
    )

    print()

    for remessa in remessas:

        print(
            f"📁 {remessa.name}"
        )

    print()

    print(
        f"TOTAL: {len(remessas)}"
    )

    return remessas


# ====================================================================
# ENCONTRAR #ÁLBUNS PRONTOS
# ====================================================================

def encontrar_pastas_prontos(
    remessa
):

    encontrados = []

    try:

        for pasta in remessa.rglob("*"):

            if not pasta.is_dir():

                continue

            if normalizar(
                pasta.name
            ) == normalizar(
                PASTA_PRONTOS
            ):

                encontrados.append(
                    pasta
                )

    except Exception:

        pass

    return encontrados


# ====================================================================
# ENCONTRAR PASTA DA PESSOA
# ====================================================================

def encontrar_cliente(
    remessa,
    nome_cliente
):

    def pontuacao_pasta(pasta):

        return pontuacao_nome(
            nome_cliente,
            pasta.name
        )

    # ==============================================================
    # PRIMEIRO:
    # #ÁLBUNS PRONTOS
    # ==============================================================

    pastas_prontos = encontrar_pastas_prontos(
        remessa
    )

    candidatos_prontos = []

    for pasta_prontos in pastas_prontos:

        try:

            for pasta in pasta_prontos.iterdir():

                if not pasta.is_dir():

                    continue

                pontuacao = pontuacao_pasta(pasta)

                if pontuacao >= 0:

                    candidatos_prontos.append(
                        (pontuacao, pasta)
                    )

        except Exception:

            pass

    # --------------------------------------------------------------
    # SE ENCONTROU EM PRONTOS
    # --------------------------------------------------------------

    if candidatos_prontos:

        # Se houver mais de uma,
        # usa a primeira encontrada.

        candidatos_prontos.sort(
            key=lambda item: (-item[0], len(item[1].parts))
        )

        return candidatos_prontos[0][1]

    # ==============================================================
    # SEGUNDO:
    # PROCURAR NA REMESSA
    # ==============================================================

    candidatos = []

    try:

        for pasta in remessa.rglob("*"):

            if not pasta.is_dir():

                continue

            pontuacao = pontuacao_pasta(pasta)

            if pontuacao >= 0:

                candidatos.append(
                    (pontuacao, pasta)
                )

    except Exception:

        pass

    if candidatos:

        candidatos.sort(
            key=lambda item: (-item[0], len(item[1].parts))
        )

        return candidatos[0][1]

    return None


# ====================================================================
# ENCONTRAR ÁLBUM
# ====================================================================

def encontrar_album(
    pasta_cliente,
    tamanho
):

    if not pasta_cliente or tamanho not in QUANTIDADE_LAMINAS:

        return None

    candidatos = []

    try:

        for pasta in pasta_cliente.rglob("*"):

            if not pasta.is_dir():

                continue

            if tamanho not in tamanhos_mencionados(pasta.name):

                continue

            if normalizar(pasta.name) == "jpeg":

                continue

            candidatos.append(pasta)

    except Exception:

        pass

    if not candidatos:

        return None

    # Pasta do tamanho mais direta (Álbum 20x20, Álbum 34x24, Álbum 30x20).
    candidatos.sort(
        key=lambda pasta: (
            len(tamanhos_mencionados(pasta.name)) != 1,
            len(pasta.parts)
        )
    )

    return candidatos[0]


# ====================================================================
# ENCONTRAR JPEG
# ====================================================================

def encontrar_jpeg(
    pasta_album
):

    try:

        # Primeiro procura JPEG diretamente

        for pasta in pasta_album.iterdir():

            if not pasta.is_dir():

                continue

            if normalizar(
                pasta.name
            ) == "jpeg":

                return pasta

        # Se não estiver diretamente,
        # procura dentro

        for pasta in pasta_album.rglob("*"):

            if not pasta.is_dir():

                continue

            if normalizar(
                pasta.name
            ) == "jpeg":

                return pasta

    except Exception:

        pass

    return None


# ====================================================================
# IDENTIFICAR LÂMINA
# ====================================================================

def eh_lamina(
    arquivo
):

    if not arquivo.is_file():

        return False

    if arquivo.suffix.lower() not in {

        ".jpg",
        ".jpeg"

    }:

        return False

    # A pasta JPEG já pertence à pessoa e ao álbum.
    # Portanto, o nome do arquivo não precisa seguir um padrão.
    return True


def eh_nome_capa(nome):

    nome = normalizar(nome)

    return nome == "capa" or nome.startswith("capa ")


def eh_arquivo_extra(arquivo):

    """Capa, Thumbs.db e arquivos do mesmo tipo não entram na contagem das lâminas."""

    nome = arquivo.name.lower()

    if nome in {
        "thumbs.db",
        "ehthumbs.db",
        "desktop.ini",
        ".ds_store",
    }:

        return True

    stem = normalizar(arquivo.stem)

    if (
        stem in {"thumbs", "ehthumbs", "desktop", "folder"}
        or stem.startswith("albumart")
        or "thumb" in stem
    ):

        return True

    return eh_nome_capa(arquivo.stem)


# ====================================================================
# NÚMERO DA LÂMINA
# ====================================================================

def numero_lamina(
    arquivo
):

    encontrado = re.search(
        r"\(\s*(\d+)\s*\)",
        arquivo.stem
    )

    if encontrado:

        return int(
            encontrado.group(1)
        )

    encontrado = re.search(
        r"^\s*(\d+)",
        arquivo.stem
    )

    if encontrado:

        return int(
            encontrado.group(1)
        )

    return 999999


# ====================================================================
# ENCONTRAR LÂMINAS
# ====================================================================

def encontrar_laminas(
    pasta_jpeg,
    quantidade
):

    laminas = []

    try:

        for arquivo in pasta_jpeg.rglob("*"):

            if eh_lamina(
                arquivo
            ):

                laminas.append(
                    arquivo
                )

    except Exception as erro:

        print(
            f"❌ ERRO LENDO JPEG: {erro}"
        )

        return []

    # --------------------------------------------------------------
    # ORDENAR
    # --------------------------------------------------------------

    laminas.sort(
        key=numero_lamina
    )

    # --------------------------------------------------------------
    # RETORNAR QUANTIDADE CORRETA
    # --------------------------------------------------------------

    return laminas[
        :quantidade
    ]


# ====================================================================
# PROCURAR CAPA NA PASTA GERAL
# ====================================================================

# JPG da capa pronta e também PNG: no lote há foto 20x20 só em PNG.
EXTENSOES_CAPA = {".jpg", ".jpeg", ".png"}

# Proporção largura/altura vista nas capas e nas fotos de capa.
# 24x34 é retrato (~2:3). 20x20 é quase quadrado ou a capa aberta (~2:1).
# 30x20 é a foto do molde (~1,3:1) ou a capa aberta (~2,7:1).
FAIXAS_PROPORCAO_CAPA = (
    ("24X34", 0.58, 0.78),
    ("20X20", 0.84, 1.12),
    ("30X20", 1.18, 1.55),
    ("20X20", 1.75, 2.20),
    ("30X20", 2.40, 3.05),
)

PALAVRAS_FRACAS_NOME = {
    "da",
    "de",
    "do",
    "das",
    "dos",
    "e",
    "di",
    "du",
    "del",
}


def listar_arquivos_capas(pasta_capas):

    cache = getattr(listar_arquivos_capas, "cache", None)

    if cache and cache[0] == pasta_capas:

        return cache[1]

    arquivos = []

    try:

        for arquivo in pasta_capas.rglob("*"):

            if not arquivo.is_file():

                continue

            if arquivo.name.startswith("._"):

                continue

            if arquivo.suffix.lower() not in EXTENSOES_CAPA:

                continue

            arquivos.append(arquivo)

    except Exception as erro:

        print(f"⚠ ERRO LENDO CAPAS: {erro}")

    listar_arquivos_capas.cache = (pasta_capas, arquivos)

    return arquivos


def dimensoes_imagem(caminho):

    """Lê só o cabeçalho. Não abre a imagem inteira."""

    try:

        with open(caminho, "rb") as arquivo:

            inicio = arquivo.read(8)

            if inicio.startswith(b"\x89PNG\r\n\x1a\n"):

                arquivo.seek(16)
                dados = arquivo.read(8)

                if len(dados) < 8:

                    return None

                largura, altura = struct.unpack(">II", dados)

                if largura and altura:

                    return largura, altura

                return None

            if not inicio.startswith(b"\xff\xd8"):

                return None

            arquivo.seek(2)
            lido = 2
            limite = 8 * 1024 * 1024

            while lido < limite:

                byte = arquivo.read(1)

                if not byte:

                    return None

                lido += 1

                if byte != b"\xff":

                    return None

                while True:

                    tipo_b = arquivo.read(1)

                    if not tipo_b:

                        return None

                    lido += 1

                    if tipo_b != b"\xff":

                        break

                tipo = tipo_b[0]

                if tipo in (0xD8, 0xD9) or tipo == 0x01 or 0xD0 <= tipo <= 0xD7:

                    continue

                tamanho_b = arquivo.read(2)

                if len(tamanho_b) < 2:

                    return None

                lido += 2
                bloco = struct.unpack(">H", tamanho_b)[0]

                if bloco < 2:

                    return None

                if tipo in (
                    0xC0, 0xC1, 0xC2, 0xC3,
                    0xC5, 0xC6, 0xC7,
                    0xC9, 0xCA, 0xCB,
                    0xCD, 0xCE, 0xCF,
                ):

                    corpo = arquivo.read(5)

                    if len(corpo) < 5:

                        return None

                    altura, largura = struct.unpack(">HH", corpo[1:5])

                    if largura and altura:

                        return largura, altura

                    return None

                if tipo == 0xDA:

                    return None

                resto = bloco - 2
                arquivo.seek(resto, 1)
                lido += resto

    except Exception:

        return None

    return None


def tamanho_pela_proporcao(caminho):

    dimensoes = dimensoes_imagem(caminho)

    if not dimensoes:

        return None

    largura, altura = dimensoes

    if not largura or not altura:

        return None

    proporcao = largura / altura

    for tamanho, minimo, maximo in FAIXAS_PROPORCAO_CAPA:

        if minimo <= proporcao <= maximo:

            return tamanho

    return None


def nomes_no_caminho(arquivo, limite):

    nomes = [arquivo.stem]

    for pai in arquivo.parents:

        if pai == limite:

            break

        nomes.append(pai.name)

    return nomes


def capa_da_pessoa(nome_cliente, nomes):

    """O nome inteiro precisa bater. 'Maria' + 'Lima' de outra pessoa não serve."""

    alvo = normalizar(nome_cliente)

    if not alvo:

        return False

    texto = normalizar(" ".join(nomes))

    if alvo in texto:

        return True

    significativas = [
        palavra
        for palavra in alvo.split()
        if palavra not in PALAVRAS_FRACAS_NOME and not palavra.isdigit()
    ]

    if len(significativas) < 2:

        return False

    palavras = set(texto.split())

    return all(palavra in palavras for palavra in significativas)


def tamanho_da_capa(arquivo, limite):

    """
    Tamanho da capa e a força da evidência.

    3 = o arquivo diz o tamanho (20X20, 24X34, 30X20)
    2 = só a pasta diz, e a proporção não contradiz
    1 = a proporção da imagem
    0 = não dá para afirmar

    Pasta 'capas 24X34 20X20' cita os dois tamanhos e não decide sozinha.
    Se o nome diz um tamanho e a foto tem a proporção de outro, vale a foto.
    """

    do_arquivo = tamanhos_mencionados(arquivo.stem)
    proporcao = None

    if len(do_arquivo) == 1:

        indicado = next(iter(do_arquivo))
        proporcao = tamanho_pela_proporcao(arquivo)

        if proporcao and proporcao != indicado:

            return proporcao, 1

        return indicado, 3

    mencoes_pasta = tamanhos_mencionados(limite.name) if limite is not None else set()

    for nome in nomes_no_caminho(arquivo, limite)[1:]:

        mencoes_pasta |= tamanhos_mencionados(nome)

    pasta_unica = next(iter(mencoes_pasta)) if len(mencoes_pasta) == 1 else None

    if proporcao is None:

        proporcao = tamanho_pela_proporcao(arquivo)

    if pasta_unica and (proporcao is None or proporcao == pasta_unica):

        return pasta_unica, 2

    if proporcao:

        return proporcao, 1

    return None, 0


# Foto de capa no molde (pixels). Empate de nome fica com esta, não com a outra exportação.
MOLDES_FOTO_CAPA = {
    "20X20": (2694, 3012),
    "24X34": (4160, 6240),
    "30X20": (3874, 3012),
}


def penalidade_copia(arquivo):

    if re.search(r"\(\s*\d+\s*\)", arquivo.stem):

        return 1

    return 0


def preferencia_foto(arquivo, tamanho):

    molde = MOLDES_FOTO_CAPA.get(tamanho)

    if not molde:

        return 0

    dimensoes = dimensoes_imagem(arquivo)

    if not dimensoes:

        return 0

    tolerancia = 8

    if (
        abs(dimensoes[0] - molde[0]) <= tolerancia
        and abs(dimensoes[1] - molde[1]) <= tolerancia
    ):

        return 1

    return 0


def escolher_melhor_capa(candidatos, tamanho=None):

    """
    Não usa a ordem da pasta.
    1. evidência de tamanho (nome do arquivo, pasta, proporção)
    2. nome da pessoa
    3. foto no molde do tamanho
    4. sem sufixo de cópia (2)
    5. caminho mais curto e, por último, ordem alfabética
    """

    if not candidatos:

        return None

    def chave(item):

        forca, pontuacao, arquivo = item

        return (
            -forca,
            -pontuacao,
            -preferencia_foto(arquivo, tamanho),
            penalidade_copia(arquivo),
            len(arquivo.parts),
            str(arquivo).lower(),
        )

    candidatos.sort(key=chave)

    return candidatos[0][2]


def procurar_capa_geral(
    albuns_colacao,
    nome_cliente,
    tamanho
):

    pasta_capas = (
        albuns_colacao
        / PASTA_CAPAS
    )

    if not pasta_capas.exists():

        return None

    candidatos = []

    try:

        for arquivo in listar_arquivos_capas(pasta_capas):

            nomes = nomes_no_caminho(arquivo, pasta_capas)

            if not capa_da_pessoa(nome_cliente, nomes):

                continue

            pontuacao = max(
                pontuacao_nome(nome_cliente, nome)
                for nome in nomes
            )

            pontuacao = max(
                pontuacao,
                pontuacao_nome(nome_cliente, " ".join(nomes)),
            )

            if pontuacao < 0:

                continue

            detectado, forca = tamanho_da_capa(
                arquivo,
                pasta_capas
            )

            # Sem o tamanho pedido, não usa capa de outro tamanho.
            if detectado != tamanho or forca <= 0:

                continue

            candidatos.append(
                (forca, pontuacao, arquivo)
            )

    except Exception as erro:

        print(
            f"⚠ ERRO PROCURANDO CAPA GERAL: {erro}"
        )

    return escolher_melhor_capa(candidatos, tamanho)


# ====================================================================
# PROCURAR CAPA ANTIGA
# ====================================================================

def procurar_capa_antiga(
    pasta_cliente,
    tamanho
):

    if not pasta_cliente:

        return None

    candidatos = []

    try:

        for arquivo in pasta_cliente.rglob("*"):

            if not arquivo.is_file():

                continue

            if arquivo.name.startswith("._"):

                continue

            if arquivo.suffix.lower() not in EXTENSOES_CAPA:

                continue

            # ======================================================
            # CAPAS ANTIGAS
            # ======================================================
            #
            # capa.jpg
            # capa final.jpg
            # capa-final.jpg
            # capa_20x20.jpg
            # capa 24x34.jpg
            #
            # ======================================================

            if not eh_nome_capa(arquivo.stem):

                continue

            detectado, forca = tamanho_da_capa(
                arquivo,
                pasta_cliente
            )

            # Capa sem tamanho, ou de outro tamanho, não entra no lugar.
            if detectado != tamanho or forca <= 0:

                continue

            candidatos.append(
                (forca, 0, arquivo)
            )

    except Exception as erro:

        print(
            f"⚠ ERRO PROCURANDO CAPA ANTIGA: {erro}"
        )

    return escolher_melhor_capa(candidatos, tamanho)


# ====================================================================
# PROCURAR CAPA
# ====================================================================

def encontrar_capa(
    albuns_colacao,
    nome_cliente,
    tamanho,
    pasta_cliente
):

    print()
    print(
        "🖼 Procurando capa..."
    )

    # ==============================================================
    # 1º - CAPA NA PASTA GERAL
    # ==============================================================

    capa = procurar_capa_geral(
        albuns_colacao,
        nome_cliente,
        tamanho
    )

    if capa:

        print()
        print(
            "✅ CAPA ENCONTRADA NA PASTA GERAL:"
        )

        print(
            f"   {capa}"
        )

        return capa

    # ==============================================================
    # 2º - CAPA ANTIGA NA PASTA DA PESSOA
    # ==============================================================

    print(
        "⚠ Capa geral não encontrada."
    )

    print(
        "🔎 Procurando formato antigo..."
    )

    capa = procurar_capa_antiga(
        pasta_cliente,
        tamanho
    )

    if capa:

        print()
        print(
            "✅ CAPA ANTIGA ENCONTRADA:"
        )

        print(
            f"   {capa}"
        )

        return capa

    print()
    print(
        "❌ CAPA NÃO ENCONTRADA"
    )

    return None


# ====================================================================
# CRIAR DESTINO
# ====================================================================

def nome_pasta_pessoa(cliente):

    nome = re.sub(
        r'[<>:"/\\|?*]',
        "",
        str(cliente.get("nome") or ""),
    ).strip().rstrip(".")

    return nome or "sem-nome"


# ====================================================================
# COPIAR ARQUIVO
# ====================================================================

def copiar_arquivo(
    origem,
    destino
):

    destino.parent.mkdir(
        parents=True,
        exist_ok=True
    )

    # --------------------------------------------------------------
    # NÃO SOBRESCREVER
    # --------------------------------------------------------------

    if destino.exists():

        contador = 2

        while True:

            novo_nome = (
                f"{destino.stem}_{contador}"
                f"{destino.suffix}"
            )

            novo_destino = (
                destino.parent
                /
                novo_nome
            )

            if not novo_destino.exists():

                destino = novo_destino

                break

            contador += 1

    shutil.copy2(
        origem,
        destino
    )

    return destino


# ====================================================================
# PROCESSAR UM TAMANHO
# ====================================================================

def pastas_do_tamanho(tamanho):

    base = DESTINO / tamanho
    pastas = [base / "CAPAS", base / "LAMINAS"]

    if tamanho in {"24X34", "30X20"}:

        pastas.append(base / "ESTOJO")

    return pastas


def caminhos_destino(cliente, tamanho):

    nome = nome_pasta_pessoa(cliente)
    base = DESTINO / tamanho

    return base / "LAMINAS" / nome, base / "CAPAS"


def garantir_pastas(cliente, tamanhos):

    if not tamanhos:

        return

    nome = nome_pasta_pessoa(cliente)

    print()
    print("📁 DESTINO:")

    for tamanho in tamanhos:

        for pasta in pastas_do_tamanho(tamanho):

            print(f"   {pasta}")

        print(f"   {DESTINO / tamanho / 'LAMINAS' / nome}")
        print(f"   {DESTINO / tamanho / 'CAPAS' / (nome + '.jpg')}")

        if MODO_TESTE:

            print("   🔵 TESTE - PASTA NÃO CRIADA")

        else:

            for pasta in pastas_do_tamanho(tamanho):

                pasta.mkdir(parents=True, exist_ok=True)

            print("   ✅ PASTA CRIADA")


def itens_por_tamanho(cliente):

    sem_tamanho = []
    escolhidos = {}

    for item in cliente.get("itens", []):

        tamanho = item.get("tamanho")

        if not tamanho:

            sem_tamanho.append(item)
            continue

        atual = escolhidos.get(tamanho)

        if atual is None or (atual.get("sintetico") and not item.get("sintetico")):

            escolhidos[tamanho] = item

    ordem = ["20X20", "24X34", "30X20"]

    tamanhos = sorted(
        escolhidos,
        key=lambda tamanho: ordem.index(tamanho) if tamanho in ordem else 99,
    )

    return sem_tamanho, [escolhidos[tamanho] for tamanho in tamanhos]


def tamanhos_nos_jpg(pasta):

    mencoes = set()

    try:

        arquivos = list(pasta.iterdir())

    except Exception:

        return mencoes

    for arquivo in arquivos:

        if arquivo.is_file() and arquivo.suffix.lower() in {".jpg", ".jpeg"}:

            mencoes |= tamanhos_mencionados(arquivo.name)

    return mencoes


def indexar_pastas_album(remessas):

    print()
    print("📁 LENDO PASTAS DE ÁLBUM NO CAMINHO INDICADO")

    indice = []

    for remessa in remessas:

        try:

            pastas = list(remessa.rglob("*"))

        except Exception as erro:

            print(f"⚠ ERRO LENDO {remessa}: {erro}")
            continue

        for pasta in pastas:

            if not pasta.is_dir():

                continue

            mencoes = tamanhos_mencionados(pasta.name)
            mencoes |= tamanhos_nos_jpg(pasta)

            if not mencoes:

                continue

            indice.append((remessa, pasta, mencoes))

    print(f"   Pastas com tamanho no nome: {len(indice)}")

    return indice


def pontuacao_no_caminho(nome, pasta, limite):

    melhor = -1
    atual = pasta

    while atual != limite and atual != atual.parent:

        melhor = max(melhor, pontuacao_nome(nome, atual.name))
        atual = atual.parent

    return melhor


def separar_laminas_e_extras(arquivos, quantidade):

    """O padrão das lâminas vale com até 2 arquivos a mais: capa, Thumbs.db ou parecido."""

    extras = []
    laminas = []

    for arquivo in arquivos:

        if eh_arquivo_extra(arquivo):

            extras.append(arquivo)

        elif eh_lamina(arquivo):

            laminas.append(arquivo)

    if quantidade < len(laminas) <= quantidade + 2:

        sem_numero = [arquivo for arquivo in laminas if numero_lamina(arquivo) == 999999]
        sobra = len(laminas) - quantidade

        if sem_numero and len(extras) + sobra <= 2 and len(sem_numero) >= sobra:

            for arquivo in sem_numero[:sobra]:

                laminas.remove(arquivo)
                extras.append(arquivo)

    laminas = sorted(laminas, key=numero_lamina)

    if len(laminas) > quantidade and len(laminas) - quantidade <= 2 and len(extras) <= 2:

        laminas = laminas[:quantidade]

    return laminas, extras


def localizar_laminas(pasta_album, quantidade):

    """Acha os JPG pelo padrão de lâmina, sem exigir uma pasta chamada JPEG ou JPG."""

    def arquivos_diretos(pasta):

        try:

            return [arquivo for arquivo in pasta.iterdir() if arquivo.is_file()]

        except Exception:

            return []

    opcoes = []

    def acrescentar(pasta, bonus):

        arquivos = arquivos_diretos(pasta)

        if not arquivos:

            return

        laminas, extras = separar_laminas_e_extras(arquivos, quantidade)

        if not laminas and not extras:

            return

        opcoes.append((laminas, extras, pasta, bonus))

    acrescentar(pasta_album, 0)

    try:

        subpastas = [pasta for pasta in pasta_album.rglob("*") if pasta.is_dir()]

    except Exception:

        subpastas = []

    for sub in subpastas:

        nome = normalizar(sub.name)
        bonus = 1 if nome in {"jpeg", "jpg", "fotos", "foto", "laminas", "lamina"} else 0
        acrescentar(sub, bonus)

    if not opcoes:

        return None, [], []

    def chave(item):

        laminas, extras, pasta, bonus = item
        qtd = len(laminas)
        completa = qtd == quantidade and len(extras) <= 2

        return (
            completa,
            -abs(qtd - quantidade),
            bonus,
            -len(extras),
            -len(pasta.parts),
        )

    laminas, extras, pasta, _ = max(opcoes, key=chave)

    return pasta, laminas, extras


def escolher_album(indice, nome, tamanho, quantidade):

    """Entre todas as pastas do tamanho, fica com a da pessoa certa, não com a primeira."""

    melhor = None

    for remessa, pasta, mencoes in indice:

        if tamanho not in mencoes:

            continue

        score = pontuacao_no_caminho(nome, pasta, remessa)

        if score < 0:

            continue

        pasta_laminas, laminas, extras = localizar_laminas(pasta, quantidade)
        qtd = len(laminas)
        chave = (
            score,
            len(mencoes) == 1,
            qtd == quantidade and len(extras) <= 2,
            qtd > 0,
            -abs(qtd - quantidade),
            -len(pasta.parts),
        )

        if melhor is None or chave > melhor[0]:

            melhor = (chave, remessa, pasta, pasta_laminas, laminas, extras)

    if not melhor:

        return None

    _, remessa, pasta, pasta_laminas, laminas, extras = melhor

    return {
        "remessa": remessa,
        "pasta_album": pasta,
        "pasta_laminas": pasta_laminas,
        "laminas": laminas,
        "extras": extras,
    }


def processar_tamanho(
    cliente,
    tamanho,
    albuns_colacao,
    indice,
):

    nome = cliente["nome"]
    quantidade = QUANTIDADE_LAMINAS[tamanho]
    encontrado = False
    pasta_laminas = None
    laminas = []
    remessa_nome = ""

    print()
    print("-" * 80)
    print(f"📦 PROCESSANDO ÁLBUM {tamanho}")
    print(f"👤 {nome}")
    print(f"📄 Lâminas esperadas: {quantidade}")
    print("-" * 80)

    escolha = escolher_album(indice, nome, tamanho, quantidade)
    pasta_album = escolha["pasta_album"] if escolha else None

    if not pasta_album:

        print()
        print(f"❌ ÁLBUM {tamanho} NÃO ENCONTRADO")

    else:

        encontrado = True
        remessa_nome = escolha["remessa"].name
        pasta_laminas = escolha["pasta_laminas"]
        laminas = escolha["laminas"]

        print()
        print("✅ ÁLBUM ENCONTRADO:")
        print(f"   {pasta_album}")
        print(f"   Remessa: {remessa_nome}")

        if not laminas:

            print()
            print("❌ LÂMINAS NÃO ENCONTRADAS")

        else:

            print()
            print("✅ LÂMINAS ENCONTRADAS:")
            print(f"   {pasta_laminas}")
            print(f"📄 LÂMINAS: {len(laminas)}/{quantidade}")

            if len(laminas) == quantidade:

                print("✅ QUANTIDADE CORRETA")

            else:

                print("⚠ FALTAM LÂMINAS")

        extras = escolha.get("extras") or []

        if extras:

            print(f"   Arquivos extras ignorados: {', '.join(arquivo.name for arquivo in extras)}")

    capa = encontrar_capa(
        albuns_colacao,
        nome,
        tamanho,
        pasta_album,
    )

    if not capa and escolha:

        extras_capa = []

        for arquivo in (escolha.get("extras") or []):

            if arquivo.name.startswith("._"):

                continue

            if arquivo.suffix.lower() not in EXTENSOES_CAPA:

                continue

            if not eh_nome_capa(arquivo.stem):

                continue

            detectado, forca = tamanho_da_capa(
                arquivo,
                escolha["pasta_album"],
            )

            if detectado != tamanho or forca <= 0:

                continue

            extras_capa.append((forca, 0, arquivo))

        capa = escolher_melhor_capa(extras_capa, tamanho)

        if capa:

            print()
            print("✅ CAPA ENCONTRADA JUNTO DAS LÂMINAS:")
            print(f"   {capa}")

    pasta_laminas, pasta_capas = caminhos_destino(cliente, tamanho)
    nome_pessoa = nome_pasta_pessoa(cliente)

    for lamina in laminas:

        destino = pasta_laminas / lamina.name

        print()
        print(f"📄 {lamina.name}")
        print(f"   → {destino}")

        if MODO_TESTE:

            print("   🔵 TESTE - NÃO COPIADO")

        else:

            try:

                copiar_arquivo(lamina, destino)
                print("   ✅ COPIADO")

            except Exception as erro:

                print(f"   ❌ ERRO: {erro}")

    if capa:

        destino = pasta_capas / f"{nome_pessoa}{capa.suffix.lower()}"

        print()
        print("🖼 CAPA")
        print(f"   {capa}")
        print(f"   → {destino}")

        if MODO_TESTE:

            print("   🔵 TESTE - NÃO COPIADO")

        else:

            try:

                copiar_arquivo(capa, destino)
                print("   ✅ CAPA COPIADA")

            except Exception as erro:

                print(f"   ❌ ERRO: {erro}")

    partes = []

    if not encontrado:

        partes.append("álbum não encontrado")

    elif not laminas:

        partes.append("lâminas não encontradas")

    elif len(laminas) < quantidade:

        partes.append(f"faltam lâminas ({len(laminas)}/{quantidade})")

    else:

        partes.append(f"lâminas {len(laminas)}/{quantidade}")

    partes.append("capa encontrada" if capa else "capa não encontrada")

    if MODO_TESTE:

        partes.append("teste, nada copiado")

    return {
        "encontrado": encontrado,
        "laminas": len(laminas),
        "esperadas": quantidade,
        "capa": capa is not None,
        "remessa": remessa_nome,
        "mensagem": "; ".join(partes),
    }


def status_do_resultado(info):

    if not info["encontrado"]:

        return "NAO ENCONTRADO"

    if info["laminas"] < info["esperadas"]:

        return "FALTAM LAMINAS"

    if not info["capa"]:

        return "SEM CAPA"

    return "OK"


def linha_relatorio(cliente, item, remessa, status, laminas, capa, mensagem):

    if item.get("sintetico"):

        mensagem = (
            "20X20 incluído porque o plano 24X34 também leva o álbum 20X20. "
            + mensagem
        )

    return {
        "STATUS": status,
        "LINHA": item.get("linha", ""),
        "NOME": cliente.get("nome", ""),
        "CODIGO": cliente.get("codigo", ""),
        "PLANO": item.get("plano", ""),
        "TAMANHO": item.get("tamanho") or "",
        "REMESSA": remessa or "",
        "LAMINAS": laminas,
        "CAPA": capa,
        "MENSAGEM": mensagem,
    }


def processar_cliente(cliente, albuns_colacao, indice):

    print()
    print()
    print("#" * 80)
    print(f"👤 CLIENTE: {cliente['nome']}")
    print(f"🔢 CÓDIGO: {cliente['codigo']}")
    print(f"📦 PRODUTOS: {', '.join(cliente.get('produtos') or [])}")
    print("📁 A pasta do álbum é escolhida entre todas as remessas, pelo nome e pelo tamanho.")
    print("#" * 80)

    sem_tamanho, itens = itens_por_tamanho(cliente)
    garantir_pastas(cliente, [item["tamanho"] for item in itens])

    linhas = []

    for item in sem_tamanho:

        print()
        print(f"⚠ PLANO NÃO IDENTIFICADO: {item.get('plano') or '(vazio)'} (linha {item.get('linha')})")

        linhas.append(linha_relatorio(
            cliente,
            item,
            "",
            "PLANO NAO IDENTIFICADO",
            "",
            "",
            "não deu para saber o tamanho do álbum a partir do plano",
        ))

    for item in itens:

        try:

            info = processar_tamanho(
                cliente,
                item["tamanho"],
                albuns_colacao,
                indice,
            )

            linhas.append(linha_relatorio(
                cliente,
                item,
                info.get("remessa", ""),
                status_do_resultado(info),
                f"{info['laminas']}/{info['esperadas']}",
                "SIM" if info["capa"] else "NAO",
                info["mensagem"],
            ))

        except Exception as erro:

            print()
            print("❌ ERRO PROCESSANDO ÁLBUM")
            print(erro)

            linhas.append(linha_relatorio(
                cliente,
                item,
                "",
                "ERRO",
                "",
                "",
                str(erro),
            ))

    return linhas


def modulo_renomeio():

    for pasta in PASTA_SCRIPT.parent.iterdir():

        if not pasta.is_dir():

            continue

        achados = list(pasta.glob("PASTA RENOMEIO.py"))

        if achados:

            return SourceFileLoader("renomeio_lote", str(achados[0])).load_module()

    raise FileNotFoundError("PASTA RENOMEIO.py não encontrado")


def houve_copia(linha):

    if linha.get("CAPA") == "SIM":

        return True

    texto = str(linha.get("LAMINAS") or "")

    return "/" in texto and not texto.startswith("0/")


def renomear_lote(clientes, linhas):

    print()
    print("=" * 80)
    print("       RENOMEANDO PELO CÓDIGO SGE")
    print("=" * 80)

    renomeio = modulo_renomeio()
    renomeio.MODO_TESTE = False
    renomeio.PAUSAR_NO_FIM = False
    renomeio.NAO_ENCONTRADOS.clear()

    for chave in renomeio.CONTADORES:

        renomeio.CONTADORES[chave] = 0

    mapa = {}
    sem_codigo = set()

    for cliente in clientes:

        chave = renomeio.normalizar(cliente.get("nome"))
        codigo = renomeio.pegar_codigo(cliente.get("codigo"))

        if codigo:

            mapa[chave] = codigo

        elif chave:

            sem_codigo.add(chave)

    for tamanho in ("20X20", "24X34", "30X20"):

        capas = DESTINO / tamanho / "CAPAS"

        if capas.is_dir():

            with os.scandir(capas) as itens:

                for item in itens:

                    if item.is_file():

                        renomeio.processar_arquivo_individual(item, mapa, tamanho, "CAPAS")

        laminas = DESTINO / tamanho / "LAMINAS"

        if laminas.is_dir() and any(laminas.iterdir()):

            renomeio.processar_laminas(str(laminas), mapa, tamanho)

    falhas = [renomeio.normalizar(item) for item in renomeio.NAO_ENCONTRADOS]

    for linha in linhas:

        if not houve_copia(linha):

            continue

        chave = renomeio.normalizar(linha.get("NOME"))

        if chave in sem_codigo or any(chave and chave in item for item in falhas):

            linha["STATUS"] = "ERRO"
            mensagem = linha.get("MENSAGEM") or ""
            linha["MENSAGEM"] = (mensagem + "; sem código SGE para renomear").strip("; ")

    print()
    print(
        "RENOMEIO — "
        f"RENOMEADOS: {renomeio.CONTADORES['renomeados']} | "
        f"JA OK: {renomeio.CONTADORES['ja_ok']} | "
        f"NAO ENCONTRADOS: {renomeio.CONTADORES['nao_encontrados']} | "
        f"ERROS: {renomeio.CONTADORES['erros']}"
    )


def salvar_relatorio(linhas):

    campos = [
        "STATUS",
        "LINHA",
        "NOME",
        "CODIGO",
        "PLANO",
        "TAMANHO",
        "REMESSA",
        "LAMINAS",
        "CAPA",
        "MENSAGEM",
    ]

    with RELATORIO.open("w", newline="", encoding="utf-8-sig") as arquivo:

        escritor = csv.DictWriter(arquivo, fieldnames=campos, delimiter=";")
        escritor.writeheader()

        for linha in linhas:

            escritor.writerow({campo: linha.get(campo, "") for campo in campos})

    print()
    print("RELATORIO:", RELATORIO)


def imprimir_resumo(linhas, remessas, clientes):

    contagem = {}

    for linha in linhas:

        status = linha.get("STATUS") or ""
        contagem[status] = contagem.get(status, 0) + 1

    print()
    print()
    print("=" * 80)
    print("                    RESUMO FINAL")
    print("=" * 80)
    print()
    print(f"📁 Remessas analisadas: {len(remessas)}")
    print(f"👤 Clientes na planilha: {len(clientes)}")
    print(f"✅ OK: {contagem.get('OK', 0)}")
    print(f"⚠ Faltam lâminas: {contagem.get('FALTAM LAMINAS', 0)}")
    print(f"🖼 Sem capa: {contagem.get('SEM CAPA', 0)}")
    print(f"❌ Não encontrados: {contagem.get('NAO ENCONTRADO', 0)}")
    print(f"⚠ Plano não identificado: {contagem.get('PLANO NAO IDENTIFICADO', 0)}")
    print(f"❌ Erros: {contagem.get('ERRO', 0)}")
    print()

    if MODO_TESTE:

        print("🔵 TESTE FINALIZADO")
        print("NENHUM ARQUIVO FOI COPIADO E NENHUMA PASTA FOI CRIADA.")

    else:

        print("✅ PROCESSAMENTO FINALIZADO")

    print()
    print(
        "RESUMO ORGANIZADOR — "
        f"OK: {contagem.get('OK', 0)} | "
        f"FALTAM LAMINAS: {contagem.get('FALTAM LAMINAS', 0)} | "
        f"SEM CAPA: {contagem.get('SEM CAPA', 0)} | "
        f"NAO ENCONTRADOS: {contagem.get('NAO ENCONTRADO', 0)} | "
        f"PLANO NAO IDENTIFICADO: {contagem.get('PLANO NAO IDENTIFICADO', 0)} | "
        f"ERROS: {contagem.get('ERRO', 0)}"
    )
    print("=" * 80)


def analisar_argumentos(argv=None):

    parser = argparse.ArgumentParser(
        description="Organizador automático de álbuns"
    )

    parser.add_argument("--link", help="Link da planilha do Google")
    parser.add_argument("--planilha", help="Arquivo xlsx local")
    parser.add_argument("--aba", help="Nome da aba")
    parser.add_argument("--origem", help="Pasta onde estão CAPAS e as remessas")
    parser.add_argument("--destino", help="Pasta onde o lote será criado")
    parser.add_argument("--col-plano", help="Coluna do plano (C ou 3)")
    parser.add_argument("--col-nome", help="Coluna do nome (A ou 1)")
    parser.add_argument("--col-codigo", help="Coluna do código SGE (B ou 2)")
    parser.add_argument("--plano", help="Filtrar por plano ou tamanho (20X20, Essencial...)")
    parser.add_argument("--simular", action="store_true", help="Não copia e não cria pastas")
    parser.add_argument("--real", action="store_true", help="Copia os arquivos e cria as pastas")
    parser.add_argument("--sem-pausa", action="store_true", help="Não espera ENTER no fim")
    parser.add_argument("--listar-abas", action="store_true", help="Lista as abas da planilha")

    return parser.parse_args(argv)


def preparar_planilha(args):

    global PLANILHA

    if args.link:

        PLANILHA = baixar_planilha_link(args.link, PLANILHA_BAIXADA)
        print("PLANILHA:", PLANILHA)

    elif args.planilha:

        PLANILHA = Path(args.planilha)


def main(argv=None):

    global MODO_TESTE, DESTINO, SEM_PAUSA

    args = analisar_argumentos(argv)
    SEM_PAUSA = args.sem_pausa

    if args.real and args.simular:

        print("Use só --simular ou só --real.")
        return 2

    if args.real:

        MODO_TESTE = False

    elif args.simular:

        MODO_TESTE = True

    if args.destino:

        DESTINO = Path(args.destino)

    try:

        preparar_planilha(args)

    except ValueError as erro:

        print()
        print("❌", erro)
        return 1

    if args.listar_abas:

        if not PLANILHA.exists():

            print("❌ PLANILHA NÃO ENCONTRADA")
            print(PLANILHA)
            return 1

        livro = load_workbook(PLANILHA, read_only=True, data_only=True)

        for nome in livro.sheetnames:

            print("ABA:", nome)

        livro.close()
        return 0

    print()
    print("=" * 80)
    print("       ORGANIZADOR AUTOMÁTICO DE ÁLBUNS")
    print("       TODAS AS REMESSAS")
    print("=" * 80)
    print()

    if MODO_TESTE:

        print("🔵 MODO TESTE ATIVADO")
        print("NENHUM ARQUIVO SERÁ COPIADO.")
        print("NENHUMA PASTA SERÁ CRIADA.")

    else:

        print("🔴 MODO REAL ATIVADO")
        print("Os arquivos serão COPIADOS.")
        print("As pastas de destino serão criadas.")
        print("Os originais serão mantidos.")

    albuns_colacao = encontrar_albuns_colacao(args.origem)

    if not albuns_colacao:

        return 1

    remessas = encontrar_remessas(albuns_colacao)

    if not remessas:

        print()
        print("❌ NENHUMA REMESSA ENCONTRADA")
        return 1

    try:

        coluna_plano = coluna_para_numero(args.col_plano, "plano") if args.col_plano else None
        coluna_nome = coluna_para_numero(args.col_nome, "nome") if args.col_nome else None
        coluna_codigo = coluna_para_numero(args.col_codigo, "código") if args.col_codigo else None

    except ValueError as erro:

        print()
        print("❌", erro)
        return 1

    clientes = carregar_planilha(
        PLANILHA,
        args.aba,
        args.plano,
        coluna_plano,
        coluna_nome,
        coluna_codigo,
    )

    if not clientes:

        print()
        print("❌ NENHUM CLIENTE ENCONTRADO")
        return 1

    linhas = []
    indice = indexar_pastas_album(remessas)

    for cliente in clientes:

        _, itens = itens_por_tamanho(cliente)

        try:

            linhas.extend(
                processar_cliente(
                    cliente,
                    albuns_colacao,
                    indice,
                )
            )

        except Exception as erro:

            print()
            print("❌ ERRO PROCESSANDO CLIENTE")
            print(cliente["nome"])
            print(erro)

            linhas.append(linha_relatorio(
                cliente,
                {"linha": "", "plano": "", "tamanho": ""},
                "",
                "ERRO",
                "",
                "",
                str(erro),
            ))

    if not MODO_TESTE:

        renomear_lote(clientes, linhas)

    salvar_relatorio(linhas)
    imprimir_resumo(linhas, remessas, clientes)
    return 0


# ====================================================================
# EXECUTAR
# ====================================================================

if __name__ == "__main__":

    codigo = 1

    try:

        codigo = main()

        if codigo is None:

            codigo = 0

    except KeyboardInterrupt:

        print()
        print("⚠ PROGRAMA INTERROMPIDO")

    except PermissionError as erro:

        print()
        print("=" * 80)
        print("❌ SEM PERMISSÃO")
        print("=" * 80)
        print(erro)

    except Exception as erro:

        print()
        print("=" * 80)
        print("❌ ERRO INESPERADO")
        print("=" * 80)
        print()
        print(type(erro).__name__)
        print(erro)

    finally:

        pausa()

    raise SystemExit(codigo)