# -*- coding: utf-8 -*-
"""
ENVIO DE ÁLBUM PARA APROVAÇÃO NO MAISBOOK

Para cada pessoa com STATUS DIAGRAMADO na planilha ALBUNS 25.2:
  1. Procura no servidor os JPG do álbum e a capa (sem copiar para o lote)
  2. Abre o MaisBook no Chrome, pesquisa pelo SGE e abre o Drive do cliente
  3. Cria a pasta ALBUM APROVAÇAO
     - um álbum: os JPG e a capa entram direto nessa pasta
     - dois álbuns (24x34, que também leva o 20x20): cria 20x20 e 24x34
  4. Espera o Drive terminar de carregar as fotos
  5. Copia o link da pasta ALBUM APROVAÇAO
  6. No cartão do cliente, abre Editar, cola o link em Álbum pronto e salva
  7. Só então escolhe o template Aprovação de álbum e envia o e-mail
  8. Só marca como enviado depois que o MaisBook confirmar

Uso:
  python ENVIO_APROVACAO_MAISBOOK.PY                 envia de verdade
  python ENVIO_APROVACAO_MAISBOOK.PY --sge 93441     uma pessoa
  python ENVIO_APROVACAO_MAISBOOK.PY --limite 1      só a primeira
  python ENVIO_APROVACAO_MAISBOOK.PY --simular       procura e abre o cliente, sem pasta, sem arquivo e sem e-mail
  python ENVIO_APROVACAO_MAISBOOK.PY --continuar     retoma quem ainda não está em envio_aprovacao_feitos.json
  python ENVIO_APROVACAO_MAISBOOK.PY --listar       só lista quem está DIAGRAMADO

Janela: guia Enviar aprovação em organizador_ui.pyw
"""

import argparse
import csv
import json
import os
import re
import shutil
import sqlite3
import sys
import tempfile
import time
from datetime import datetime
from importlib.machinery import SourceFileLoader
from pathlib import Path
from urllib.parse import quote

PASTA = Path(__file__).resolve().parent
PASTA_MB = PASTA.parent / "cpf_maisbook"
_ARQ_PROC = PASTA / "PROCUARAÇAO_SERVIDOR.PY"
if not _ARQ_PROC.exists():
    _ARQ_PROC = PASTA / "PROCUARACAO_SERVIDOR.PY"
proc = SourceFileLoader("procura_servidor", str(_ARQ_PROC)).load_module()

_ARQ_MB = PASTA_MB / "CPF_MAISBOOK.PY"
if _ARQ_MB.exists():
    mb = SourceFileLoader("cpf_maisbook", str(_ARQ_MB)).load_module()
    base = mb.base
    nomes_batem = mb.nomes_batem
else:
    class _FallbackBase:
        @staticmethod
        def ir(pagina, url, timeout=20000):
            pagina.goto(url, wait_until="domcontentloaded", timeout=timeout)
        @staticmethod
        def texto_da_tela(pagina):
            try:
                return proc.normalizar(pagina.inner_text("body"))
            except Exception:
                return ""
    class _FallbackMB:
        PERFIL = PASTA / "maisbook_perfil"
        @staticmethod
        def carregar_config():
            return {"url_maisbook": "https://www.maisbook.com.br/admin/index.php?pagina=clientes", "timeout_ms": 20000}
        @staticmethod
        def login_maisbook(pagina, timeout=20000):
            return False
    mb = _FallbackMB()
    base = _FallbackBase()
    def nomes_batem(a, b):
        return proc.normalizar(a) == proc.normalizar(b)

PLANILHA_LINK = "https://docs.google.com/spreadsheets/d/1FeYQop9z6rVAFQuIs1EOiuQcE84t0iBgamYtMWUIAU0/edit?usp=sharing"
PLANILHA = PASTA / "planilha_envio_aprovacao.xlsx"
FEITOS = PASTA / "envio_aprovacao_feitos.json"
RELATORIO = PASTA / "relatorio_envio_aprovacao.csv"
PARAR = PASTA / "envio_aprovacao_parar.flag"
PRINTS = PASTA / "envio_prints"

PASTA_RAIZ = "ALBUM APROVAÇAO"
PASTA_TAMANHO = {"20X20": "20x20", "24X34": "24x34"}
ESPERA_DRIVE_S = 30
ESPERA_UPLOAD_S = 900
TEMPLATE_VALOR = "aprovacao_album"


def _configurar_saida():
    for canal in (sys.stdout, sys.stderr):
        if hasattr(canal, "reconfigure"):
            try:
                canal.reconfigure(encoding="utf-8", errors="replace")
            except Exception:
                pass


_configurar_saida()


class PararPessoa(Exception):
    def __init__(self, resultado, detalhe):
        super().__init__(detalhe)
        self.resultado = resultado
        self.detalhe = detalhe


def normalizar(texto):
    return proc.normalizar(texto)


def carregar_json(arquivo):
    try:
        return json.loads(Path(arquivo).read_text(encoding="utf-8"))
    except Exception:
        return {}


def salvar_json(arquivo, dados):
    Path(arquivo).write_text(json.dumps(dados, indent=2, ensure_ascii=False), encoding="utf-8")


def salvar_relatorio(linhas):
    campos = ["NOME", "SGE", "CONTRATO", "CODIGO", "COLUNA_M_CAPA", "TAMANHOS", "JPG", "CAPA", "RESULTADO", "MENSAGEM", "AVISO"]
    try:
        with RELATORIO.open("w", newline="", encoding="utf-8-sig") as arquivo:
            escritor = csv.DictWriter(arquivo, fieldnames=campos, delimiter=";")
            escritor.writeheader()
            for linha in linhas:
                escritor.writerow({campo: linha.get(campo, "") for campo in campos})
    except PermissionError:
        print(f"    (não consegui gravar {RELATORIO.name} — ele está aberto no Excel?)")


def tirar_print(pagina, nome):
    if pagina is None:
        return ""
    PRINTS.mkdir(exist_ok=True)
    arquivo = PRINTS / f"{datetime.now():%Y%m%d_%H%M%S}_{re.sub(r'[^0-9A-Za-z_-]', '_', nome)}.png"
    try:
        pagina.screenshot(path=str(arquivo), full_page=True)
        return str(arquivo)
    except Exception:
        return ""


# ==========================================================
# Planilha
# ==========================================================

COLUNA_CODIGO = 9
COLUNA_CAPA_PADRAO = 13  # Coluna M (CAPA)
NOMES_CABECALHO_CLIENTE = (
    "nome",
    "nome completo",
    "nome do cliente",
    "nome do aluno",
    "nome do formando",
    "cliente",
    "aluno",
    "formando",
)

BLOQUEIOS_COLUNA_M = {
    "",
    "em diagramacao",
    "em producao",
    "aguardando",
    "aguardando capa",
    "aguardando foto",
    "pendente",
    "nao",
    "sem capa",
    "sem album",
    "falta",
    "faltando",
}


def capa_liberada_na_planilha(valor):
    """
    Verifica se a Coluna M (CAPA) recebeu o OK para enviar (ex.: DIAGRAMADO, OK, CAPA ENTREGUE, SIM, PRONTO).
    Se estiver vazia ou em 'EM DIAGRAMAÇÃO' / 'EM PRODUÇÃO', retorna False.
    """
    texto = normalizar(valor)
    if not texto or texto in BLOQUEIOS_COLUNA_M:
        return False
    return True


def limpar_sge(valor):
    texto = str(valor or "").strip()
    if texto.upper().startswith("633."):
        texto = texto[4:].strip()
    texto = re.sub(r"\.\d+$", "", texto)
    return texto.strip()


def cabecalho_da_folha(folha):
    """Acha a linha que tem NOME/CLIENTE, SGE, TAMANHO, STATUS e CAPA (Coluna M) juntos."""

    linhas_amostra = list(folha.iter_rows(max_row=80, values_only=True))
    for numero, linha in enumerate(linhas_amostra[:40], start=1):
        textos = [normalizar(valor) for valor in linha]
        conjunto = set(textos)
        if not {"sge", "tamanho", "status"} <= conjunto:
            continue
        candidatos_nome = [
            (0 if texto == "nome" else 1, i + 1)
            for i, texto in enumerate(textos)
            if texto in NOMES_CABECALHO_CLIENTE
        ]
        if not candidatos_nome:
            continue

        melhor_nome = candidatos_nome[0][1]
        melhor_chave = None
        for prio, ind in candidatos_nome:
            preenchidas = sum(
                1
                for l_abaixo in linhas_amostra[numero:]
                if ind <= len(l_abaixo) and normalizar(l_abaixo[ind - 1])
            )
            chave = (preenchidas > 0, -prio, preenchidas, -ind)
            if melhor_chave is None or chave > melhor_chave:
                melhor_chave = chave
                melhor_nome = ind

        contrato = next((i + 1 for i, texto in enumerate(textos) if texto in ("cont", "contrato", "ctt")), 0)
        codigo = textos.index("maisbook") + 1 if "maisbook" in textos else COLUNA_CODIGO
        status_idx = textos.index("status") + 1
        # A Coluna M (13) tem o título CAPA (depois da coluna STATUS)
        capa_col = next(
            (i + 1 for i, texto in enumerate(textos) if texto == "capa" and (i + 1) > status_idx),
            0,
        )
        if not capa_col:
            capa_col = textos.index("capa") + 1 if "capa" in textos else COLUNA_CAPA_PADRAO

        return numero, {
            "nome": melhor_nome,
            "sge": textos.index("sge") + 1,
            "tamanho": textos.index("tamanho") + 1,
            "status": status_idx,
            "capa": capa_col,
            "contrato": contrato,
            "codigo": codigo,
        }
    return None


def so_contrato(valor):
    digitos = re.sub(r"\D", "", str(valor or "").split("/")[-1].split(".")[0])
    return digitos.lstrip("0") or ("0" if digitos else "")


def celula(linha, indice):
    if not indice or indice > len(linha):
        return ""
    return proc.texto_celula(linha[indice - 1])


def tamanhos_do_plano(plano):
    tamanho = proc.tamanho_do_plano(plano)
    if not tamanho:
        return []
    if tamanho == "24X34":
        return ["20X20", "24X34"]
    return [tamanho]


EDICOES_LOCAIS = PASTA / "edicoes_planilha_local.json"


def aplicar_edicoes_locais_no_xlsx(caminho_xlsx):
    """Aplica edições recentes feitas pelo Dashboard na cópia local .xlsx (evita atraso de cache do export Google)."""
    if not EDICOES_LOCAIS.is_file() or not Path(caminho_xlsx).is_file():
        return
    try:
        from openpyxl.utils import column_index_from_string

        dados = carregar_json(EDICOES_LOCAIS)
        registros = dados.get("linhas", {})
        if not registros:
            return
        agora = time.time()
        wb = proc.load_workbook(caminho_xlsx)
        alterou = False
        for chave, info in list(registros.items()):
            # Mantém o overlay ativo por 2 horas após a edição em tempo real
            if agora - float(info.get("ts", 0)) > 7200:
                continue
            aba = info.get("aba") or "GERAL ENVIOS"
            linha_num = int(info.get("linha") or 0)
            celulas = info.get("celulas") or {}
            if linha_num < 1 or aba not in wb.sheetnames or not celulas:
                continue
            ws = wb[aba]
            for col_ref, novo_val in celulas.items():
                col_idx = int(col_ref) if str(col_ref).isdigit() else column_index_from_string(str(col_ref).upper())
                ws.cell(row=linha_num, column=col_idx, value=novo_val)
                alterou = True
        if alterou:
            wb.save(caminho_xlsx)
        wb.close()
    except Exception:
        pass


def _ler_diagramados_de_csv_local(etapa_filtro="diagramado", silencioso=False):
    etapa_norm = normalizar(etapa_filtro or "diagramado")
    caminhos_csv = [
        PASTA.parent / "public" / "assets" / "aistudio" / "ALBUNS_25_ffb84621.2",
        PASTA.parent / "public" / "assets" / "aistudio" / "ALBUNS_25_73835ec6.2",
        PASTA.parent / "ALBUNS_25_2_COLACAO.csv",
        PASTA / "ALBUNS_25_2_COLACAO.csv",
    ]
    arq_csv = next((c for c in caminhos_csv if c.is_file()), None)
    if not arq_csv:
        return []
    edicoes_overlay = {}
    try:
        if EDICOES_LOCAIS.is_file():
            dados_ov = json.loads(EDICOES_LOCAIS.read_text(encoding="utf-8"))
            edicoes_overlay = dados_ov.get("linhas") or {}
    except Exception:
        edicoes_overlay = {}

    pessoas = {}
    with arq_csv.open("r", encoding="utf-8", errors="replace") as f:
        leitor = csv.reader(f)
        for numero, linha in enumerate(leitor, start=1):
            if numero <= 18 or len(linha) < 11:
                continue
            ov_celulas = (edicoes_overlay.get(f"GERAL ENVIOS:{numero}") or {}).get("celulas") or {}
            status_k = (ov_celulas.get("K") if "K" in ov_celulas else (linha[10] if len(linha) > 10 else "")).strip()
            if etapa_norm not in ("todos", "todas", "*") and normalizar(status_k) != etapa_norm:
                continue
            nome = (ov_celulas.get("C") if "C" in ov_celulas else (linha[2] if len(linha) > 2 else "")).strip()
            cont = so_contrato(ov_celulas.get("D") if "D" in ov_celulas else (linha[3] if len(linha) > 3 else ""))
            sge = limpar_sge(ov_celulas.get("E") if "E" in ov_celulas else (linha[4] if len(linha) > 4 else ""))
            plano = (ov_celulas.get("F") if "F" in ov_celulas else (linha[5] if len(linha) > 5 else "")).strip()
            codigo = re.sub(
                r"\s+",
                "",
                (ov_celulas.get("I") if "I" in ov_celulas else (linha[8] if len(linha) > 8 else "")),
            ).upper()
            capa_plan = (ov_celulas.get("M") if "M" in ov_celulas else (linha[12] if len(linha) > 12 else "")).strip()
            if not nome or not re.fullmatch(r"\d+", sge):
                continue
            chave_pessoa = f"{sge}|{normalizar(nome)}"
            pessoa = pessoas.setdefault(chave_pessoa, {
                "nome": nome,
                "sge": sge,
                "contrato": cont,
                "codigo": codigo,
                "linha": numero,
                "aba": "GERAL ENVIOS",
                "col_status": 11,
                "col_capa": 13,
                "status_planilha": status_k,
                "capa_planilha": capa_plan,
                "liberado_coluna_m": capa_liberada_na_planilha(capa_plan),
                "planos": [],
                "tamanhos": [],
                "avisos": [],
                "duplicados": [],
            })
            if status_k and not pessoa.get("status_planilha"):
                pessoa["status_planilha"] = status_k
            if capa_plan and not pessoa.get("capa_planilha"):
                pessoa["capa_planilha"] = capa_plan
            if capa_liberada_na_planilha(capa_plan):
                pessoa["liberado_coluna_m"] = True
            if plano and plano not in pessoa["planos"]:
                pessoa["planos"].append(plano)
            for tamanho in tamanhos_do_plano(plano):
                if tamanho not in pessoa["tamanhos"]:
                    pessoa["tamanhos"].append(tamanho)
    lista = list(pessoas.values())
    if not silencioso and etapa_norm == "diagramado":
        qtd_liberados_m = sum(1 for p in lista if p.get("liberado_coluna_m"))
        qtd_aguardando_m = len(lista) - qtd_liberados_m
        print(
            f"Aba: GERAL ENVIOS (CSV) | STATUS DIAGRAMADO (Coluna K): {len(lista)} | "
            f"Liberados na Coluna M (CAPA): {qtd_liberados_m} | Aguardando OK na Coluna M: {qtd_aguardando_m}"
        )
    return lista


def ler_diagramados(etapa_filtro="diagramado", silencioso=False):
    etapa_norm = normalizar(etapa_filtro or "diagramado")
    if proc.load_workbook is None:
        return _ler_diagramados_de_csv_local(etapa_filtro, silencioso=silencioso)
    caminho = proc.baixar_planilha_link(PLANILHA_LINK, PLANILHA)
    aplicar_edicoes_locais_no_xlsx(caminho)
    livro = proc.load_workbook(caminho, data_only=True)
    melhor = None
    aba_padrao = None
    for nome in livro.sheetnames:
        folha = livro[nome]
        achado = cabecalho_da_folha(folha)
        if not achado:
            continue
        linha_cab, colunas = achado
        if aba_padrao is None:
            aba_padrao = (0, nome, linha_cab, colunas)
        qtd = 0
        for numero, linha in enumerate(folha.iter_rows(values_only=True), start=1):
            if numero <= linha_cab:
                continue
            if etapa_norm in ("todos", "todas", "*") or normalizar(celula(linha, colunas["status"])) == etapa_norm:
                qtd += 1
        if melhor is None or qtd > melhor[0]:
            melhor = (qtd, nome, linha_cab, colunas)
    if not melhor:
        melhor = aba_padrao
    if not melhor:
        livro.close()
        raise ValueError("Não achei uma aba com o cabeçalho NOME/CLIENTE, SGE, TAMANHO e STATUS")

    _, nome_aba, linha_cab, colunas = melhor
    folha = livro[nome_aba]
    pessoas = {}
    for numero, linha in enumerate(folha.iter_rows(values_only=True), start=1):
        if numero <= linha_cab:
            continue
        status_k = celula(linha, colunas["status"]).strip()
        if etapa_norm not in ("todos", "todas", "*") and normalizar(status_k) != etapa_norm:
            continue
        nome = celula(linha, colunas["nome"]).strip()
        sge = limpar_sge(celula(linha, colunas["sge"]))
        plano = celula(linha, colunas["tamanho"]).strip()
        capa_plan = celula(linha, colunas.get("capa") or COLUNA_CAPA_PADRAO).strip()
        if not nome or not re.fullmatch(r"\d+", sge):
            continue
        chave_pessoa = f"{sge}|{normalizar(nome)}"
        pessoa = pessoas.setdefault(chave_pessoa, {
            "nome": nome,
            "sge": sge,
            "contrato": so_contrato(celula(linha, colunas["contrato"])),
            "codigo": re.sub(r"\s+", "", celula(linha, colunas["codigo"])).upper(),
            "linha": numero,
            "aba": nome_aba,
            "col_status": colunas["status"],
            "col_capa": colunas.get("capa") or COLUNA_CAPA_PADRAO,
            "status_planilha": status_k,
            "capa_planilha": capa_plan,
            "liberado_coluna_m": capa_liberada_na_planilha(capa_plan),
            "planos": [],
            "tamanhos": [],
            "avisos": [],
            "duplicados": [],
        })
        if status_k and not pessoa.get("status_planilha"):
            pessoa["status_planilha"] = status_k
        if capa_plan and not pessoa.get("capa_planilha"):
            pessoa["capa_planilha"] = capa_plan
        if capa_liberada_na_planilha(capa_plan):
            pessoa["liberado_coluna_m"] = True
        if plano and plano not in pessoa["planos"]:
            pessoa["planos"].append(plano)
        for tamanho in tamanhos_do_plano(plano):
            if tamanho not in pessoa["tamanhos"]:
                pessoa["tamanhos"].append(tamanho)
    livro.close()
    lista = list(pessoas.values())
    for pessoa in lista:
        if not pessoa["codigo"]:
            pessoa["avisos"].append("sem código MaisBook na coluna I — pesquisado pelo SGE")
        if not pessoa["contrato"]:
            pessoa["avisos"].append("sem contrato na planilha — conferido pelo SGE")
    if not silencioso and etapa_norm == "diagramado":
        qtd_liberados_m = sum(1 for p in lista if p.get("liberado_coluna_m"))
        qtd_aguardando_m = len(lista) - qtd_liberados_m
        print(
            f"Aba: {nome_aba} | STATUS DIAGRAMADO (Coluna K): {len(lista)} | "
            f"Liberados na Coluna M (CAPA): {qtd_liberados_m} | Aguardando OK na Coluna M: {qtd_aguardando_m}"
        )
    return lista


def diagnosticar_elegibilidade_lote(origem_srv=None, sge_filtro="", ignorar_col_m=False):
    """
    Analisa todos os formandos da planilha e calcula exatamente por qual critério de filtragem
    cada formando foi aprovado ou reprovado para o envio em lote no MaisBook:
      1. Filtro 1 — Status na Coluna K (precisa ser DIAGRAMADO)
      2. Filtro 2 — Histórico de Já Enviados (envio_aprovacao_feitos.json)
      3. Filtro 3 — Trava da Coluna M (CAPA preenchida/liberada vs. capa nula/vazia/em produção)
      4. Filtro 4 — Conferência Física no Servidor (TEM ÁLBUM JPG + TEM CAPA)
    """
    todos = ler_diagramados(etapa_filtro="todos", silencioso=True)
    feitos = carregar_json(FEITOS)
    sge_alvo = limpar_sge(sge_filtro) if sge_filtro else ""

    if sge_alvo:
        todos_filtrados = [p for p in todos if p["sge"] == sge_alvo]
    else:
        todos_filtrados = list(todos)

    contagem_status_k = {}
    aprovados_status_k = []
    reprovados_status_k = []

    for p in todos_filtrados:
        st_bruto = (p.get("status_planilha") or "").strip() or "(vazio)"
        contagem_status_k[st_bruto] = contagem_status_k.get(st_bruto, 0) + 1
        if normalizar(st_bruto) == "diagramado":
            aprovados_status_k.append(p)
        else:
            reprovados_status_k.append({
                "linha": p["linha"],
                "sge": p["sge"],
                "nome": p["nome"],
                "status_k": st_bruto,
                "capa_m": p.get("capa_planilha") or "(vazia)",
                "criterio": "STATUS_INCORRETO_COL_K",
                "motivo": f"Coluna K está em '{st_bruto}' (exige DIAGRAMADO)",
            })

    ja_enviados = []
    nao_enviados = []
    for p in aprovados_status_k:
        if p["sge"] in feitos:
            info_feito = feitos.get(p["sge"]) or {}
            ja_enviados.append({
                "linha": p["linha"],
                "sge": p["sge"],
                "nome": p["nome"],
                "status_k": p.get("status_planilha") or "DIAGRAMADO",
                "capa_m": p.get("capa_planilha") or "(vazia)",
                "criterio": "JA_ENVIADO_HISTORICO",
                "motivo": f"Consta em envio_aprovacao_feitos.json ({info_feito.get('data', 'anteriormente')})",
            })
        else:
            nao_enviados.append(p)

    aprovados_col_m = []
    reprovados_capa_nula = []
    for p in nao_enviados:
        if ignorar_col_m or p.get("liberado_coluna_m"):
            aprovados_col_m.append(p)
        else:
            val_m = p.get("capa_planilha") or ""
            reprovados_capa_nula.append({
                "linha": p["linha"],
                "sge": p["sge"],
                "nome": p["nome"],
                "status_k": p.get("status_planilha") or "DIAGRAMADO",
                "capa_m": val_m or "(nula/vazia)",
                "criterio": "CAPA_NULA_COLUNA_M",
                "motivo": f"Coluna M (CAPA) está {'nula/vazia' if not val_m else repr(val_m)} — exige OK/DIAGRAMADO/CAPA",
            })

    albuns_dir = None
    indice_srv = {}
    erro_servidor = ""
    try:
        albuns_dir, indice_srv = preparar_indice(origem_srv)
    except Exception as e:
        erro_servidor = str(e)

    elegiveis_confirmados = []
    bloqueados_falta_capa_srv = []
    bloqueados_falta_album_srv = []
    bloqueados_sem_arquivos_srv = []
    detalhes_diagramados = []

    for p in aprovados_status_k:
        ja_env = p["sge"] in feitos
        lib_m = bool(ignorar_col_m or p.get("liberado_coluna_m"))
        info_srv = None
        if albuns_dir is not None:
            try:
                info_srv = conferir_pessoa_no_servidor(p, albuns_dir, indice_srv)
            except Exception as err_p:
                info_srv = {
                    "completo": False,
                    "temAlbum": False,
                    "temCapa": False,
                    "resumoAlbum": "Erro ao conferir",
                    "resumoCapa": "Erro ao conferir",
                    "resumoGeral": str(err_p),
                    "mensagem": str(err_p),
                }

        if ja_env:
            criterio_final = "JA_ENVIADO_HISTORICO"
            motivo_final = "Já consta em envio_aprovacao_feitos.json"
            elegivel = False
        elif not lib_m:
            criterio_final = "CAPA_NULA_COLUNA_M"
            motivo_final = f"Coluna M (CAPA) nula/vazia ({p.get('capa_planilha') or 'vazia'})"
            elegivel = False
        elif info_srv is None:
            criterio_final = "CONFERENCIA_SERVIDOR_PENDENTE"
            motivo_final = f"Pasta do servidor não acessível ({erro_servidor or 'não conferido'})"
            elegivel = False
        elif info_srv.get("completo"):
            criterio_final = "ELEGIVEL_PRONTO"
            motivo_final = f"Elegível: Col.K=DIAGRAMADO + Col.M='{p.get('capa_planilha')}' + Servidor OK ({info_srv.get('resumoAlbum')} | {info_srv.get('resumoCapa')})"
            elegivel = True
            elegiveis_confirmados.append({
                "linha": p["linha"],
                "sge": p["sge"],
                "nome": p["nome"],
                "tamanho": ", ".join(p.get("tamanhos") or []),
                "status_k": p.get("status_planilha") or "DIAGRAMADO",
                "capa_m": p.get("capa_planilha") or "OK",
                "resumo_album": info_srv.get("resumoAlbum"),
                "resumo_capa": info_srv.get("resumoCapa"),
                "criterio": criterio_final,
                "motivo": motivo_final,
            })
        elif info_srv.get("temAlbum") and not info_srv.get("temCapa"):
            criterio_final = "SERVIDOR_FALTA_CAPA"
            motivo_final = f"Tem Álbum ({info_srv.get('resumoAlbum')}), mas FALTA CAPA no servidor ({info_srv.get('resumoCapa')})"
            elegivel = False
            bloqueados_falta_capa_srv.append({
                "linha": p["linha"],
                "sge": p["sge"],
                "nome": p["nome"],
                "capa_m": p.get("capa_planilha") or "OK",
                "criterio": criterio_final,
                "motivo": motivo_final,
            })
        elif info_srv.get("temCapa") and not info_srv.get("temAlbum"):
            criterio_final = "SERVIDOR_FALTA_ALBUM"
            motivo_final = f"Tem Capa ({info_srv.get('resumoCapa')}), mas FALTA ÁLBUM/LÂMINAS no servidor ({info_srv.get('resumoAlbum')})"
            elegivel = False
            bloqueados_falta_album_srv.append({
                "linha": p["linha"],
                "sge": p["sge"],
                "nome": p["nome"],
                "capa_m": p.get("capa_planilha") or "OK",
                "criterio": criterio_final,
                "motivo": motivo_final,
            })
        else:
            criterio_final = "SERVIDOR_SEM_ALBUM_E_SEM_CAPA"
            motivo_final = f"Sem Álbum e Sem Capa no servidor ({info_srv.get('resumoAlbum')} | {info_srv.get('resumoCapa')})"
            elegivel = False
            bloqueados_sem_arquivos_srv.append({
                "linha": p["linha"],
                "sge": p["sge"],
                "nome": p["nome"],
                "capa_m": p.get("capa_planilha") or "OK",
                "criterio": criterio_final,
                "motivo": motivo_final,
            })

        detalhes_diagramados.append({
            "linha": p["linha"],
            "sge": p["sge"],
            "nome": p["nome"],
            "contrato": p.get("contrato") or "",
            "tamanhos": p.get("tamanhos") or [],
            "status_k": p.get("status_planilha") or "DIAGRAMADO",
            "capa_m": p.get("capa_planilha") or "",
            "liberado_col_m": lib_m,
            "ja_enviado": ja_env,
            "servidor_conferido": info_srv is not None,
            "servidor_completo": bool(info_srv and info_srv.get("completo")),
            "resumo_album_srv": info_srv.get("resumoAlbum") if info_srv else "Conferência pendente",
            "resumo_capa_srv": info_srv.get("resumoCapa") if info_srv else "Conferência pendente",
            "elegivel": elegivel,
            "criterio": criterio_final,
            "motivo": motivo_final,
        })

    causa_zero = ""
    if len(elegiveis_confirmados) == 0:
        if sge_alvo and not todos_filtrados:
            causa_zero = f"Filtro --sge {sge_alvo} ativo, mas esse SGE não existe na planilha."
        elif sge_alvo and not aprovados_status_k:
            st_atual = todos_filtrados[0].get("status_planilha") or "(vazio)"
            causa_zero = (
                f"Filtro --sge {sge_alvo} ({todos_filtrados[0]['nome']}) ativo, mas o STATUS na Coluna K é "
                f"'{st_atual}' (precisa ser DIAGRAMADO). Limpe o campo SGE para enviar o lote completo!"
            )
        elif not aprovados_status_k:
            causa_zero = "Nenhum formando está com STATUS = DIAGRAMADO na Coluna K da planilha atual."
        elif not nao_enviados:
            causa_zero = f"Todos os {len(aprovados_status_k)} formandos DIAGRAMADO já constam como enviados em envio_aprovacao_feitos.json."
        elif not aprovados_col_m:
            causa_zero = f"Todos os {len(nao_enviados)} formandos DIAGRAMADO estão com a Coluna M (CAPA) nula/vazia."
        else:
            total_bloq_srv = len(bloqueados_falta_capa_srv) + len(bloqueados_falta_album_srv) + len(bloqueados_sem_arquivos_srv)
            causa_zero = (
                f"Há {len(aprovados_col_m)} formando(s) com STATUS=DIAGRAMADO e Coluna M liberada, "
                f"mas todos os {total_bloq_srv} foram bloqueados na Conferência Física do Servidor "
                f"(Falta Capa: {len(bloqueados_falta_capa_srv)}, Falta Álbum: {len(bloqueados_falta_album_srv)}, "
                f"Sem Álbum/Capa: {len(bloqueados_sem_arquivos_srv)})."
            )

    return {
        "verificadoEm": f"{datetime.now():%Y-%m-%d %H:%M:%S}",
        "servidor": str(albuns_dir) if albuns_dir else (origem_srv or ""),
        "erroServidor": erro_servidor,
        "sgeFiltroAtivo": sge_alvo,
        "ignorarColunaM": bool(ignorar_col_m),
        "funil": {
            "totalPlanilha": len(todos_filtrados),
            "reprovadosStatusIncorretoColK": len(reprovados_status_k),
            "aprovadosStatusDiagramadoColK": len(aprovados_status_k),
            "reprovadosJaEnviadosHistorico": len(ja_enviados),
            "reprovadosCapaNulaColM": len(reprovados_capa_nula),
            "aprovadosColunaMCapa": len(aprovados_col_m),
            "reprovadosServidorFaltaCapa": len(bloqueados_falta_capa_srv),
            "reprovadosServidorFaltaAlbum": len(bloqueados_falta_album_srv),
            "reprovadosServidorSemArquivos": len(bloqueados_sem_arquivos_srv),
            "elegiveisConfirmados": len(elegiveis_confirmados),
        },
        "distribuicaoStatusColK": contagem_status_k,
        "causaZeroElegiveis": causa_zero,
        "elegiveis": elegiveis_confirmados,
        "reprovadosCapaNula": reprovados_capa_nula,
        "jaEnviados": ja_enviados,
        "bloqueadosFaltaCapaServidor": bloqueados_falta_capa_srv,
        "bloqueadosFaltaAlbumServidor": bloqueados_falta_album_srv,
        "bloqueadosSemArquivosServidor": bloqueados_sem_arquivos_srv,
        "detalhesDiagramados": detalhes_diagramados,
    }


def imprimir_diagnostico_elegibilidade(diag):
    f = diag["funil"]
    print("\n" + "=" * 82)
    print("DIAGNÓSTICO DE ELEGIBILIDADE — ENVIO EM LOTE MAISBOOK (CRITÉRIOS DE FILTRAGEM)")
    print("=" * 82)
    if diag.get("sgeFiltroAtivo"):
        print(f"⚠ ATENÇÃO — FILTRO INDIVIDUAL SGE ATIVO: --sge {diag['sgeFiltroAtivo']}")
    print(f"Pasta do Servidor Analisada : {diag.get('servidor') or '(não encontrada)'}")
    print("-" * 82)
    print(f"  [0] Total de formandos analisados na planilha .............: {f['totalPlanilha']}")
    print(f"  [1] Reprovados por Status Incorreto (Coluna K != DIAGRAMADO): -{f['reprovadosStatusIncorretoColK']}")
    for st_nome, st_qtd in sorted(diag.get("distribuicaoStatusColK", {}).items(), key=lambda x: -x[1]):
        marcador = "✔ ELEGÍVEL NA COL. K" if normalizar(st_nome) == "diagramado" else "✖ Ignorado no lote"
        print(f"        • {st_nome:<24}: {st_qtd:>3} formando(s) [{marcador}]")
    print(f"      -> Passaram no Filtro 1 (Coluna K = DIAGRAMADO) .......: {f['aprovadosStatusDiagramadoColK']}")
    print(f"  [2] Reprovados por Já Constarem Enviados (feitos.json) ....: -{f['reprovadosJaEnviadosHistorico']}")
    print(f"  [3] Reprovados por Capa Nula/Vazia na Coluna M (CAPA) .....: -{f['reprovadosCapaNulaColM']}")
    print(f"      -> Passaram nos Filtros 1+2+3 (Prontos na Planilha) ...: {f['aprovadosColunaMCapa']}")
    print(f"  [4] Conferência Física no Servidor (Álbum JPG + Capa):")
    print(f"        • Bloqueados — Falta Capa no Servidor (Só Álbum) ....: -{f['reprovadosServidorFaltaCapa']}")
    print(f"        • Bloqueados — Falta Álbum/Lâminas (Só Capa/Incomp.) : -{f['reprovadosServidorFaltaAlbum']}")
    print(f"        • Bloqueados — Sem Álbum e Sem Capa no Servidor .....: -{f['reprovadosServidorSemArquivos']}")
    print("-" * 82)
    print(f"  [✔] PESSOAS ELEGÍVEIS CONFIRMADAS PARA ENVIO EM LOTE ......: {f['elegiveisConfirmados']}")
    if diag.get("causaZeroElegiveis"):
        print(f"\n  🚨 CAUSA RAIZ DE 0 PESSOAS ELEGÍVEIS:\n     {diag['causaZeroElegiveis']}")
    print("=" * 82)

    if diag.get("detalhesDiagramados"):
        print("\nDETALHAMENTO INDIVIDUAL DOS FORMANDOS COM STATUS = DIAGRAMADO:")
        for d in diag["detalhesDiagramados"]:
            icone = "✔ ELEGÍVEL" if d["elegivel"] else f"✖ BLOQUEADO [{d['criterio']}]"
            print(
                f"  Linha {d['linha']:>3} | SGE {d['sge']:<6} | {d['nome']:<35} | "
                f"Col.M={d['capa_m'] or '(nula)':<10} | {icone} -> {d['motivo']}"
            )
        print("=" * 82)



# ==========================================================
# Servidor — só procura, não copia
# ==========================================================

IGNORAR_NO_NOME = {"da", "de", "do", "dos", "das", "e", "di", "du"}


PALAVRAS_ESTRUTURA_SERVIDOR = {
    "capa", "capas", "album", "albuns", "lamina", "laminas", "jpeg", "jpg", "png",
    "ctt", "contrato", "remessa", "prontos", "colacao", "envio", "lote",
    "30x20", "20x20", "24x34", "berlim", "diamond", "essencial", "classic", "premium",
}


def capa_e_da_pessoa(nome, arquivo, sge=""):
    """Confere se o caminho do arquivo/pasta pertence exclusivamente à pessoa (sem confundir homônimos como 'Ana Paula')."""
    texto_norm = normalizar(str(arquivo))
    palavras_texto = set(texto_norm.split())
    sge_limpo = str(sge or "").strip()
    if sge_limpo and len(sge_limpo) >= 3 and sge_limpo in palavras_texto:
        return True
    if proc.capa_da_pessoa(nome, [str(arquivo)]):
        return True
    palavras = [p for p in normalizar(nome).split() if len(p) > 2 and p not in IGNORAR_NO_NOME]
    if not palavras:
        return False
    # Só aceita abreviação (2+ palavras iniciais) se o nome do arquivo/pasta não tiver sobrenomes conflitantes de outra pessoa
    if len(palavras) >= 2 and all(palavra in palavras_texto for palavra in palavras[:2]):
        try:
            p_arq = Path(str(arquivo))
            partes_cand = [p_arq.stem]
            if p_arq.parent and p_arq.parent.name:
                partes_cand.append(p_arq.parent.name)
            palavras_cand = [
                w for w in normalizar(" ".join(partes_cand)).split()
                if len(w) > 2
                and not w.isdigit()
                and w not in IGNORAR_NO_NOME
                and w not in PALAVRAS_ESTRUTURA_SERVIDOR
            ]
            if palavras_cand and all(w in set(palavras) for w in palavras_cand):
                return True
        except Exception:
            pass
    return False


def procurar_capa_da_pessoa(albuns, nome, tamanho, sge=""):
    # 1. Tenta primeiro pelo buscador oficial de capas (que valida subpastas de tamanho + proporção da imagem)
    capa_geral = proc.procurar_capa_geral(albuns, nome, tamanho)
    if capa_geral and Path(capa_geral).is_file():
        return Path(capa_geral)

    # 2. Procura na pasta CAPAS geral por nome ou código SGE
    pasta_capas = albuns / proc.PASTA_CAPAS
    candidatos_capas = []
    if pasta_capas.exists():
        candidatos_capas.extend(proc.listar_arquivos_capas(pasta_capas))

    # 3. Também procura se existir estrutura de lote <TAMANHO>/CAPAS
    pasta_capas_lote = albuns / tamanho / "CAPAS"
    if pasta_capas_lote.exists():
        candidatos_capas.extend(proc.listar_arquivos_capas(pasta_capas_lote))

    melhor = None
    for arquivo in candidatos_capas:
        if not capa_e_da_pessoa(nome, arquivo, sge=sge):
            continue
        nomes = [arquivo.stem]
        for pai in arquivo.parents:
            if pai in (pasta_capas, pasta_capas_lote, albuns):
                break
            nomes.append(pai.name)
        mencoes = set()
        for item in nomes:
            mencoes |= proc.tamanhos_mencionados(item)
        if mencoes and tamanho not in mencoes:
            continue
        if tamanho in mencoes and mencoes == {tamanho}:
            prioridade = 2
        elif tamanho in mencoes:
            prioridade = 1
        else:
            prioridade = 0
        pts = max(proc.pontuacao_nome(nome, " ".join(nomes)), 1 if (sge and str(sge) in normalizar(arquivo.stem).split()) else -1)
        candidato = (prioridade, pts, -len(arquivo.parts))
        if melhor is None or candidato > melhor[0]:
            melhor = (candidato, arquivo)
    return melhor[1] if melhor else None


def escolher_album_da_pessoa(indice, pessoa, tamanho, quantidade, albuns=None):
    """Procura a pasta do álbum da pessoa nas Remessas/#ÁLBUNS PRONTOS ou na estrutura de Lote (<TAMANHO>/LAMINAS)."""
    nome = pessoa["nome"]
    sge = str(pessoa.get("sge") or "").strip()
    contrato = pessoa.get("contrato") or ""
    melhor = None
    for remessa, pasta, mencoes in indice:
        if tamanho not in mencoes:
            continue
        try:
            relativo = pasta.relative_to(remessa)
        except ValueError:
            relativo = pasta
        if not capa_e_da_pessoa(nome, relativo, sge=sge):
            continue
        do_contrato = bool(contrato) and contrato in normalizar(str(relativo)).split()
        pasta_laminas, laminas, extras = proc.localizar_laminas(pasta, quantidade)
        qtd = len(laminas)
        chave = (
            do_contrato,
            proc.pontuacao_no_caminho(nome, pasta, remessa),
            len(mencoes) == 1,
            qtd == quantidade and len(extras) <= 2,
            qtd > 0,
            -abs(qtd - quantidade),
            -len(pasta.parts),
        )
        if melhor is None or chave > melhor[0]:
            melhor = (chave, remessa, pasta, pasta_laminas, laminas, extras)

    # Fallback: verifica também estrutura <TAMANHO>/LAMINAS/<NOME ou SGE>
    if melhor is None and albuns is not None:
        pasta_laminas_lote = albuns / tamanho / "LAMINAS"
        if pasta_laminas_lote.is_dir():
            try:
                for sub in pasta_laminas_lote.iterdir():
                    if not sub.is_dir():
                        continue
                    if not capa_e_da_pessoa(nome, sub.name, sge=sge):
                        continue
                    p_lam, laminas, extras = proc.localizar_laminas(sub, quantidade)
                    qtd = len(laminas)
                    chave = (
                        True,
                        max(proc.pontuacao_nome(nome, sub.name), 1),
                        True,
                        qtd == quantidade,
                        qtd > 0,
                        -abs(qtd - quantidade),
                        -len(sub.parts),
                    )
                    if melhor is None or chave > melhor[0]:
                        melhor = (chave, albuns / tamanho, sub, p_lam, laminas, extras)
            except Exception:
                pass

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


def capa_do_album(albuns, nome, tamanho, escolha, sge=""):
    pasta = escolha["pasta_album"] if escolha else None
    capa = procurar_capa_da_pessoa(albuns, nome, tamanho, sge=sge)
    if capa:
        print(f"   capa ({tamanho}): {capa}")
        return capa
    if pasta:
        capa = proc.procurar_capa_antiga(pasta, tamanho)
    if not capa and escolha:
        extras = [
            arquivo for arquivo in (escolha.get("extras") or [])
            if proc.eh_lamina(arquivo) and proc.eh_nome_capa(arquivo.stem)
        ]
        capa = extras[0] if extras else None
    if capa:
        print(f"   capa ({tamanho}): {capa}")
    return capa


def localizar_albuns(pessoa, albuns, indice):
    """Devolve os arquivos de cada tamanho (capa + lâminas JPG), ou um resultado detalhado de falha."""
    tamanhos_alvo = list(pessoa.get("tamanhos") or [])
    if not tamanhos_alvo:
        for pl in (pessoa.get("planos") or []):
            for t in tamanhos_do_plano(pl):
                if t not in tamanhos_alvo:
                    tamanhos_alvo.append(t)

    if not tamanhos_alvo:
        return None, "plano/tamanho não identificado na Coluna F", "faltou arquivo"

    sge = str(pessoa.get("sge") or "").strip()
    achados = []
    problemas = []
    for tamanho in tamanhos_alvo:
        quantidade = proc.QUANTIDADE_LAMINAS[tamanho]
        escolha = escolher_album_da_pessoa(indice, pessoa, tamanho, quantidade, albuns=albuns)
        laminas = list(escolha["laminas"]) if escolha else []
        if escolha:
            print(f"   álbum {tamanho}: {escolha['pasta_album']} ({len(laminas)}/{quantidade} lâminas JPG)")
        else:
            print(f"   álbum {tamanho}: NÃO ENCONTRADO no servidor")
        capa = capa_do_album(albuns, pessoa["nome"], tamanho, escolha, sge=sge)
        if not capa:
            print(f"   capa {tamanho}: NÃO ENCONTRADA no servidor")

        if not escolha:
            problemas.append(f"{tamanho}: álbum não encontrado")
        elif len(laminas) < quantidade:
            problemas.append(f"{tamanho}: faltam lâminas ({len(laminas)}/{quantidade})")
        if not capa:
            problemas.append(f"{tamanho}: capa não encontrada")
        arquivos = list(laminas)
        if capa and capa not in arquivos:
            arquivos.append(capa)
        achados.append({
            "tamanho": tamanho,
            "laminas": len(laminas),
            "esperadas": quantidade,
            "tem_pasta_album": escolha is not None,
            "tem_album": bool(escolha is not None and len(laminas) >= quantidade),
            "pasta_album": str(escolha["pasta_album"]) if escolha else "",
            "pasta_laminas": str(escolha["pasta_laminas"]) if escolha and escolha.get("pasta_laminas") else "",
            "capa": capa is not None,
            "arquivo_capa": str(capa) if capa else "",
            "arquivos": arquivos,
            "remessa": escolha["remessa"].name if escolha and escolha.get("remessa") else "",
        })

    if problemas:
        return achados, "; ".join(problemas), "faltou arquivo"
    return achados, "", ""


def resolver_pasta_servidor(caminho_servidor=None):
    """Localiza a pasta física do servidor de álbuns e capas (aceita caminho customizado da UI ou config)."""
    candidatos = []
    if caminho_servidor and str(caminho_servidor).strip():
        candidatos.append(Path(str(caminho_servidor).strip()))
    try:
        cfg_org = carregar_json(PASTA / "organizador_config.json")
        if cfg_org.get("origem"):
            candidatos.append(Path(str(cfg_org["origem"]).strip()))
    except Exception:
        pass
    candidatos.extend([
        proc.SERVIDOR,
        Path(r"D:\ÁLBUNS COLAÇÃO 2025.2"),
        Path(r"\\192.168.0.29\hd-61\ÁLBUNS COLAÇÃO 2025.2"),
        Path(r"\\192.168.0.29\hd-61\ÁLBUNS COLAÇÃO"),
        Path(r"D:\ÁLBUNS COLAÇÃO"),
        PASTA.parent / "servidor_albuns_colacao",
        PASTA / "servidor_albuns_colacao",
    ])
    vistos = set()
    for cand in candidatos:
        try:
            chave = str(cand)
            if chave in vistos:
                continue
            vistos.add(chave)
            if cand.exists() and cand.is_dir():
                return cand
        except Exception:
            continue
    alvo_erro = candidatos[0] if candidatos else proc.SERVIDOR
    raise FileNotFoundError(f"Servidor de álbuns/capas não encontrado: {alvo_erro}")


def preparar_indice(caminho_servidor=None):
    setattr(proc.listar_arquivos_capas, "cache", None)
    pasta_srv = resolver_pasta_servidor(caminho_servidor)
    albuns = proc.encontrar_albuns_colacao(pasta_srv)
    if albuns is None:
        raise FileNotFoundError(f"Servidor não encontrado: {pasta_srv}")
    remessas = proc.encontrar_remessas(albuns)
    if not remessas:
        raise FileNotFoundError(f"Nenhuma remessa em {albuns}")
    return albuns, proc.indexar_pastas_album(remessas)


def conferir_pessoa_no_servidor(pessoa, albuns, indice):
    """
    Executa a conferência real de CAPA e ÁLBUM no servidor para um formando,
    retornando separadamente se TEM ÁLBUM (e quantas lâminas JPG) e se TEM CAPA (e o caminho).
    """
    achados, msg, resultado = localizar_albuns(pessoa, albuns, indice)
    itens = []
    for it in (achados or []):
        itens.append({
            "tamanho": it["tamanho"],
            "laminasEncontradas": it["laminas"],
            "laminasEsperadas": it["esperadas"],
            "temPastaAlbum": bool(it.get("tem_pasta_album")),
            "temAlbum": bool(it.get("tem_album")),
            "pastaAlbum": it.get("pasta_album") or "",
            "pastaLaminas": it.get("pasta_laminas") or "",
            "temCapa": bool(it.get("capa")),
            "arquivoCapa": it.get("arquivo_capa") or "",
            "remessa": it.get("remessa") or "",
        })

    if not itens:
        return {
            "id": str(pessoa.get("id") or ""),
            "linha": pessoa.get("linha") or 0,
            "sge": str(pessoa.get("sge") or ""),
            "nome": pessoa.get("nome") or "",
            "contrato": str(pessoa.get("contrato") or ""),
            "temAlbum": False,
            "temPastaAlbum": False,
            "temCapa": False,
            "completo": False,
            "resumoAlbum": "✖ Plano s/ tamanho",
            "resumoCapa": "✖ Plano s/ tamanho",
            "resumoGeral": f"⚠ {msg or 'Plano não identificado'}",
            "mensagem": msg or "Plano não identificado",
            "pastaAlbum": "",
            "nomePastaAluno": "",
            "arquivoCapa": "",
            "itens": [],
        }

    tem_album = all(i["temAlbum"] for i in itens)
    tem_pasta = any(i["temPastaAlbum"] for i in itens)
    tem_capa = all(i["temCapa"] for i in itens)
    completo = bool(tem_album and tem_capa)

    primeira_pasta = next((i["pastaAlbum"] for i in itens if i.get("pastaAlbum")), "")
    primeira_capa = next((i["arquivoCapa"] for i in itens if i.get("arquivoCapa")), "")

    nome_pasta_aluno = ""
    if primeira_pasta:
        try:
            p_obj = Path(primeira_pasta)
            if p_obj.name.upper().startswith(("ÁLBUM", "ALBUM", "JPEG", "JPG", "LAMINAS", "LÂMINAS")) and p_obj.parent.name:
                nome_pasta_aluno = p_obj.parent.name
            else:
                nome_pasta_aluno = p_obj.name
        except Exception:
            nome_pasta_aluno = ""

    partes_alb = []
    for i in itens:
        if i["temAlbum"]:
            partes_alb.append(f"{i['tamanho']}: {i['laminasEncontradas']}/{i['laminasEsperadas']} JPG")
        elif i["temPastaAlbum"]:
            partes_alb.append(f"{i['tamanho']}: INCOMPLETO ({i['laminasEncontradas']}/{i['laminasEsperadas']} JPG)")
        else:
            partes_alb.append(f"{i['tamanho']}: NÃO")

    sufixo_pasta = f" · Pasta: {nome_pasta_aluno}" if nome_pasta_aluno else ""
    if tem_album:
        resumo_album = "✔ SIM (" + ", ".join(partes_alb) + sufixo_pasta + ")"
    elif tem_pasta:
        resumo_album = "⚠ " + ", ".join(partes_alb) + sufixo_pasta
    else:
        resumo_album = "✖ SEM ÁLBUM (" + ", ".join(partes_alb) + ")"

    if tem_capa:
        nomes_capas = [Path(i["arquivoCapa"]).name for i in itens if i.get("arquivoCapa")]
        det_c = nomes_capas[0] if len(nomes_capas) == 1 else f"{len(nomes_capas)} capas ({', '.join(nomes_capas)})"
        resumo_capa = f"✔ SIM ({det_c})"
    elif any(i["temCapa"] for i in itens):
        partes_c = [f"{i['tamanho']}: {'SIM' if i['temCapa'] else 'NÃO'}" for i in itens]
        resumo_capa = "⚠ PARCIAL (" + ", ".join(partes_c) + ")"
    else:
        resumo_capa = "✖ SEM CAPA"

    if completo:
        resumo_geral = f"✔ COMPLETO ({', '.join(partes_alb)} + Capa OK)"
    elif tem_album and not tem_capa:
        resumo_geral = f"⚠ SÓ ÁLBUM — FALTA CAPA ({', '.join(partes_alb)})"
    elif not tem_album and tem_capa:
        if tem_pasta:
            resumo_geral = f"⚠ CAPA OK — FALTAM LÂMINAS ({', '.join(partes_alb)})"
        else:
            resumo_geral = "⚠ SÓ CAPA — FALTA ÁLBUM NO SERVIDOR"
    else:
        if tem_pasta:
            resumo_geral = f"✖ INCOMPLETO ({', '.join(partes_alb)} e Sem Capa)"
        else:
            resumo_geral = "✖ SEM ÁLBUM E SEM CAPA NO SERVIDOR"

    return {
        "id": str(pessoa.get("id") or ""),
        "linha": pessoa.get("linha") or 0,
        "sge": str(pessoa.get("sge") or ""),
        "nome": pessoa.get("nome") or "",
        "contrato": str(pessoa.get("contrato") or ""),
        "temAlbum": tem_album,
        "temPastaAlbum": tem_pasta,
        "temCapa": tem_capa,
        "completo": completo,
        "resumoAlbum": resumo_album,
        "resumoCapa": resumo_capa,
        "resumoGeral": resumo_geral,
        "mensagem": msg,
        "pastaAlbum": primeira_pasta,
        "nomePastaAluno": nome_pasta_aluno,
        "arquivoCapa": primeira_capa,
        "itens": itens,
    }


def texto_jpg(achados):
    return "; ".join(f"{item['tamanho']} {item['laminas']}/{item['esperadas']}" for item in achados)


def texto_capa(achados):
    if not achados:
        return "NAO"
    return "SIM" if all(item["capa"] for item in achados) else "NAO"


# ==========================================================
# MaisBook e Drive
# ==========================================================

def abrir_navegador(playwright):
    """O mesmo perfil do robô do MaisBook. O Chrome instalado passa no login do Google."""

    mb.PERFIL.mkdir(exist_ok=True)
    opcoes = dict(
        headless=False,
        no_viewport=True,
        ignore_default_args=["--enable-automation"],
        args=["--start-maximized", "--disable-blink-features=AutomationControlled"],
    )
    try:
        return playwright.chromium.launch_persistent_context(str(mb.PERFIL), channel="chrome", **opcoes)
    except Exception:
        return playwright.chromium.launch_persistent_context(
            str(mb.PERFIL), headless=False, no_viewport=True, args=["--start-maximized"])


def ir_clientes(pagina, sge):
    cfg = mb.carregar_config()
    url = f"{cfg['url_maisbook']}&acao=listar&search={quote(sge)}"
    base.ir(pagina, url, int(cfg.get("timeout_ms") or 20000))
    if pagina.locator("#user_login").count():
        if not mb.login_maisbook(pagina, int(cfg.get("timeout_ms") or 20000)):
            raise PararPessoa("erro no e-mail", "o MaisBook pediu login e não consegui entrar")
        base.ir(pagina, url, int(cfg.get("timeout_ms") or 20000))
    return url


def normalizar_codigo_mb(valor):
    return re.sub(r"\s+", "", str(valor or "")).upper()


def termo_busca(pessoa):
    return (
        (pessoa.get("termo_busca_efetivo") or "").strip()
        or (pessoa.get("codigo") or "").strip()
        or (pessoa.get("sge") or "").strip()
        or (pessoa.get("nome") or "").strip()
    )


def termos_busca_candidatos(pessoa):
    """
    Ordem de busca no MaisBook:
      1. Código do MaisBook (Coluna I) — ou SGE caso a Coluna I esteja vazia
      2. Exceção (fallback): Nome do formando (Coluna C), caso não encontre pelo código do MaisBook
      3. Exceção (fallback): 2 primeiras palavras do nome (caso haja diferença de preposição no MaisBook)
      4. Código SGE (Coluna E), caso ainda não tenha sido tentado
    Em todas as etapas, continua comparando obrigatoriamente se TANTO o Nome QUANTO o Código batem.
    """
    candidatos = []
    vistos = set()

    def _adicionar(rotulo, valor):
        v = str(valor or "").strip()
        chave = normalizar(v)
        if v and chave not in vistos:
            vistos.add(chave)
            candidatos.append((rotulo, v))

    nome_completo = (pessoa.get("nome") or "").strip()
    palavras_nome = [p for p in nome_completo.split() if len(p) > 2 and normalizar(p) not in IGNORAR_NO_NOME]
    nome_curto = " ".join(palavras_nome[:2]) if len(palavras_nome) >= 2 else ""

    if (pessoa.get("codigo") or "").strip():
        _adicionar("código MaisBook", pessoa["codigo"])
        _adicionar("nome", nome_completo)
        if nome_curto and normalizar(nome_curto) != normalizar(nome_completo):
            _adicionar("nome (2 primeiras palavras)", nome_curto)
        _adicionar("código SGE", pessoa.get("sge"))
    else:
        _adicionar("código SGE", pessoa.get("sge"))
        _adicionar("nome", nome_completo)
        if nome_curto and normalizar(nome_curto) != normalizar(nome_completo):
            _adicionar("nome (2 primeiras palavras)", nome_curto)
    return candidatos


JS_LISTAR_CARTOES = r"""
() => {
  const norm = (s) => (s || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/\s+/g, ' ').trim().toUpperCase();
  const cabecalho = document.querySelector('.clientes-grid-header');
  const titulos = cabecalho ? [...cabecalho.children].map(c => norm(c.innerText)) : [];
  return [...document.querySelectorAll('.clientes-grid-row[data-cliente-id]')].map(linha => {
    const principal = linha.querySelector('.clientes-grid-main') || linha;
    const celulas = [...principal.children].map(c => (c.innerText || '').trim());
    const campo = (rotulos, posicao) => {
      const lista = Array.isArray(rotulos) ? rotulos : [rotulos];
      for (const t of lista) {
        const i = titulos.indexOf(t);
        if (i >= 0 && i < celulas.length && celulas[i]) {
          return celulas[i].split('\n')[0].trim();
        }
      }
      const valor = celulas[posicao] || '';
      return valor.split('\n')[0].trim();
    };
    return {
      id: linha.getAttribute('data-cliente-id'),
      sge: campo(['SGE', 'CODIGO SGE', 'COD. SGE'], 0),
      ctt: campo(['CTT', 'CONTRATO', 'CONT'], 1),
      nome: campo(['NOME', 'CLIENTE', 'ALUNO'], 2),
      codigo: campo(['CODIGO', 'COD. ACESSO', 'COD ACESSO', 'COD'], 3),
      album: campo(['ALBUM', 'STATUS ALBUM'], 10),
    };
  });
}
"""


def descrever_cartao(cartao):
    return f"{cartao['nome']} (SGE {cartao['sge']}, CTT {cartao['ctt']}, código {cartao['codigo']})"


def nome_cartao_bate(nome_planilha, nome_cartao):
    """Confere se o nome na planilha e o nome no cartão/modal do MaisBook pertencem à mesma pessoa."""
    if not str(nome_planilha or "").strip() or not str(nome_cartao or "").strip():
        return False
    if capa_e_da_pessoa(nome_planilha, nome_cartao):
        return True
    palavras_plan = [
        p for p in normalizar(nome_planilha).split()
        if len(p) > 2 and p not in IGNORAR_NO_NOME
    ]
    palavras_cart = [
        p for p in normalizar(nome_cartao).split()
        if len(p) > 2 and p not in IGNORAR_NO_NOME
    ]
    if not palavras_plan or not palavras_cart:
        return False
    if nomes_batem(nome_planilha, nome_cartao):
        set_plan, set_cart = set(palavras_plan), set(palavras_cart)
        if set_plan <= set_cart or set_cart <= set_plan:
            return True
    return False


def codigo_cartao_bate(pessoa, cartao):
    """
    Compara se o código do formando na planilha (Código MaisBook da Coluna I ou SGE da Coluna E)
    bate com o código do cartão no MaisBook (campo CODIGO ou campo SGE).
    """
    sge_plan = limpar_sge(pessoa.get("sge", ""))
    cod_plan = normalizar_codigo_mb(pessoa.get("codigo", ""))
    sge_cod_plan = limpar_sge(cod_plan) if cod_plan else ""

    sge_cart = limpar_sge(cartao.get("sge", ""))
    cod_cart = normalizar_codigo_mb(cartao.get("codigo", ""))
    sge_cod_cart = limpar_sge(cod_cart) if cod_cart else ""

    # 1. Compara pelo Código do MaisBook (Coluna I) se preenchido na planilha
    if cod_plan:
        if cod_cart and cod_cart == cod_plan:
            return True
        if sge_cod_plan and sge_cart and sge_cart == sge_cod_plan:
            return True
        if sge_cod_plan and sge_cod_cart and sge_cod_cart == sge_cod_plan:
            return True

    # 2. Compara pelo Código SGE (Coluna E)
    if sge_plan:
        if sge_cart and sge_cart == sge_plan:
            return True
        if cod_cart and (cod_cart == sge_plan.upper() or sge_cod_cart == sge_plan):
            return True

    return False


def cartao_e_da_pessoa(pessoa, cartao):
    """
    Valida obrigatoriamente que TANTO o Nome QUANTO o Código (MaisBook ou SGE) batem no cartão.
    Mesmo quando a busca é feita pelo nome (exceção), exige que o nome e o código batam.
    """
    if not nome_cartao_bate(pessoa.get("nome", ""), cartao.get("nome", "")):
        return False
    if not codigo_cartao_bate(pessoa, cartao):
        return False
    return True


def _filtrar_cartoes_da_pessoa(pessoa, cartoes):
    certos = [c for c in cartoes if cartao_e_da_pessoa(pessoa, c)]
    # Se o mesmo aluno (mesmo nome + mesmo código) tiver mais de um cartão por contratos distintos, desempata pelo contrato
    if len(certos) > 1 and pessoa.get("contrato"):
        por_contrato = [
            c for c in certos
            if so_contrato(c.get("ctt")) == pessoa["contrato"]
        ]
        if por_contrato:
            certos = por_contrato
    return certos


def _aguardar_cartoes_na_tela(pagina, timeout_s=20):
    limite = time.time() + timeout_s
    cartoes = []
    vazio_desde = None
    while time.time() < limite:
        cartoes = [c for c in pagina.evaluate(JS_LISTAR_CARTOES) if c.get("id")]
        if cartoes:
            break
        texto = base.texto_da_tela(pagina)
        if "mostrando 0" in texto or "nenhum cliente" in texto:
            vazio_desde = vazio_desde or time.time()
            if time.time() - vazio_desde >= 5:
                break
        else:
            vazio_desde = None
        time.sleep(0.5)
    return cartoes


def achar_cliente(pagina, pessoa):
    termo_atual = termo_busca(pessoa)
    cartoes = _aguardar_cartoes_na_tela(pagina)
    todos_cartoes_vistos = list(cartoes)
    certos = _filtrar_cartoes_da_pessoa(pessoa, cartoes)

    # EXCEÇÃO: Caso não encontre pelo código do MaisBook (ou termo inicial), procura pelo nome
    # (e demais candidatos), mas continua comparando obrigatoriamente se o NOME e o CÓDIGO batem!
    if not certos:
        candidatos = termos_busca_candidatos(pessoa)
        ja_tentados = {normalizar(termo_atual)}
        for rotulo_cand, termo_cand in candidatos:
            chave_cand = normalizar(termo_cand)
            if not chave_cand or chave_cand in ja_tentados:
                continue
            ja_tentados.add(chave_cand)
            print(
                f"    exceção MaisBook: não encontrado por '{termo_atual}' — "
                f"procurando por {rotulo_cand} '{termo_cand}' (comparando nome + código)..."
            )
            ir_clientes(pagina, termo_cand)
            cartoes_cand = _aguardar_cartoes_na_tela(pagina)
            for c in cartoes_cand:
                if not any(v.get("id") == c.get("id") for v in todos_cartoes_vistos):
                    todos_cartoes_vistos.append(c)
            certos = _filtrar_cartoes_da_pessoa(pessoa, cartoes_cand)
            if certos:
                cartoes = cartoes_cand
                pessoa["termo_busca_efetivo"] = termo_cand
                aviso_fb = (
                    f"exceção MaisBook: não achou pelo código '{termo_atual}', "
                    f"encontrado por {rotulo_cand} (nome + código conferidos)"
                )
                if aviso_fb not in pessoa.setdefault("avisos", []):
                    pessoa["avisos"].append(aviso_fb)
                print(
                    f"    encontrado por {rotulo_cand} ('{termo_cand}') — "
                    f"nome e código batem: {descrever_cartao(certos[0])}"
                )
                break

    if len(todos_cartoes_vistos) > 1:
        for cartao in todos_cartoes_vistos:
            descricao = descrever_cartao(cartao)
            if descricao not in pessoa.setdefault("duplicados", []):
                pessoa["duplicados"].append(descricao)

    if not todos_cartoes_vistos:
        cod_desc = pessoa.get("codigo") or pessoa.get("sge")
        raise PararPessoa(
            "não encontrado",
            f"MaisBook: nenhum cliente na pesquisa pelo código '{cod_desc}' nem pelo nome '{pessoa.get('nome')}'",
        )

    if not certos:
        outros = [c for c in todos_cartoes_vistos if c not in certos]
        achados = "; ".join(descrever_cartao(c) for c in outros[:3])
        cod_esperado = (
            f"código {pessoa['codigo']} / SGE {pessoa['sge']}"
            if pessoa.get("codigo")
            else f"SGE {pessoa['sge']}"
        )
        raise PararPessoa(
            "não encontrado",
            f"MaisBook: nenhum cartão onde o nome ('{pessoa.get('nome')}') e o código ({cod_esperado}) batam — achei: {achados}",
        )
    if len(certos) > 1:
        raise PararPessoa(
            "duplicado",
            "duplicado — não enviado: " + "; ".join(descrever_cartao(c) for c in certos),
        )

    cliente_id = certos[0]["id"]
    if certos[0].get("sge"):
        pessoa["sge_maisbook"] = certos[0]["sge"].strip()
    if certos[0].get("codigo"):
        pessoa["codigo_maisbook"] = certos[0]["codigo"].strip()
    if certos[0].get("album"):
        pessoa["album_maisbook"] = certos[0]["album"].strip()
    pessoa["codigo_cartao_confirmado"] = True
    linha = pagina.locator(f'.clientes-grid-row[data-cliente-id="{cliente_id}"]')
    linha.first.wait_for(state="visible", timeout=15000)
    return cliente_id, linha.first


def no_dialogo(dialog):
    try:
        if dialog.type == "beforeunload":
            dialog.accept()
        elif dialog.type == "alert":
            print(f"    aviso do navegador: {dialog.message}")
            dialog.accept()
        else:
            dialog.dismiss()
    except Exception:
        pass


def botao_novo(pagina):
    for seletor in ('button[guidedhelpid="new_menu_button"]', 'button[guidedhelpid="td_new_menu_button"]'):
        botoes = pagina.locator(seletor)
        for i in range(min(botoes.count(), 4)):
            botao = botoes.nth(i)
            if botao.is_visible():
                return botao
    return None


def esperar_drive_carregar(pagina):
    limite = time.time() + ESPERA_DRIVE_S
    while time.time() < limite:
        if botao_novo(pagina):
            return
        time.sleep(0.5)
    imagem = tirar_print(pagina, "drive_nao_carregou")
    detalhe = f"{pagina.url[:120]} | {imagem}"
    if "accounts.google" in pagina.url.lower():
        raise PararPessoa("erro no Drive", f"o Drive abriu na tela de login do Google — nada foi enviado | {detalhe}")
    raise PararPessoa("erro no Drive", f"o Drive não carregou a pasta do aluno | {detalhe}")


def pasta_atual(pagina):
    botoes = pagina.locator('[guidedhelpid="folder_path_button"]')
    textos = []
    for i in range(min(botoes.count(), 8)):
        try:
            texto = botoes.nth(i).inner_text(timeout=1000).strip()
        except Exception:
            texto = ""
        if texto:
            textos.append(texto)
    return textos[-1] if textos else ""


def esta_na_pasta(pagina, nome):
    alvo = normalizar(nome)
    if not alvo:
        return False
    titulo = normalizar(pagina.title())
    if titulo == alvo or titulo.startswith(alvo + " "):
        return True
    return normalizar(pasta_atual(pagina)) == alvo


def conferir_drive_da_pessoa(pagina, pessoa):
    limite = time.time() + 25
    while time.time() < limite:
        titulo = pagina.title()
        trilha = " ".join(
            pagina.locator('[guidedhelpid="folder_path_button"]').all_inner_texts()
        )
        if nomes_batem(pessoa["nome"], f"{titulo} {trilha}"):
            return
        time.sleep(1)
    raise PararPessoa("erro no Drive", "a pasta do Drive não é a desta pessoa — nada foi enviado")


JS_MARCAR_PASTA = r"""
(nome) => {
  const norm = (s) => (s || '')
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim();
  const alvo = norm(nome);
  const bate = (rotulo) => {
    const n = norm(String(rotulo).split('\n')[0]);
    if (!n || n === alvo) return n === alvo && !!n;
    if (!n.startsWith(alvo + ' ')) return false;
    return /pasta|folder/.test(n);
  };
  const els = [];
  const visitar = (raiz) => {
    if (!raiz || !raiz.querySelectorAll) return;
    for (const el of raiz.querySelectorAll('[data-robo-pasta]')) el.removeAttribute('data-robo-pasta');
    for (const el of raiz.querySelectorAll('[data-id], [role="row"], [role="option"], [role="gridcell"], [aria-label], [data-tooltip]')) {
      els.push(el);
    }
    for (const el of raiz.querySelectorAll('*')) if (el.shadowRoot) visitar(el.shadowRoot);
  };
  visitar(document);
  for (const el of els) {
    const rotulo = el.getAttribute('data-tooltip') || el.getAttribute('aria-label') || el.innerText || '';
    if (!bate(rotulo)) continue;
    let alvoEl = el;
    let pai = el;
    for (let i = 0; i < 8 && pai; i++) {
      if (pai.getAttribute && (pai.getAttribute('data-id') || pai.getAttribute('role') === 'row')) {
        alvoEl = pai;
        break;
      }
      pai = pai.parentElement;
    }
    alvoEl.setAttribute('data-robo-pasta', '1');
    try { alvoEl.scrollIntoView({block: 'center'}); } catch (e) {}
    return true;
  }
  return false;
}
"""

JS_ROLAR_TOPO = r"""
() => {
  const visitar = (raiz) => {
    if (!raiz || !raiz.querySelectorAll) return;
    const todos = [raiz, ...raiz.querySelectorAll('*')];
    for (const el of todos) {
      if (el.shadowRoot) visitar(el.shadowRoot);
      try {
        if (el.scrollTop > 0) el.scrollTop = 0;
      } catch (e) {}
    }
  };
  visitar(document);
}
"""

JS_ROLAR_LISTA = r"""
() => {
  let rolou = false;
  const visitar = (raiz) => {
    if (!raiz || !raiz.querySelectorAll) return;
    const todos = [raiz, ...raiz.querySelectorAll('*')];
    for (const el of todos) {
      if (el.shadowRoot) visitar(el.shadowRoot);
      try {
        if (el.scrollHeight > el.clientHeight + 40 && el.clientHeight > 120) {
          const antes = el.scrollTop;
          el.scrollTop = Math.min(el.scrollTop + Math.floor(el.clientHeight * 0.8), el.scrollHeight);
          if (el.scrollTop > antes) rolou = true;
        }
      } catch (e) {}
    }
  };
  visitar(document);
  return rolou;
}
"""

JS_FALTANDO = r"""
(nomes) => {
  const norm = (s) => String(s || '')
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .toLowerCase().replace(/\s+/g, ' ').trim();

  const labels = [];
  const visitar = (raiz) => {
    if (!raiz || !raiz.querySelectorAll) return;
    for (const el of raiz.querySelectorAll('[data-id], [role="row"], [role="gridcell"], [role="option"], [aria-label], [data-tooltip]')) {
      if (el.closest && el.closest('[role="dialog"], [role="alertdialog"]')) continue;
      const tip = el.getAttribute('data-tooltip') || '';
      const aria = el.getAttribute('aria-label') || '';
      if (tip) labels.push(norm(tip));
      if (aria) labels.push(norm(aria));
      if (el.getAttribute('data-id') || el.getAttribute('role') === 'row' || el.getAttribute('role') === 'gridcell') {
        const txt = (el.innerText || '').split('\n')[0].trim();
        if (txt) labels.push(norm(txt));
      }
    }
    for (const el of raiz.querySelectorAll('*')) {
      if (el.shadowRoot) visitar(el.shadowRoot);
    }
  };
  visitar(document.querySelector('[role="main"]') || document.body);

  const contemNomeExato = (rotulo, alvo) => {
    if (!rotulo || !alvo) return false;
    let pos = rotulo.indexOf(alvo);
    while (pos !== -1) {
      const antes = pos > 0 ? rotulo[pos - 1] : '';
      const depois = pos + alvo.length < rotulo.length ? rotulo[pos + alvo.length] : '';
      const antesOk = !antes || !/[a-z0-9]/.test(antes);
      const depoisOk = !depois || !/[a-z0-9]/.test(depois);
      if (antesOk && depoisOk) return true;
      pos = rotulo.indexOf(alvo, pos + 1);
    }
    return false;
  };

  return nomes.filter(nome => {
    const alvo = norm(nome);
    if (!alvo) return false;
    return !labels.some(rotulo => contemNomeExato(rotulo, alvo));
  });
}
"""

JS_STATUS_UPLOAD = r"""
() => {
  const norm = (s) => (s || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase();
  const juntar = (no) => {
    if (!no) return '';
    let texto = '';
    try { texto = no.innerText || ''; } catch (e) {}
    if (!no.querySelectorAll) return texto;
    for (const el of no.querySelectorAll('*')) {
      if (el.shadowRoot) texto += ' ' + juntar(el.shadowRoot);
    }
    return texto;
  };
  const texto = norm(juntar(document.body));
  const numero = texto.match(/(\d+)\s+uploads?\s+conclu/);
  const concluido = (
    texto.includes('upload conclu') ||
    texto.includes('uploads conclu') ||
    texto.includes('upload complete')
  );
  const fazendo = (
    texto.includes('fazendo upload') ||
    texto.includes('fazendo o upload') ||
    texto.includes('iniciando upload') ||
    texto.includes('iniciando o upload') ||
    texto.includes('preparando upload') ||
    texto.includes('preparando o upload') ||
    texto.includes('uploading') ||
    texto.includes('starting upload') ||
    /\b\d+\s*(min|seg)\s+restante/.test(texto) ||
    texto.includes('tempo restante')
  );
  const falha = (
    texto.includes('nao foi enviado') ||
    texto.includes('falha no upload') ||
    texto.includes('upload failed') ||
    texto.includes('erro no upload') ||
    texto.includes('upload cancelado')
  );
  return {
    fazendo,
    concluido,
    falha,
    quantidade: numero ? Number(numero[1]) : (concluido ? 1 : 0),
  };
}
"""


def clicar_item_menu(pagina, texto):
    itens = pagina.get_by_role("menuitem").filter(has_text=texto)
    limite = time.time() + 8
    while time.time() < limite:
        for i in range(min(itens.count(), 8)):
            item = itens.nth(i)
            if item.is_visible():
                item.click()
                return
        time.sleep(0.3)
    raise PararPessoa("erro no Drive", f"não achei '{texto}' no menu Novo")


def criar_pasta(pagina, nome):
    pagina.keyboard.press("Escape")
    botao = botao_novo(pagina)
    if botao is None:
        raise PararPessoa("erro no Drive", "não achei o botão Novo do Drive")
    botao.click()
    clicar_item_menu(pagina, "Nova pasta")
    dialog = pagina.get_by_role("dialog").last
    dialog.wait_for(state="visible", timeout=10000)
    campo = dialog.locator("input").first
    campo.fill(nome)
    clicou = False
    for rotulo in ("Criar", "Create"):
        botao_criar = dialog.get_by_role("button", name=rotulo, exact=True)
        if botao_criar.count() and botao_criar.first.is_visible():
            botao_criar.first.click()
            clicou = True
            break
    if not clicou:
        campo.press("Enter")
    try:
        dialog.wait_for(state="hidden", timeout=8000)
    except Exception:
        try:
            texto = normalizar(dialog.inner_text(timeout=1000))
        except Exception:
            texto = ""
        if "ja existe" in texto or "already exists" in texto:
            pagina.keyboard.press("Escape")
            return False
        pagina.keyboard.press("Escape")
    return True


def esperar_entrar(pagina, nome, segundos):
    limite = time.time() + segundos
    while time.time() < limite:
        if esta_na_pasta(pagina, nome):
            pagina.wait_for_timeout(1200)
            return True
        time.sleep(0.4)
    return False


def abrir_pasta(pagina, nome):
    if esta_na_pasta(pagina, nome):
        return
    limite = time.time() + 25
    while time.time() < limite:
        if not pagina.evaluate(JS_MARCAR_PASTA, nome):
            time.sleep(0.6)
            continue
        alvo = pagina.locator('[data-robo-pasta="1"]').last
        id_pasta = alvo.evaluate(
            """(el) => el.getAttribute('data-id')
                || (el.closest('[data-id]') && el.closest('[data-id]').getAttribute('data-id'))
                || (el.querySelector('[data-id]') && el.querySelector('[data-id]').getAttribute('data-id'))
                || ''"""
        )
        if id_pasta:
            pagina.goto(f"https://drive.google.com/drive/folders/{id_pasta}", wait_until="domcontentloaded", timeout=60000)
            if esperar_entrar(pagina, nome, 15):
                return
        try:
            alvo.click(timeout=5000)
            pagina.keyboard.press("Enter")
        except Exception:
            try:
                alvo.dblclick(timeout=5000)
            except Exception:
                pass
        if esperar_entrar(pagina, nome, 8):
            return
        time.sleep(0.4)
    raise PararPessoa("erro no Drive", f"não consegui abrir a pasta {nome}")


def garantir_pasta(pagina, nome):
    if esta_na_pasta(pagina, nome):
        print(f"    já dentro de {nome}")
        return
    limite = time.time() + 18
    lista_parada = 0
    while time.time() < limite:
        if pagina.evaluate(JS_MARCAR_PASTA, nome):
            print(f"    pasta {nome} já existe — abrindo")
            abrir_pasta(pagina, nome)
            return
        if pagina.evaluate(JS_ROLAR_LISTA):
            lista_parada = 0
        else:
            lista_parada += 1
            if lista_parada >= 3:
                break
        time.sleep(0.5)
    print(f"    criando pasta {nome}")
    if not criar_pasta(pagina, nome):
        print(f"    pasta {nome} já existe — abrindo")
        abrir_pasta(pagina, nome)
        return
    url_antes = pagina.url
    pagina.wait_for_timeout(800)
    pagina.keyboard.press("Enter")
    if esperar_entrar(pagina, nome, 8):
        return
    if pagina.url != url_antes and "drive.google.com" in pagina.url:
        try:
            pagina.go_back(timeout=15000)
        except Exception:
            pass
        pagina.wait_for_timeout(800)
    abrir_pasta(pagina, nome)


def tratar_conflito(pagina):
    dialogos = pagina.get_by_role("dialog")
    for i in range(min(dialogos.count(), 3)):
        dialog = dialogos.nth(i)
        try:
            if not dialog.is_visible():
                continue
            texto = normalizar(dialog.inner_text(timeout=1000))
        except Exception:
            continue
        if "ja existe" not in texto and "already exists" not in texto:
            continue
        caixa = dialog.get_by_role("checkbox")
        if caixa.count() and caixa.first.is_visible():
            try:
                if not caixa.first.is_checked():
                    caixa.first.check()
            except Exception:
                pass
        for rotulo in ("Manter os dois arquivos", "Manter ambos", "Keep both", "Substituir", "Replace"):
            botao = dialog.get_by_role("button", name=rotulo)
            if botao.count() and botao.first.is_visible():
                botao.first.click()
                return


def arquivos_faltando(pagina, arquivos, rolar_tudo=True):
    """
    Verifica na grade/lista real da pasta do Google Drive ([role="main"], ignorando a bandeja de upload)
    quais arquivos da lista ainda não estão presentes na pasta.
    Rola a lista do topo até o fim para garantir leitura de pastas com 20+ lâminas virtualizadas.
    """
    restantes = [arquivo.name for arquivo in arquivos]
    if not restantes:
        return []
    try:
        if rolar_tudo:
            pagina.evaluate(JS_ROLAR_TOPO)
            pagina.wait_for_timeout(300)
        for _ in range(12 if rolar_tudo else 1):
            restantes = pagina.evaluate(JS_FALTANDO, restantes)
            if not restantes or not rolar_tudo:
                break
            if not pagina.evaluate(JS_ROLAR_LISTA):
                break
            pagina.wait_for_timeout(300)
        if rolar_tudo:
            pagina.evaluate(JS_ROLAR_TOPO)
    except Exception:
        return list(arquivos)
    por_nome = {arquivo.name: arquivo for arquivo in arquivos}
    return [por_nome[nome] for nome in restantes if nome in por_nome]


def enviar_arquivos(pagina, arquivos):
    if not arquivos:
        return
    pagina.keyboard.press("Escape")
    botao = botao_novo(pagina)
    if botao is None:
        raise PararPessoa("erro no Drive", "não achei o botão Novo para enviar os arquivos")
    botao.click()
    with pagina.expect_file_chooser(timeout=20000) as escolha:
        clicar_item_menu(pagina, "Upload de arquivo")
    escolha.value.set_files([str(arquivo) for arquivo in arquivos])


def esperar_carregamento(pagina, arquivos, enviados):
    nomes = [arquivo.name for arquivo in arquivos]
    inicio = time.time()
    limite = inicio + ESPERA_UPLOAD_S
    viu_envio = False
    ciclos_sem_fazendo = 0

    # Dá 2.5s para a bandeja de upload do Drive aparecer e iniciar
    pagina.wait_for_timeout(2500)

    while time.time() < limite:
        tratar_conflito(pagina)
        try:
            status = pagina.evaluate(JS_STATUS_UPLOAD)
            faltando = pagina.evaluate(JS_FALTANDO, nomes)
        except Exception:
            time.sleep(2)
            continue

        if status.get("fazendo"):
            viu_envio = True
            ciclos_sem_fazendo = 0
            time.sleep(2)
            continue

        ciclos_sem_fazendo += 1
        quantidade = status.get("quantidade") or 0

        if status.get("falha") and faltando:
            return

        # Assim que todos os arquivos estão visíveis na pasta e não há upload em andamento, prossegue!
        if not faltando:
            pagina.wait_for_timeout(2000)
            return

        # Se a bandeja do Drive confirmou "X uploads concluídos"
        if status.get("concluido") and (quantidade >= len(enviados) or viu_envio):
            pagina.wait_for_timeout(2000)
            return

        # Se não há texto de upload ativo por 3 ciclos (~6s), faz uma varredura com scroll na pasta e libera
        if ciclos_sem_fazendo >= 3:
            faltando_scroll = arquivos_faltando(pagina, arquivos, rolar_tudo=True)
            if not faltando_scroll or viu_envio:
                pagina.wait_for_timeout(1500)
                return
            if ciclos_sem_fazendo >= 5:
                return

        time.sleep(2)
    raise PararPessoa("erro no Drive", "as fotos não terminaram de carregar no Drive")


def colocar_arquivos(pagina, arquivos):
    """
    Envia os arquivos que faltam para a pasta atual do Drive, aguarda o término do upload
    e confere rolando a pasta se todos os arquivos foram enviados (reenviando se faltar algum).
    """
    total = len(arquivos)
    faltam = arquivos_faltando(pagina, arquivos, rolar_tudo=True)
    if not faltam:
        print(f"    os {total} arquivo(s) já estavam na pasta")
        pagina.wait_for_timeout(1500)
        return

    for tentativa in range(1, 3):
        if tentativa == 1:
            print(f"    enviando {len(faltam)} arquivo(s)")
        else:
            print(f"    reenviando {len(faltam)} arquivo(s) que faltaram ({', '.join(a.name for a in faltam[:4])})")
        enviar_arquivos(pagina, faltam)
        esperar_carregamento(pagina, arquivos, faltam)
        faltam = arquivos_faltando(pagina, arquivos, rolar_tudo=True)
        if not faltam:
            print(f"    conferido no Drive: todos os {total}/{total} arquivo(s) na pasta")
            return
        try:
            status = pagina.evaluate(JS_STATUS_UPLOAD)
            if status.get("concluido") and not status.get("falha") and (status.get("quantidade") or 0) >= total:
                print(f"    conferido no Drive: {total} upload(s) concluído(s)")
                return
        except Exception:
            pass

    nomes_faltam = ", ".join(a.name for a in faltam[:6])
    raise PararPessoa(
        "erro no Drive",
        f"faltaram {len(faltam)} de {total} arquivo(s) no Drive ({nomes_faltam})",
    )


def voltar_para_pasta(pagina, nome):
    if esta_na_pasta(pagina, nome):
        return
    botoes = pagina.locator('[guidedhelpid="folder_path_button"]').filter(has_text=nome)
    for i in range(botoes.count() - 1, -1, -1):
        btn = botoes.nth(i)
        try:
            if btn.is_visible():
                btn.click(timeout=10000)
                if esperar_entrar(pagina, nome, 12):
                    return
                break
        except Exception:
            continue
    if "drive.google.com" in pagina.url:
        try:
            pagina.go_back(timeout=15000)
            if esperar_entrar(pagina, nome, 12):
                return
        except Exception:
            pass
    raise PararPessoa("erro no Drive", f"não voltei para a pasta {nome}")


def abrir_drive(contexto, linha):
    links = linha.locator('a[title="Abrir pasta no Drive"], a[title*="Drive" i], a[href*="drive.google.com"]')
    qtd = links.count()
    if not qtd:
        raise PararPessoa("erro no Drive", "não achei o botão do Drive neste cliente")

    # 1. Se qualquer link já tiver o href direto da pasta do Google Drive, abre direto sem depender de clique visual
    for i in range(qtd):
        try:
            endereco = (links.nth(i).get_attribute("href") or "").strip()
            if "drive.google.com" in endereco:
                drive = contexto.new_page()
                drive.goto(endereco, wait_until="domcontentloaded", timeout=60000)
                return drive, True
        except Exception:
            continue

    # 2. Escolhe o link visível (no modo Compacto do MaisBook o primeiro link pode estar oculto nos detalhes colapsados)
    alvo = links.first
    for i in range(qtd):
        cand = links.nth(i)
        try:
            if cand.is_visible():
                alvo = cand
                break
        except Exception:
            continue

    antes = list(contexto.pages)
    try:
        alvo.scroll_into_view_if_needed(timeout=5000)
        alvo.click(timeout=10000)
    except Exception:
        try:
            alvo.evaluate("el => el.click()")
        except Exception as erro_click:
            raise PararPessoa("erro no Drive", f"não consegui clicar no botão do Drive ({base.erro_curto(erro_click)})")

    limite = time.time() + 10
    while time.time() < limite:
        novas = [pagina for pagina in contexto.pages if pagina not in antes]
        if novas:
            novas[-1].wait_for_load_state("domcontentloaded")
            return novas[-1], True
        time.sleep(0.3)
    return contexto.pages[-1], False


def extrair_link_drive_do_cartao(linha):
    """Lê o href direto do botão 'Abrir pasta no Drive' no cartão do MaisBook sem precisar abrir aba."""
    links = linha.locator('a[title="Abrir pasta no Drive"], a[title*="Drive" i], a[href*="drive.google.com"]')
    qtd = links.count()
    for i in range(qtd):
        try:
            endereco = (links.nth(i).get_attribute("href") or "").strip()
            if "drive.google.com" in endereco:
                return endereco
        except Exception:
            continue
    return ""


def id_da_pasta(url_ou_id):
    texto = str(url_ou_id or "").strip()
    achado = re.search(r"/folders/([a-zA-Z0-9_-]+)", texto)
    if achado:
        return achado.group(1)
    achado_id = re.search(r"[?&]id=([a-zA-Z0-9_-]+)", texto)
    if achado_id:
        return achado_id.group(1)
    if re.fullmatch(r"[a-zA-Z0-9_-]{15,}", texto):
        return texto
    return ""


def link_da_pasta(url):
    pasta_id = id_da_pasta(url)
    if not pasta_id:
        return ""
    return "https://drive.google.com/drive/folders/" + pasta_id


# ==========================================================
# Google Drive para Computador (Unidade G:\ + DriveFS SQLite)
# ==========================================================

NOMES_RAIZ_DRIVEFS = {
    "my drive",
    "meu drive",
    "root",
    "shared drives",
    "drives compartilhados",
    "team drives",
    "computers",
    "outros computadores",
}


def localizar_bancos_drivefs():
    """Localiza os bancos locais metadata_sqlite_db do Google Drive para Desktop no Windows."""
    bases = []
    local_app = os.environ.get("LOCALAPPDATA", "")
    if local_app:
        bases.append(Path(local_app) / "Google" / "DriveFS")
    bases.append(Path.home() / "AppData" / "Local" / "Google" / "DriveFS")

    bancos = []
    vistos = set()
    for base_dir in bases:
        if not base_dir.is_dir():
            continue
        try:
            for sub in base_dir.iterdir():
                if not sub.is_dir():
                    continue
                db_file = sub / "metadata_sqlite_db"
                chave = str(db_file).lower()
                if db_file.is_file() and chave not in vistos:
                    try:
                        if db_file.stat().st_size > 0:
                            vistos.add(chave)
                            bancos.append(db_file)
                    except Exception:
                        pass
        except Exception:
            continue
    bancos.sort(key=lambda p: p.stat().st_mtime if p.exists() else 0, reverse=True)
    return bancos


def raizes_drive_pc():
    """
    Retorna as pastas raízes candidatas do Google Drive no computador,
    priorizando a Unidade G:\\ padrão do Google Drive Desktop.
    """
    candidatas = []
    vistas = set()

    def _add(caminho):
        try:
            p = Path(caminho)
            chave = str(p).lower()
            if chave not in vistas and p.exists() and p.is_dir():
                vistas.add(chave)
                candidatas.append(p)
        except Exception:
            pass

    # Prioridade 1: Unidade G:\ padrão do Google Drive Desktop confirmada pelo usuário
    for sub in ("Meu Drive", "My Drive", "Drives compartilhados", "Shared drives", "Outros computadores", ""):
        _add(Path("G:\\") / sub if sub else Path("G:\\"))

    # Prioridade 2: Outras letras caso o Windows tenha montado em letra adicional
    for letra in "DEFHIJKLMNOPQRSTUVWXYZ":
        raiz = Path(f"{letra}:\\")
        try:
            if not raiz.exists():
                continue
        except Exception:
            continue
        for sub in ("Meu Drive", "My Drive", "Drives compartilhados", "Shared drives"):
            _add(raiz / sub)

    return candidatas


def _abrir_copia_sqlite_drivefs(db_path):
    """
    Faz uma cópia rápida do metadata_sqlite_db (incluindo -wal e -shm) para pasta temporária
    para ler sem travar o Google Drive Desktop e enxergar pastas recém-criadas no WAL.
    """
    pasta_tmp = Path(tempfile.mkdtemp(prefix="drivefs_ro_"))
    destino_db = pasta_tmp / "metadata_sqlite_db"
    shutil.copy2(db_path, destino_db)
    for sufixo in ("-wal", "-shm"):
        arq_aux = Path(str(db_path) + sufixo)
        if arq_aux.is_file():
            try:
                shutil.copy2(arq_aux, Path(str(destino_db) + sufixo))
            except Exception:
                pass
    conn = sqlite3.connect(str(destino_db), timeout=5.0)
    return conn, pasta_tmp


def _fechar_copia_sqlite_drivefs(conn, pasta_tmp):
    try:
        if conn:
            conn.close()
    except Exception:
        pass
    try:
        if pasta_tmp and Path(pasta_tmp).exists():
            shutil.rmtree(pasta_tmp, ignore_errors=True)
    except Exception:
        pass


def _descobrir_esquema_drivefs(conn):
    """Identifica dinamicamente as colunas de items e parents no metadata_sqlite_db do DriveFS."""
    cur = conn.cursor()
    tabelas = {
        r[0]: r[1] or ""
        for r in cur.execute("SELECT name, sql FROM sqlite_master WHERE type='table'").fetchall()
    }
    if "items" not in tabelas:
        return None

    cols_items = [r[1] for r in cur.execute("PRAGMA table_info(items)").fetchall()]
    col_id = "stable_id" if "stable_id" in cols_items else ("id" if "id" in cols_items else None)
    col_cloud = "cloud_id" if "cloud_id" in cols_items else None
    col_title = "local_title" if "local_title" in cols_items else ("title" if "title" in cols_items else None)
    col_trashed = "trashed" if "trashed" in cols_items else ("is_deleted" if "is_deleted" in cols_items else None)
    if not (col_id and col_cloud and col_title):
        return None

    tabela_pais = None
    col_filho_pai = None
    col_pai_id = None
    for cand_tab in ("parents", "item_parents", "parent_R"):
        if cand_tab in tabelas:
            cols_p = [r[1] for r in cur.execute(f"PRAGMA table_info({cand_tab})").fetchall()]
            c_filho = next((c for c in ("item_stable_id", "stable_id", "item_id", "child_stable_id", "child_id") if c in cols_p), None)
            c_pai = next((c for c in ("parent_stable_id", "parent_id") if c in cols_p), None)
            if c_filho and c_pai:
                tabela_pais = cand_tab
                col_filho_pai = c_filho
                col_pai_id = c_pai
                break

    if not tabela_pais and "parent_stable_id" in cols_items:
        tabela_pais = "items"
        col_filho_pai = col_id
        col_pai_id = "parent_stable_id"

    if not tabela_pais:
        return None

    return {
        "col_id": col_id,
        "col_cloud": col_cloud,
        "col_title": col_title,
        "col_trashed": col_trashed,
        "tabela_pais": tabela_pais,
        "col_filho_pai": col_filho_pai,
        "col_pai_id": col_pai_id,
    }


def resolver_pasta_aluno_no_drive_pc(pasta_cloud_id):
    """
    Traduz diretamente o cloud_id do botão do MaisBook para o caminho físico exato na Unidade G:\\
    consultando o índice local do Google Drive Desktop (sem varrer o Drive inteiro).
    Retorna (caminho_local_path, db_path, stable_id_aluno) ou (None, None, None).
    """
    if not pasta_cloud_id:
        return None, None, None

    raizes = raizes_drive_pc()
    bancos = localizar_bancos_drivefs()

    for db_path in bancos:
        conn, pasta_tmp = None, None
        try:
            conn, pasta_tmp = _abrir_copia_sqlite_drivefs(db_path)
            esq = _descobrir_esquema_drivefs(conn)
            if not esq:
                continue
            cur = conn.cursor()
            filtro_lixo = f" AND COALESCE({esq['col_trashed']}, 0) = 0" if esq["col_trashed"] else ""
            row = cur.execute(
                f"SELECT {esq['col_id']}, {esq['col_title']} FROM items "
                f"WHERE {esq['col_cloud']} = ?{filtro_lixo} LIMIT 1",
                (pasta_cloud_id,),
            ).fetchone()
            if not row:
                continue

            stable_id_aluno, titulo_aluno = row[0], str(row[1] or "").strip()
            partes = [titulo_aluno] if titulo_aluno else []
            atual_id = stable_id_aluno
            visitados = {atual_id}

            for _ in range(35):
                pai_row = cur.execute(
                    f"SELECT {esq['col_pai_id']} FROM {esq['tabela_pais']} WHERE {esq['col_filho_pai']} = ? LIMIT 1",
                    (atual_id,),
                ).fetchone()
                if not pai_row or pai_row[0] is None or pai_row[0] in visitados:
                    break
                pai_id = pai_row[0]
                visitados.add(pai_id)
                info_pai = cur.execute(
                    f"SELECT {esq['col_title']} FROM items WHERE {esq['col_id']} = ? LIMIT 1",
                    (pai_id,),
                ).fetchone()
                if not info_pai:
                    break
                nome_pai = str(info_pai[0] or "").strip()
                if nome_pai:
                    partes.append(nome_pai)
                atual_id = pai_id

            partes.reverse()
            # Remove prefixos virtuais de raiz ('My Drive', 'Meu Drive', 'Root', etc.)
            partes_sem_raiz = list(partes)
            while partes_sem_raiz and normalizar(partes_sem_raiz[0]) in NOMES_RAIZ_DRIVEFS:
                partes_sem_raiz.pop(0)

            for seq in (partes_sem_raiz, partes):
                if not seq:
                    continue
                for raiz in raizes:
                    candidato = raiz.joinpath(*seq)
                    if candidato.is_dir():
                        return candidato, db_path, stable_id_aluno
        except Exception:
            continue
        finally:
            _fechar_copia_sqlite_drivefs(conn, pasta_tmp)

    # Tentativa extra via diretório virtual .shortcut-targets-by-id do Google Drive Desktop no G:\
    for letra in ("G", "D", "E", "F"):
        atalho_id = Path(f"{letra}:\\.shortcut-targets-by-id") / pasta_cloud_id
        try:
            if atalho_id.is_dir():
                subpastas = [p for p in atalho_id.iterdir() if p.is_dir()]
                if len(subpastas) == 1:
                    return subpastas[0], (bancos[0] if bancos else None), None
                return atalho_id, (bancos[0] if bancos else None), None
        except Exception:
            pass

    return None, None, None


def obter_cloud_id_subpasta_drivefs(db_path, stable_id_pai, nome_subpasta, timeout_s=28):
    """
    Aguarda o Google Drive Desktop sincronizar a pasta recém-criada (ex.: ALBUM APROVAÇAO)
    no banco metadata_sqlite_db e devolve o cloud_id oficial para montar o link web.
    """
    if not db_path or not Path(db_path).is_file():
        return ""

    alvo_norm = normalizar(nome_subpasta)
    limite = time.time() + timeout_s
    while time.time() < limite:
        conn, pasta_tmp = None, None
        try:
            conn, pasta_tmp = _abrir_copia_sqlite_drivefs(db_path)
            esq = _descobrir_esquema_drivefs(conn)
            if esq:
                cur = conn.cursor()
                filtro_lixo = f" AND COALESCE(i.{esq['col_trashed']}, 0) = 0" if esq["col_trashed"] else ""
                if stable_id_pai is not None:
                    linhas = cur.execute(
                        f"SELECT i.{esq['col_cloud']}, i.{esq['col_title']} "
                        f"FROM items i "
                        f"JOIN {esq['tabela_pais']} p ON i.{esq['col_id']} = p.{esq['col_filho_pai']} "
                        f"WHERE p.{esq['col_pai_id']} = ?{filtro_lixo}",
                        (stable_id_pai,),
                    ).fetchall()
                    for cloud_id, titulo in linhas:
                        cid = str(cloud_id or "").strip()
                        if normalizar(str(titulo or "")) == alvo_norm and cid and not cid.lower().startswith("local"):
                            return cid
        except Exception:
            pass
        finally:
            _fechar_copia_sqlite_drivefs(conn, pasta_tmp)
        time.sleep(1.5)
    return ""


def copiar_arquivos_incremental_drive_pc(pasta_destino, arquivos):
    """
    Copia para a pasta do Google Drive no PC (G:\\) apenas os arquivos faltantes ou atualizados,
    e confere 100% dos arquivos e tamanhos em bytes no Windows Explorer.
    """
    pasta_destino = Path(pasta_destino)
    pasta_destino.mkdir(parents=True, exist_ok=True)
    copiados = 0
    mantidos = 0

    for arq_origem in arquivos:
        origem = Path(arq_origem)
        destino = pasta_destino / origem.name
        tam_origem = origem.stat().st_size
        precisa_copiar = True
        if destino.is_file():
            try:
                if destino.stat().st_size == tam_origem and tam_origem > 0:
                    precisa_copiar = False
            except Exception:
                precisa_copiar = True
        if precisa_copiar:
            shutil.copy2(str(origem), str(destino))
            copiados += 1
        else:
            mantidos += 1

    # Conferência rigorosa no Explorer: verifica se 100% dos arquivos existem na pasta com tamanho exato
    faltando = []
    for arq_origem in arquivos:
        origem = Path(arq_origem)
        destino = pasta_destino / origem.name
        if not destino.is_file() or destino.stat().st_size != origem.stat().st_size or destino.stat().st_size == 0:
            faltando.append(origem.name)

    if faltando:
        raise PararPessoa(
            "erro no Drive",
            f"Drive do PC (G:\\): falha ao confirmar cópia de {len(faltando)} arquivo(s): {', '.join(faltando[:5])}",
        )

    print(
        f"    Drive do PC ({pasta_destino.name}): {copiados} copiado(s), {mantidos} já atualizado(s) "
        f"— conferidos {len(arquivos)}/{len(arquivos)} arquivo(s) no Explorer"
    )


def tentar_subir_pelo_drive_pc(linha_cliente, pessoa, achados):
    """
    Tenta realizar o envio diretamente pelo Google Drive para Computador (Unidade G:\\ + DriveFS):
      1. Extrai o ID da pasta do aluno direto do botão do MaisBook
      2. Resolve instantaneamente o caminho na Unidade G:\\ via metadata_sqlite_db
      3. Confere se o nome da pasta bate com o aluno
      4. Cria ALBUM APROVAÇAO (e 20x20 / 24x34 se aplicável) e copia apenas arquivos faltantes/atualizados
      5. Obtém o link web (https://drive.google.com/drive/folders/<ID>) da pasta ALBUM APROVAÇAO
    Retorna (link, pasta_aprovacao_criada_no_pc) — se link for vazio, o chamador usa o fallback do navegador.
    """
    url_drive_cliente = extrair_link_drive_do_cartao(linha_cliente)
    id_pasta_cliente = id_da_pasta(url_drive_cliente)
    if not id_pasta_cliente:
        return "", False

    pasta_aluno_pc, db_path, stable_id_aluno = resolver_pasta_aluno_no_drive_pc(id_pasta_cliente)
    if not pasta_aluno_pc or not pasta_aluno_pc.is_dir():
        return "", False

    # Confere se o nome da pasta no G:\ bate com o nome do formando
    caminho_str = f"{pasta_aluno_pc.parent.name} {pasta_aluno_pc.name}"
    if not nomes_batem(pessoa["nome"], caminho_str) and not capa_e_da_pessoa(pessoa["nome"], pasta_aluno_pc.name):
        raise PararPessoa(
            "erro no Drive",
            f"a pasta localizada no Drive do PC ('{pasta_aluno_pc.name}') não bate com '{pessoa['nome']}' — nada foi enviado",
        )

    print(f"    pasta localizada no Drive do PC: {pasta_aluno_pc}")
    pasta_aprovacao = pasta_aluno_pc / PASTA_RAIZ
    pasta_aprovacao.mkdir(parents=True, exist_ok=True)

    if len(achados) >= 2:
        for item in achados:
            nome_sub = PASTA_TAMANHO.get(item["tamanho"], item["tamanho"].lower())
            copiar_arquivos_incremental_drive_pc(pasta_aprovacao / nome_sub, item["arquivos"])
    else:
        copiar_arquivos_incremental_drive_pc(pasta_aprovacao, achados[0]["arquivos"])

    # Aguarda o DriveFS atribuir o cloud_id da pasta ALBUM APROVAÇAO para montar o link web
    print("    aguardando o Google Drive do PC gerar o link da pasta ALBUM APROVAÇAO...")
    cloud_id_aprovacao = obter_cloud_id_subpasta_drivefs(db_path, stable_id_aluno, PASTA_RAIZ, timeout_s=28)
    if cloud_id_aprovacao:
        link = f"https://drive.google.com/drive/folders/{cloud_id_aprovacao}"
        print(f"    link obtido pelo Drive do PC: {link}")
        return link, True

    print("    arquivos já copiados no G:\\, mas o ID do link ainda não apareceu no índice local — abrindo navegador apenas para pegar o link...")
    return "", True


def subir_para_o_drive_navegador(contexto, pagina_mb, pessoa, achados, linha_cliente=None):
    if linha_cliente is None:
        _, linha_cliente = achar_cliente(pagina_mb, pessoa)
    drive, aba_nova = abrir_drive(contexto, linha_cliente)
    drive.on("dialog", no_dialogo)
    try:
        esperar_drive_carregar(drive)
        conferir_drive_da_pessoa(drive, pessoa)
        garantir_pasta(drive, PASTA_RAIZ)
        link_raiz = link_da_pasta(drive.url)
        if len(achados) >= 2:
            for item in achados:
                nome_pasta = PASTA_TAMANHO.get(item["tamanho"], item["tamanho"].lower())
                print(f"    pasta {nome_pasta}")
                garantir_pasta(drive, nome_pasta)
                colocar_arquivos(drive, item["arquivos"])
                voltar_para_pasta(drive, PASTA_RAIZ)
        else:
            colocar_arquivos(drive, achados[0]["arquivos"])
        if not link_raiz and not esta_na_pasta(drive, PASTA_RAIZ):
            voltar_para_pasta(drive, PASTA_RAIZ)
        link = link_raiz or link_da_pasta(drive.url)
        if not link:
            raise PararPessoa("erro no Drive", "não consegui o link da pasta ALBUM APROVAÇAO")
        print(f"    link: {link}")
        return link
    finally:
        if aba_nova:
            try:
                drive.close()
            except Exception:
                pass
        else:
            try:
                pagina_mb.go_back(timeout=15000)
            except Exception:
                pass
        try:
            pagina_mb.bring_to_front()
        except Exception:
            pass


def subir_para_o_drive(contexto, pagina_mb, pessoa, achados):
    """
    Envia as fotos prioritariamente pelo Google Drive do Computador (Unidade G:\\ + DriveFS).
    Se a pasta do aluno ainda não estiver sincronizada no PC ou o link não puder ser lido localmente,
    cai automaticamente para o envio pelo navegador sem quebrar o fluxo.
    """
    _, linha = achar_cliente(pagina_mb, pessoa)
    try:
        link_pc, copiou_no_pc = tentar_subir_pelo_drive_pc(linha, pessoa, achados)
        if link_pc:
            return link_pc
        if copiou_no_pc:
            print("    usando fallback do navegador para capturar o link da pasta ALBUM APROVAÇAO...")
    except PararPessoa:
        raise
    except Exception as erro_pc:
        print(f"    aviso Drive do PC: {base.erro_curto(erro_pc)} — usando navegador como fallback...")

    return subir_para_o_drive_navegador(contexto, pagina_mb, pessoa, achados, linha_cliente=linha)


def abrir_edicao(pagina, pessoa):
    cliente_id, linha = achar_cliente(pagina, pessoa)
    botao_editar = pagina.locator(
        f'button[title="Editar"][data-cliente-id="{cliente_id}"], '
        f'[title*="Editar" i][data-cliente-id="{cliente_id}"]'
    )
    if not botao_editar.count():
        botao_editar = linha.locator('button[title*="Editar" i], [title*="Editar" i]')
    botao_editar.first.scroll_into_view_if_needed()
    botao_editar.first.click(timeout=15000)
    modal = pagina.locator("#modal-cliente")
    modal.wait_for(state="visible", timeout=15000)
    campo_sge = pagina.locator("#cliente_codigo_interno")
    campo_nome = pagina.locator("#cliente_nome")
    limite = time.time() + 15
    confirmado = False
    sge_modal_ultimo, nome_modal_ultimo = "", ""
    while time.time() < limite:
        sge_modal = (campo_sge.input_value() or "").strip() if campo_sge.count() else ""
        nome_modal = (campo_nome.input_value() or "").strip() if campo_nome.count() else ""
        sge_modal_ultimo, nome_modal_ultimo = sge_modal, nome_modal
        nome_ok = bool(nome_modal) and nome_cartao_bate(pessoa["nome"], nome_modal)
        cartao_modal = {
            "sge": sge_modal or pessoa.get("sge_maisbook", ""),
            "codigo": pessoa.get("codigo_maisbook", ""),
        }
        codigo_ok = codigo_cartao_bate(pessoa, cartao_modal)
        if nome_ok and codigo_ok:
            if sge_modal and sge_modal != pessoa["sge"]:
                pessoa["sge_maisbook"] = sge_modal
            confirmado = True
            break
        time.sleep(0.3)
    if not confirmado:
        fechar_edicao(pagina)
        raise PararPessoa(
            "erro no e-mail",
            f"o Editar abriu outro cliente (nome='{nome_modal_ultimo}', código='{sge_modal_ultimo}') — nome e código precisam bater",
        )
    return modal


def campo_album_pronto(pagina):
    campo = pagina.locator("#cliente_album_link")
    campo.wait_for(state="visible", timeout=10000)
    anterior, desde = None, time.time()
    limite = time.time() + 12
    while time.time() < limite:
        atual = campo.input_value()
        if atual != anterior:
            anterior, desde = atual, time.time()
        elif time.time() - desde >= 1.5:
            return campo
        time.sleep(0.3)
    return campo


def edicao_aberta(pagina):
    modal = pagina.locator("#modal-cliente.active, #modal-cliente:visible")
    try:
        return modal.count() > 0 and modal.first.is_visible()
    except Exception:
        return False


def fechar_edicao(pagina):
    if not edicao_aberta(pagina):
        return
    try:
        btn_cancelar = pagina.locator("#modal-cliente").get_by_role("button", name="Cancelar", exact=True)
        if btn_cancelar.count() and btn_cancelar.first.is_visible():
            btn_cancelar.first.click(timeout=5000)
        else:
            btn_fechar = pagina.locator("#modal-cliente button.modal-close, #modal-cliente .close")
            if btn_fechar.count() and btn_fechar.first.is_visible():
                btn_fechar.first.click(timeout=5000)
            else:
                pagina.keyboard.press("Escape")
    except Exception:
        try:
            pagina.keyboard.press("Escape")
        except Exception:
            pass
    try:
        pagina.locator("#modal-cliente.active").wait_for(state="hidden", timeout=8000)
    except Exception:
        pass


def esperar_salvar(pagina):
    inicio = time.time()
    while time.time() - inicio < 25:
        if not edicao_aberta(pagina):
            return
        botao = pagina.locator("#modal-cliente button[type='submit']")
        try:
            texto = normalizar(botao.first.inner_text(timeout=1000)) if botao.count() else ""
        except Exception:
            texto = ""
        if "salvando" not in texto and time.time() - inicio > 2:
            fechar_edicao(pagina)
            return
        time.sleep(0.4)
    fechar_edicao(pagina)
    raise PararPessoa("erro no e-mail", "o MaisBook não terminou de salvar o Álbum pronto")


def colar_link_album(pagina, pessoa, link):
    modal = abrir_edicao(pagina, pessoa)
    campo = campo_album_pronto(pagina)
    campo.fill(link)
    if link_da_pasta(campo.input_value()) != link:
        fechar_edicao(pagina)
        raise PararPessoa("erro no e-mail", "não consegui colar o link em Álbum pronto")
    modal.get_by_role("button", name="Salvar", exact=True).first.click(timeout=15000)
    esperar_salvar(pagina)
    pagina.wait_for_timeout(1200)

    ir_clientes(pagina, termo_busca(pessoa))
    abrir_edicao(pagina, pessoa)
    gravado = pagina.locator("#cliente_album_link").input_value()
    fechar_edicao(pagina)
    if link_da_pasta(gravado) != link:
        raise PararPessoa("erro no e-mail", "salvei, mas o link não ficou em Álbum pronto")
    print("    link salvo em Álbum pronto")
    return link


def enviar_email(pagina, pessoa):
    fechar_edicao(pagina)
    cliente_id, linha = achar_cliente(pagina, pessoa)
    album_mb = normalizar(pessoa.get("album_maisbook") or "")
    ja_enviado_mb = "aguardando aprovacao" in album_mb or "aprovado" in album_mb

    btn_email = pagina.locator(
        f'button[title="Enviar e-mail"][data-cliente-id="{cliente_id}"], '
        f'button[title*="e-mail" i][data-cliente-id="{cliente_id}"], '
        f'button[title*="email" i][data-cliente-id="{cliente_id}"], '
        f'button[title*="notifica" i][data-cliente-id="{cliente_id}"], '
        f'button[title*="reenviar" i][data-cliente-id="{cliente_id}"]'
    )
    if not btn_email.count():
        btn_email = linha.locator(
            'button[title*="e-mail" i], button[title*="email" i], '
            'button[title*="notifica" i], button[title*="reenviar" i], '
            'button:has(.dashicons-email), button:has(svg.fa-envelope), button:has(i[class*="envelope"])'
        )
    if not btn_email.count():
        if ja_enviado_mb:
            return f"E-mail já consta como enviado no MaisBook ({pessoa.get('album_maisbook')})"
        raise PararPessoa("erro no e-mail", "não achei o botão Enviar e-mail neste cliente")
    try:
        btn_email.first.scroll_into_view_if_needed()
        btn_email.first.click(timeout=15000)
        modal = pagina.locator("#modal-notificacao-cliente")
        modal.wait_for(state="visible", timeout=15000)
    except Exception:
        if ja_enviado_mb:
            return f"E-mail já consta como enviado no MaisBook ({pessoa.get('album_maisbook')})"
        raise
    info = modal.locator("#notificacao_cliente_info").inner_text()
    if pessoa["nome"] and not nomes_batem(pessoa["nome"], info):
        try:
            modal.locator("button.modal-close").first.click(timeout=5000)
        except Exception:
            pagina.keyboard.press("Escape")
        raise PararPessoa("erro no e-mail", "o pop-up abriu outro cliente — e-mail não enviado")

    pagina.locator("#notificacao_template").evaluate(
        """(el, valor) => {
          el.value = valor;
          el.dispatchEvent(new Event('change', { bubbles: true }));
          if (window.jQuery) window.jQuery(el).val(valor).trigger('change');
        }""",
        TEMPLATE_VALOR,
    )
    valor = pagina.locator("#notificacao_template").input_value()
    assunto = pagina.locator("#notificacao_assunto").input_value().strip()
    if valor != TEMPLATE_VALOR or not assunto:
        try:
            modal.locator("button.modal-close").first.click(timeout=5000)
        except Exception:
            pagina.keyboard.press("Escape")
        raise PararPessoa("erro no e-mail", "não consegui selecionar o template Aprovação de álbum")

    pagina.locator("#btn-enviar-notificacao").click(timeout=15000)
    resultado = pagina.locator("#notificacao_result")
    resultado.wait_for(state="visible", timeout=45000)
    info_resultado = resultado.evaluate(
        """el => ({ texto: el.innerText || '', fundo: getComputedStyle(el).backgroundColor || '' })"""
    )
    texto = normalizar(info_resultado.get("texto"))
    verde = "209, 250, 229" in info_resultado.get("fundo", "")
    try:
        fechar_notif = modal.locator("button.modal-close")
        if fechar_notif.count() and fechar_notif.first.is_visible():
            fechar_notif.first.click(timeout=3000)
    except Exception:
        pass
    if "erro" in texto and "sucesso" not in texto:
        raise PararPessoa("erro no e-mail", info_resultado.get("texto") or "o MaisBook recusou o envio")
    if "sucesso" in texto or "enviado" in texto or verde:
        return info_resultado.get("texto") or "E-mail enviado com sucesso."
    raise PararPessoa("erro no e-mail", "o MaisBook não confirmou o envio")


# ==========================================================
# Atualização da Coluna K (Aguardando Aprovação) + Baixa SGE
# ==========================================================

ORDEM_COLUNA_K = [
    "Sem álbum",
    "AGUARDANDO FOTO",
    "EM DIAGRAMAÇÃO",
    "DIAGRAMADO",
    "Aguardando Aprovação",
    "Com Modificações",
    "Aprovado",
    "IMPRESSÃO",
    "Finalizado",
]


def indice_etapa_k(valor):
    alvo = normalizar(valor)
    for idx, etapa in enumerate(ORDEM_COLUNA_K):
        if normalizar(etapa) == alvo:
            return idx
    return -1


def atualizar_status_planilha(contexto, pessoa, novo_status="Aguardando Aprovação"):
    """
    Atualiza a Coluna K (STATUS) da pessoa para 'Aguardando Aprovação' na planilha do Google
    (e na cópia local), respeitando a regra de que a Coluna K só anda para baixo.
    """
    from openpyxl.utils import get_column_letter

    aba = pessoa.get("aba") or "GERAL ENVIOS"
    linha_num = int(pessoa.get("linha") or 0)
    col_status = int(pessoa.get("col_status") or 11)
    if linha_num < 1:
        return "linha inválida"

    # 1. Atualiza na cópia local .xlsx se existir
    try:
        if PLANILHA.is_file():
            wb = proc.load_workbook(PLANILHA)
            if aba in wb.sheetnames:
                ws = wb[aba]
                atual_local = ws.cell(row=linha_num, column=col_status).value
                idx_atual = indice_etapa_k(atual_local)
                idx_novo = indice_etapa_k(novo_status)
                if idx_atual < idx_novo and idx_atual not in (7, 8):
                    ws.cell(row=linha_num, column=col_status, value=novo_status)
                    wb.save(PLANILHA)
            wb.close()
    except Exception:
        pass

    # 2. Atualiza no Google Sheets usando o navegador já autenticado na conta Google
    letra_col = get_column_letter(col_status)
    ref_celula = f"{letra_col}{linha_num}"
    plan_pag = None
    criou_aba = False
    try:
        for pag in contexto.pages:
            if "/spreadsheets/d/" in (pag.url or ""):
                plan_pag = pag
                break
        if plan_pag is None:
            plan_pag = contexto.new_page()
            criou_aba = True
            plan_pag.goto(PLANILHA_LINK, wait_until="domcontentloaded", timeout=60000)

        plan_pag.wait_for_function(
            "() => { const el = document.querySelector('#t-name-box'); return el && !el.hasAttribute('disabled'); }",
            timeout=30000,
        )
        caixa = plan_pag.locator("#t-name-box").first
        caixa.click()
        caixa.fill("'" + aba.replace("'", "''") + "'!" + ref_celula)
        caixa.press("Enter")
        plan_pag.wait_for_timeout(700)

        barra_formula = plan_pag.locator("#t-formula-bar-input").first
        try:
            atual_barra = (
                barra_formula
                .inner_text(timeout=3000)
                .replace("\ufeff", "")
                .replace("\u200b", "")
                .strip()
            )
        except Exception:
            atual_barra = ""

        idx_atual = indice_etapa_k(atual_barra)
        idx_novo = indice_etapa_k(novo_status)
        if idx_atual in (7, 8) or (idx_atual >= idx_novo and idx_atual != -1):
            print(f"    planilha: Coluna K ({ref_celula}) já está em '{atual_barra}' (não retrocede)")
            return f"Coluna K mantida em {atual_barra}"

        barra_formula.click()
        plan_pag.keyboard.press("Control+A")
        plan_pag.keyboard.press("Backspace")
        plan_pag.keyboard.type(novo_status, delay=15)
        plan_pag.keyboard.press("Enter")
        plan_pag.wait_for_timeout(2500)
        print(f"    planilha: Coluna K ({ref_celula}) atualizada para '{novo_status}'")
        try:
            gravar_edicoes_planilha(
                aba=aba,
                linha_num=linha_num,
                celulas_dict={letra_col: novo_status},
                sge=pessoa.get("sge", ""),
                usar_navegador=False,
            )
        except Exception:
            pass
        return f"Coluna K ({ref_celula}) -> {novo_status}"
    except Exception as erro:
        aviso = f"não consegui editar a célula {ref_celula} direto no Google Sheets ({base.erro_curto(erro)})"
        print(f"    aviso planilha: {aviso}")
        return aviso
    finally:
        if criou_aba and plan_pag is not None:
            try:
                plan_pag.close()
            except Exception:
                pass


def gravar_edicoes_planilha(aba, linha_num, celulas_dict, sge="", webhook_url="", usar_navegador=True):
    """
    Reescreve em tempo real uma ou mais células de uma linha da planilha:
      1. Salva no overlay instantâneo (edicoes_planilha_local.json)
      2. Atualiza todas as cópias locais .xlsx (planilha_envio_aprovacao.xlsx, planilha_organizador.xlsx, baixa_planilha.xlsx)
      3. Envia via Webhook Google Apps Script (se configurado)
      4. Abre/usa Chromium em background (Playwright) para gravar direto nas células do Google Sheets online
    """
    import urllib.request
    from openpyxl.utils import column_index_from_string, get_column_letter

    aba = (aba or "GERAL ENVIOS").strip()
    linha_num = int(linha_num or 0)
    if linha_num < 1 or not celulas_dict:
        return {"ok": False, "mensagem": "Linha ou células inválidas"}

    normalizadas = {}
    for col_ref, valor in celulas_dict.items():
        if str(col_ref).isdigit():
            letra = get_column_letter(int(col_ref))
        else:
            letra = str(col_ref).strip().upper()
        normalizadas[letra] = "" if valor is None else str(valor)

    # 1. Grava no overlay local em tempo real (0ms)
    try:
        dados_locais = carregar_json(EDICOES_LOCAIS)
        linhas_map = dados_locais.setdefault("linhas", {})
        chave = f"{aba}:{linha_num}"
        atual = linhas_map.get(chave, {"aba": aba, "linha": linha_num, "sge": str(sge or ""), "celulas": {}})
        atual["aba"] = aba
        atual["linha"] = linha_num
        if sge:
            atual["sge"] = str(sge)
        atual.setdefault("celulas", {}).update(normalizadas)
        atual["ts"] = time.time()
        atual["data"] = f"{datetime.now():%Y-%m-%d %H:%M:%S}"
        linhas_map[chave] = atual
        salvar_json(EDICOES_LOCAIS, dados_locais)
    except Exception as e:
        print(f"    aviso overlay local: {e}")

    # 2. Atualiza as cópias locais .xlsx imediatamente
    arquivos_xlsx = [
        PLANILHA,
        PASTA / "planilha_organizador.xlsx",
        PASTA.parent / "baixa_sge" / "baixa_planilha.xlsx",
    ]
    for arq in arquivos_xlsx:
        if not arq.is_file():
            continue
        try:
            wb = proc.load_workbook(arq)
            if aba in wb.sheetnames:
                ws = wb[aba]
                for letra, val in normalizadas.items():
                    idx = column_index_from_string(letra)
                    ws.cell(row=linha_num, column=idx, value=val)
                wb.save(arq)
            wb.close()
        except Exception:
            pass

    mensagens = [f"Local (.xlsx) atualizado ({', '.join(f'{k}{linha_num}={v!r}' for k, v in normalizadas.items())})"]

    # 3. Se houver Webhook do Google Apps Script configurado, envia POST instantâneo
    if not webhook_url:
        cfg_org = carregar_json(PASTA / "organizador_config.json")
        webhook_url = (cfg_org.get("webhook_apps_script") or "").strip()

    if webhook_url and webhook_url.startswith("http"):
        try:
            payload = json.dumps({
                "action": "update_cells",
                "spreadsheetId": "1FeYQop9z6rVAFQuIs1EOiuQcE84t0iBgamYtMWUIAU0",
                "sheetName": aba,
                "aba": aba,
                "row": linha_num,
                "linha": linha_num,
                "sge": str(sge or ""),
                "cells": normalizadas,
                "celulas": normalizadas,
            }, ensure_ascii=False).encode("utf-8")
            req = urllib.request.Request(
                webhook_url,
                data=payload,
                headers={"Content-Type": "application/json"},
                method="POST",
            )
            with urllib.request.urlopen(req, timeout=10) as resp:
                resp.read()
            mensagens.append("Google Sheets atualizado via Webhook Apps Script")
        except Exception as e:
            mensagens.append(f"Aviso Webhook: {e}")

    # 4. Grava direto nas células do Google Sheets online via Playwright (headless em background)
    if usar_navegador:
        try:
            from playwright.sync_api import sync_playwright

            with sync_playwright() as p:
                try:
                    navegador = p.chromium.launch(headless=True, channel="chrome")
                except Exception:
                    navegador = p.chromium.launch(headless=True)
                plan_pag = navegador.new_page()
                try:
                    plan_pag.goto(PLANILHA_LINK, wait_until="domcontentloaded", timeout=60000)
                    plan_pag.wait_for_function(
                        "() => { const el = document.querySelector('#t-name-box'); return el && !el.hasAttribute('disabled'); }",
                        timeout=30000,
                    )
                    caixa = plan_pag.locator("#t-name-box").first
                    barra_formula = plan_pag.locator("#t-formula-bar-input").first
                    gravadas_ok = []
                    for letra, val in normalizadas.items():
                        ref_celula = f"{letra.upper()}{linha_num}"
                        caixa.click()
                        caixa.fill("'" + aba.replace("'", "''") + "'!" + ref_celula)
                        caixa.press("Enter")
                        plan_pag.wait_for_timeout(600)
                        celula_atual = (caixa.input_value() or "").strip().upper()
                        if celula_atual != ref_celula:
                            raise RuntimeError(
                                f"Linha {linha_num} está oculta por filtro no Google Sheets (cursor foi para {celula_atual})"
                            )
                        if val:
                            barra_formula.click()
                            plan_pag.keyboard.press("Control+A")
                            plan_pag.keyboard.press("Backspace")
                            plan_pag.keyboard.type(val, delay=15)
                            plan_pag.keyboard.press("Enter")
                        else:
                            plan_pag.keyboard.press("Backspace")
                            plan_pag.keyboard.press("Delete")
                            plan_pag.keyboard.press("Enter")
                        plan_pag.wait_for_timeout(800)
                        gravadas_ok.append(ref_celula)
                    # Aguarda o salvamento XHR do Google Sheets no Drive antes de fechar
                    plan_pag.wait_for_timeout(2500)
                    mensagens.append(f"Google Sheets Online atualizado ({', '.join(gravadas_ok)})")
                finally:
                    try:
                        navegador.close()
                    except Exception:
                        pass
        except Exception as e:
            mensagens.append(f"Aviso Google Sheets Online: {e}")

    resumo = " | ".join(mensagens)
    print("REESCRITA PLANILHA:", resumo)
    return {"ok": True, "mensagem": resumo}


def dar_baixa_sge_enviado(playwright, pessoa, estado_sge):
    """
    Aciona o robô do SGE (baixa_sge/BAIXA.PY) para registrar 'Enviado para Aprovação'
    para o formando assim que o álbum é enviado no MaisBook.
    """
    try:
        arq_baixa = PASTA.parent / "baixa_sge" / "BAIXA.PY"
        sge_mod = estado_sge.get("mod") or SourceFileLoader("baixa_sge_mod", str(arq_baixa)).load_module()

        if estado_sge.get("robo") is None:
            sge_cfg = sge_mod.carregar_config()
            sge_cfg["aba"] = "GERAL ENVIOS"
            sge_contexto, sge_pagina = sge_mod.abrir_navegador(playwright)
            robo_sge = sge_mod.Robo(sge_contexto, sge_pagina, sge_cfg, True, False)
            sge_pagina.on("dialog", robo_sge._dialogo)
            contas_sge = sge_mod.ContasSGE()
            robo_sge.contas = contas_sge
            sge_mod.ir(sge_pagina, sge_cfg["url_clientes"], robo_sge.timeout)
            sge_mod.entrar_sge(robo_sge, contas_sge, PARAR)
            estado_sge["robo"] = robo_sge
            estado_sge["contas"] = contas_sge
            estado_sge["contexto"] = sge_contexto
            estado_sge["cfg"] = sge_cfg
            estado_sge["mod"] = sge_mod

        robo_sge = estado_sge["robo"]
        contas_sge = estado_sge["contas"]
        sge_cfg = estado_sge["cfg"]
        sge_mod = estado_sge.get("mod") or sge_mod
        robo_sge.contas = contas_sge

        planos_lista = pessoa.get("planos") or []
        item_aluno = planos_lista[0] if planos_lista else ", ".join(pessoa.get("tamanhos") or [])

        codigos_tentar = [str(pessoa.get("sge") or "").strip()]
        sge_mb = str(pessoa.get("sge_maisbook") or "").strip()
        if sge_mb and sge_mb not in codigos_tentar:
            codigos_tentar.append(sge_mb)

        linhas_csv_sge = []
        resultados_sge = []
        aluno_sge = None

        for cod_sge in codigos_tentar:
            if not cod_sge:
                continue
            aluno_sge = {
                "linha": int(pessoa.get("linha") or 0),
                "sge": sge_mod.codigo_base(cod_sge, sge_cfg),
                "nome": pessoa.get("nome") or "",
                "status": "Aguardando Aprovação",
                "item": item_aluno,
                "cpf": pessoa.get("cpf") or "",
                "maisbook": pessoa.get("codigo") or "",
            }
            linhas_csv_sge = []
            resultados_sge = sge_mod.processar_com_contas(robo_sge, aluno_sge, linhas_csv_sge)
            if resultados_sge and not all(r == "NAO ENCONTRADO" for r in resultados_sge):
                break

        status_sge = resultados_sge[0] if resultados_sge else "SEM RETORNO"
        det_sge = linhas_csv_sge[0][7] if linhas_csv_sge and len(linhas_csv_sge[0]) > 7 else ""

        if any(r in ("ALTERADO", "REENVIADO", "JA NO STATUS") for r in resultados_sge) and aluno_sge:
            feitas_sge = sge_mod.carregar_feitas()
            chave_feita = (
                f"{aluno_sge['sge']}|"
                f"{sge_mod.normalizar(sge_mod.status_desejado(sge_cfg, aluno_sge['status']))}|"
                f"{sge_mod.normalizar(aluno_sge['item'])}"
            )
            feitas_sge[chave_feita] = {
                "nome": pessoa["nome"],
                "status": "Enviado para Aprovação",
                "cpf": pessoa.get("cpf", ""),
                "data": f"{datetime.now():%Y-%m-%d %H:%M:%S}",
            }
            sge_mod.salvar_feitas(feitas_sge)

        resumo_sge = f"SGE ({aluno_sge['sge'] if aluno_sge else pessoa['sge']}): {status_sge} (Enviado para Aprovação)" + (
            f" — {det_sge}" if det_sge else ""
        )
        print(f"    {resumo_sge}")
        return resumo_sge
    except Exception as erro:
        aviso = f"SGE erro: {base.erro_curto(erro) if hasattr(base, 'erro_curto') else erro}"
        print(f"    {aviso}")
        return aviso


# ==========================================================
# Execução e Modo Vigia Contínuo
# ==========================================================

def linha_relatorio(pessoa, achados, resultado, mensagem):
    achados = achados or []
    avisos = list(pessoa.get("avisos") or [])
    if pessoa.get("duplicados"):
        avisos.append("pesquisa trouxe mais de um cartão: " + "; ".join(pessoa["duplicados"]))
    return {
        "NOME": pessoa["nome"],
        "SGE": pessoa["sge"],
        "CONTRATO": pessoa.get("contrato", ""),
        "CODIGO": pessoa.get("codigo", ""),
        "COLUNA_M_CAPA": pessoa.get("capa_planilha") or "(vazia)",
        "TAMANHOS": ", ".join(pessoa["tamanhos"]),
        "JPG": texto_jpg(achados),
        "CAPA": texto_capa(achados),
        "RESULTADO": resultado,
        "MENSAGEM": mensagem,
        "AVISO": " | ".join(avisos),
    }


def pediu_parar():
    return PARAR.exists()


def analisar_argumentos():
    parser = argparse.ArgumentParser(description="Envia o álbum diagramado para aprovação no MaisBook")
    parser.add_argument("--sge", help="processa só este código")
    parser.add_argument("--limite", type=int)
    parser.add_argument("--origem", help="pasta raiz do servidor onde estão CAPAS e as REMESSAS de álbuns")
    parser.add_argument("--simular", action="store_true", help="procura e abre o cliente, sem pasta, sem arquivo e sem e-mail")
    parser.add_argument("--continuar", action="store_true", help="retoma quem ainda não foi enviado")
    parser.add_argument("--listar", action="store_true", help="só lista quem está com STATUS DIAGRAMADO e o status da Coluna M (CAPA)")
    parser.add_argument("--diagnostico", action="store_true", help="exibe diagnóstico detalhado do funil de elegibilidade do envio em lote")
    parser.add_argument("--diagnostico-json", action="store_true", help="retorna JSON com o diagnóstico detalhado de elegibilidade")
    parser.add_argument("--conferir-servidor", action="store_true", help="confere fisicamente no servidor se cada formando tem CAPA e ÁLBUM completos")
    parser.add_argument("--etapa", default="DIAGRAMADO", help="etapa da Coluna K a ser conferida no servidor (padrão: DIAGRAMADO)")
    parser.add_argument("--conferir-json", help="JSON de entrada ou caminho para conferir lista de alunos no servidor e devolver JSON")
    parser.add_argument("--vigia", action="store_true", help="monitora continuamente a planilha e envia automaticamente quem tiver STATUS=DIAGRAMADO e Coluna M (CAPA) liberada")
    parser.add_argument("--intervalo", type=int, default=60, help="segundos entre cada verificação no modo --vigia (padrão: 60)")
    parser.add_argument("--ignorar-coluna-m", action="store_true", help="permite enviar todos os DIAGRAMADO mesmo se a Coluna M (CAPA) estiver vazia")
    parser.add_argument("--sem-planilha", action="store_true", help="não atualiza a Coluna K para Aguardando Aprovação após enviar")
    parser.add_argument("--sem-sge", action="store_true", help="não dá baixa automática no SGE (Enviado para Aprovação) após enviar")
    parser.add_argument("--gravar-celulas", help="JSON com {'aba', 'linha', 'sge', 'celulas', 'usar_navegador'} para reescrever células na planilha")
    return parser.parse_args()


def executar_rodada(args, em_vigia=False):
    feitos = carregar_json(FEITOS)
    pessoas = ler_diagramados()
    if args.sge:
        codigo = limpar_sge(args.sge)
        pessoas = [pessoa for pessoa in pessoas if pessoa["sge"] == codigo]
        if not pessoas:
            print(f"O código {args.sge} não está com STATUS DIAGRAMADO.")
            diag = diagnosticar_elegibilidade_lote(
                origem_srv=getattr(args, "origem", None),
                sge_filtro=codigo,
                ignorar_col_m=getattr(args, "ignorar_coluna_m", False),
            )
            imprimir_diagnostico_elegibilidade(diag)
            return 0
    if args.listar:
        liberados = 0
        aguardando_m = 0
        for pessoa in pessoas:
            col_m = pessoa.get("capa_planilha") or "(vazia)"
            if pessoa["sge"] in feitos:
                marca = "JÁ ENVIADO"
            elif pessoa.get("liberado_coluna_m"):
                marca = f"LIBERADO PELA COLUNA M (CAPA='{col_m}') — VAI CONFERIR NO SERVIDOR E ENVIAR"
                liberados += 1
            else:
                marca = f"AGUARDANDO OK NA COLUNA M (CAPA={col_m})"
                aguardando_m += 1
            print(
                f"linha {pessoa['linha']}: SGE {pessoa['sge']} — {pessoa['nome']} | "
                f"código MB {pessoa.get('codigo') or '-'} | "
                f"{', '.join(pessoa['tamanhos']) or '(plano não identificado)'} | {marca}"
            )
        print(
            f"\nTotal com STATUS DIAGRAMADO (Coluna K): {len(pessoas)} | "
            f"Liberados na Coluna M (CAPA) pendentes: {liberados} | "
            f"Retidos sem OK na Coluna M: {aguardando_m}"
        )
        return len(pessoas)

    ja_enviadas = [pessoa for pessoa in pessoas if pessoa["sge"] in feitos]
    pessoas = [pessoa for pessoa in pessoas if pessoa["sge"] not in feitos]
    if ja_enviadas:
        print(f"Já enviadas, sem mandar de novo ({len(ja_enviadas)}): {', '.join(pessoa['sge'] for pessoa in ja_enviadas)}")

    # REGRA COLUNA M (CAPA): Só envia realmente quem estiver com a Coluna M (CAPA) marcada (ex.: DIAGRAMADO / OK / CAPA ENTREGUE)
    # a menos que o usuário passe --ignorar-coluna-m ou tenha pedido um --sge específico.
    exigir_col_m = not getattr(args, "ignorar_coluna_m", False)
    retidos_coluna_m = []
    if exigir_col_m:
        liberados_m = []
        for pessoa in pessoas:
            if pessoa.get("liberado_coluna_m"):
                liberados_m.append(pessoa)
            else:
                retidos_coluna_m.append(pessoa)
        pessoas = liberados_m

    if retidos_coluna_m:
        print(
            f"Retidos aguardando OK na Coluna M (CAPA vazia/em produção) ({len(retidos_coluna_m)}): "
            + ", ".join(f"{p['nome']} ({p['sge']})" for p in retidos_coluna_m[:10])
            + (f" ... e mais {len(retidos_coluna_m) - 10}" if len(retidos_coluna_m) > 10 else "")
        )

    if args.continuar:
        print(f"Continuando com {len(pessoas)} pessoa(s) liberada(s) na Coluna M (CAPA).")
    if args.limite:
        pessoas = pessoas[: args.limite]

    print("=" * 70)
    print(
        "ENVIO DE ÁLBUM PARA APROVAÇÃO —",
        "SIMULAÇÃO" if args.simular else ("MODO VIGIA CONTÍNUO" if em_vigia else "ENVIO REAL"),
    )
    print(
        f"Liberados na Coluna M (CAPA) para conferir no servidor e enviar: {len(pessoas)} "
        f"| Aguardando OK na Coluna M: {len(retidos_coluna_m)}"
    )
    print("=" * 70)
    if not pessoas:
        print("Nenhuma pessoa liberada na Coluna M (CAPA) para processar agora.")
        diag = diagnosticar_elegibilidade_lote(
            origem_srv=getattr(args, "origem", None),
            sge_filtro=getattr(args, "sge", "") or "",
            ignorar_col_m=getattr(args, "ignorar_coluna_m", False),
        )
        imprimir_diagnostico_elegibilidade(diag)
        return 0

    if not em_vigia:
        PARAR.unlink(missing_ok=True)
    print("\n" + "=" * 70)
    print("FASE 1 — CONFERÊNCIA DE SEGURANÇA NO SERVIDOR (capa e lâminas JPG)")
    print("(Mesmo com OK na Coluna M, confere se capa e álbum estão completos no servidor antes de enviar)")
    print("=" * 70)
    albuns, indice = preparar_indice(getattr(args, "origem", None))
    prontos = []
    fora = []
    linhas = []
    for pessoa in pessoas:
        if pediu_parar():
            print("\nPARADO A PEDIDO antes de procurar o restante.")
            break
        col_m_txt = pessoa.get("capa_planilha") or "OK"
        print(
            f"\n{pessoa['nome']} — SGE {pessoa['sge']} — Coluna M (CAPA)='{col_m_txt}' — "
            f"{', '.join(pessoa['tamanhos']) or pessoa['planos']}"
        )
        achados, mensagem, resultado = localizar_albuns(pessoa, albuns, indice)
        if resultado:
            pessoa.setdefault("avisos", []).append(
                f"Coluna M (CAPA='{col_m_txt}') estava marcada para enviar, mas faltou arquivo no servidor — envio bloqueado por segurança"
            )
            for parte in mensagem.split("; "):
                if parte.strip():
                    print(
                        f"NÃO ENCONTREI: {pessoa['nome']} (SGE {pessoa['sge']}) — "
                        f"Coluna M (CAPA='{col_m_txt}') tinha OK, mas no servidor: {parte.strip()} [ENVIO BLOQUEADO]"
                    )
            print(f"    {resultado} (envio bloqueado mesmo com OK na Coluna M) — {mensagem}")
            linhas.append(linha_relatorio(pessoa, achados, resultado, mensagem))
            salvar_relatorio(linhas)
            fora.append((pessoa, mensagem))
            continue
        print("    CONFERIDO NO SERVIDOR OK: " + texto_jpg(achados) + " | capa " + texto_capa(achados))
        prontos.append((pessoa, achados))

    print("\n" + "-" * 70)
    print(f"CONFIRMADOS (COLUNA M OK + SERVIDOR COMPLETO) — VÃO PARA O MAISBOOK ({len(prontos)}):")
    for pessoa, _ in prontos:
        print(f"    {pessoa['nome']} — SGE {pessoa['sge']} — CAPA='{pessoa.get('capa_planilha')}' — {', '.join(pessoa['tamanhos'])}")
    print(f"BLOQUEADOS NA CONFERÊNCIA DO SERVIDOR (ERRO AO DAR OK NA COLUNA M) ({len(fora)}):")
    for pessoa, motivo in fora:
        print(f"    {pessoa['nome']} — SGE {pessoa['sge']} — CAPA='{pessoa.get('capa_planilha')}' — {motivo}")
    print("-" * 70)

    if not prontos:
        print("\nNinguém com capa e álbum completos — o MaisBook não foi aberto.")
        print(f"RELATORIO: {RELATORIO}")
        return 0

    from playwright.sync_api import sync_playwright

    print("\n" + "=" * 70)
    print("FASE 2 — MAISBOOK + PLANILHA (Coluna K) + SGE (um aluno por vez)")
    print("=" * 70)
    enviados_rodada = 0
    with sync_playwright() as p:
        contexto = abrir_navegador(p)
        pagina = contexto.pages[0] if contexto.pages else contexto.new_page()
        pagina.on("dialog", no_dialogo)
        estado_sge = {"robo": None, "contas": None, "contexto": None, "cfg": None}
        try:
            for numero, (pessoa, achados) in enumerate(prontos, start=1):
                if pediu_parar():
                    print(f"\nPARADO A PEDIDO antes de {pessoa['nome']}.")
                    break
                print(
                    f"\n[{numero}/{len(prontos)}] MaisBook — {pessoa['nome']} ({pessoa['sge']}) "
                    f"| contrato {pessoa.get('contrato') or '-'} | pesquisa {termo_busca(pessoa)}"
                )
                try:
                    ir_clientes(pagina, termo_busca(pessoa))
                    if args.simular:
                        achar_cliente(pagina, pessoa)
                        print("    simulado — cliente encontrado, sem pasta e sem e-mail")
                        linhas.append(linha_relatorio(pessoa, achados, "simulado", "cliente encontrado no MaisBook"))
                    else:
                        link = subir_para_o_drive(contexto, pagina, pessoa, achados)
                        ir_clientes(pagina, termo_busca(pessoa))
                        colar_link_album(pagina, pessoa, link)
                        ir_clientes(pagina, termo_busca(pessoa))
                        confirmacao = enviar_email(pagina, pessoa)
                        print(f"    enviado — {confirmacao}")

                        extras_msg = [confirmacao]
                        if not getattr(args, "sem_planilha", False):
                            info_plan = atualizar_status_planilha(contexto, pessoa, "Aguardando Aprovação")
                            if info_plan:
                                extras_msg.append(info_plan)
                            try:
                                pagina.bring_to_front()
                            except Exception:
                                pass

                        if not getattr(args, "sem_sge", False):
                            info_sge = dar_baixa_sge_enviado(p, pessoa, estado_sge)
                            if info_sge:
                                extras_msg.append(info_sge)
                            try:
                                pagina.bring_to_front()
                            except Exception:
                                pass

                        feitos[pessoa["sge"]] = {
                            "nome": pessoa["nome"],
                            "tamanhos": pessoa["tamanhos"],
                            "status_coluna_k": "Aguardando Aprovação",
                            "etapa_sge": "Enviado para Aprovação",
                            "data": f"{datetime.now():%Y-%m-%d %H:%M:%S}",
                        }
                        salvar_json(FEITOS, feitos)
                        enviados_rodada += 1
                        linhas.append(linha_relatorio(pessoa, achados, "enviado", " | ".join(extras_msg)))
                except PararPessoa as parada:
                    imagem = tirar_print(pagina, f"erro_{pessoa['sge']}")
                    detalhe = parada.detalhe + (f" | {imagem}" if imagem else "")
                    if parada.resultado == "duplicado":
                        print(f"DUPLICADO: {pessoa['nome']} (SGE {pessoa['sge']}) — {parada.detalhe}")
                    elif parada.resultado == "não encontrado" or "não encontr" in parada.detalhe.lower():
                        print(f"NÃO ENCONTREI: {pessoa['nome']} (SGE {pessoa['sge']}) — {parada.detalhe}")
                    print(f"    {parada.resultado} — {detalhe}")
                    linhas.append(linha_relatorio(pessoa, achados, parada.resultado, detalhe))
                except Exception as erro:
                    if "has been closed" in str(erro):
                        print("    o navegador foi fechado — parando")
                        linhas.append(linha_relatorio(pessoa, achados, "erro no Drive", "navegador fechado"))
                        break
                    imagem = tirar_print(pagina, f"erro_{pessoa['sge']}")
                    detalhe = f"{base.erro_curto(erro)}" + (f" | {imagem}" if imagem else "")
                    print(f"    erro — {detalhe}")
                    linhas.append(linha_relatorio(pessoa, achados, "erro no Drive", detalhe))
                salvar_relatorio(linhas)
        finally:
            if estado_sge.get("contexto") is not None:
                try:
                    estado_sge["contexto"].close()
                except Exception:
                    pass
            try:
                contexto.close()
            except Exception:
                pass

    duplicadas = [pessoa for pessoa, _ in prontos if pessoa.get("duplicados")]
    if duplicadas:
        print("\n" + "-" * 70)
        print(f"PESSOAS DUPLICADAS NO MAISBOOK ({len(duplicadas)}):")
        for pessoa in duplicadas:
            print(
                f"DUPLICADO: {pessoa['nome']} (SGE {pessoa['sge']}) — pesquisa {termo_busca(pessoa)} "
                f"trouxe {len(pessoa['duplicados'])} cartões"
            )
            for cartao in pessoa["duplicados"]:
                print(f"    {cartao}")
        print("-" * 70)

    print("\n" + "=" * 70)
    print(f"RELATORIO: {RELATORIO}")
    for resultado in ("enviado", "simulado", "faltou arquivo", "erro no Drive", "erro no e-mail", "não encontrado", "duplicado"):
        qtd = sum(1 for linha in linhas if linha["RESULTADO"] == resultado)
        if qtd:
            print(f"{resultado}: {qtd}")
    print("=" * 70)
    return enviados_rodada


def executar_conferencia_servidor_cli(args):
    """Executa apenas a conferência física de CAPA e ÁLBUM no servidor (somente para a etapa solicitada)."""
    albuns, indice = preparar_indice(getattr(args, "origem", None))
    etapa_alvo = getattr(args, "etapa", "DIAGRAMADO") or "DIAGRAMADO"
    pessoas = ler_diagramados(etapa_filtro=etapa_alvo)
    if args.sge:
        codigo = limpar_sge(args.sge)
        pessoas = [p for p in pessoas if p["sge"] == codigo]
    if args.limite:
        pessoas = pessoas[: args.limite]

    print("=" * 78)
    print(f"CONFERÊNCIA FÍSICA NO SERVIDOR: {albuns}")
    print(f"Etapa conferida: {etapa_alvo} | Total de formandos verificados: {len(pessoas)}")
    print("=" * 78)

    completos = 0
    so_album = 0
    so_capa = 0
    sem_nada = 0
    resultados = {}

    for p in pessoas:
        info = conferir_pessoa_no_servidor(p, albuns, indice)
        resultados[p["sge"]] = info
        col_m = p.get("capa_planilha") or "(vazia)"
        print(
            f"Linha {p['linha']} | SGE {p['sge']} — {p['nome']} | Col.M='{col_m}'\n"
            f"   -> ÁLBUM: {info['resumoAlbum']}\n"
            f"   -> CAPA : {info['resumoCapa']}\n"
            f"   -> GERAL: {info['resumoGeral']}"
        )
        if info["completo"]:
            completos += 1
        elif info["temAlbum"] and not info["temCapa"]:
            so_album += 1
        elif not info["temAlbum"] and info["temCapa"]:
            so_capa += 1
        else:
            sem_nada += 1

    print("=" * 78)
    print(
        f"RESUMO DA CONFERÊNCIA NO SERVIDOR:\n"
        f"  ✔ Completos (Álbum + Capa): {completos}\n"
        f"  ⚠ Só Álbum (Falta Capa)   : {so_album}\n"
        f"  ⚠ Só Capa (Falta Álbum)   : {so_capa}\n"
        f"  ✖ Sem Álbum e Sem Capa    : {sem_nada}"
    )
    print("=" * 78)
    salvar_json(PASTA / "conferencia_servidor.json", {
        "servidor": str(albuns),
        "verificadoEm": f"{datetime.now():%Y-%m-%d %H:%M:%S}",
        "resultados": resultados,
    })
    return resultados


def main():
    args = analisar_argumentos()
    if getattr(args, "gravar_celulas", None):
        dados = json.loads(args.gravar_celulas)
        gravar_edicoes_planilha(
            aba=dados.get("aba") or "GERAL ENVIOS",
            linha_num=int(dados.get("linha") or 0),
            celulas_dict=dados.get("celulas") or {},
            sge=dados.get("sge") or "",
            webhook_url=dados.get("webhook_url") or "",
            usar_navegador=bool(dados.get("usar_navegador", False)),
        )
        return
    if getattr(args, "conferir_json", None):
        entrada = args.conferir_json
        if Path(entrada).is_file():
            payload = json.loads(Path(entrada).read_text(encoding="utf-8"))
        else:
            payload = json.loads(entrada)
        origem_srv = payload.get("origem") or getattr(args, "origem", None)
        albuns, indice = preparar_indice(origem_srv)
        alunos_in = payload.get("alunos") or []
        saida = {}
        lista_resultados = []
        for idx_a, a in enumerate(alunos_in):
            tamanhos = a.get("tamanhos") or tamanhos_do_plano(a.get("tamanho") or "")
            id_unico = str(a.get("id") or "")
            linha_num = a.get("linha") or a.get("linhaPlanilha") or (idx_a + 1)
            pessoa = {
                "id": id_unico,
                "linha": linha_num,
                "nome": a.get("nome") or "",
                "sge": str(a.get("sge") or ""),
                "contrato": so_contrato(a.get("contrato") or a.get("cont") or ""),
                "tamanhos": tamanhos,
                "planos": [a.get("tamanho")] if a.get("tamanho") else [],
            }
            info_srv = conferir_pessoa_no_servidor(pessoa, albuns, indice)
            lista_resultados.append(info_srv)
            if id_unico:
                saida[id_unico] = info_srv
            saida[f"linha_{linha_num}"] = info_srv
            chave_composta = f"{pessoa['sge']}|{normalizar(pessoa['nome'])}"
            saida[chave_composta] = info_srv
            if pessoa["sge"] and pessoa["sge"] not in saida:
                saida[pessoa["sge"]] = info_srv
        resposta_json = {
            "ok": True,
            "servidor": str(albuns),
            "verificadoEm": f"{datetime.now():%Y-%m-%d %H:%M:%S}",
            "resultados": saida,
            "lista": lista_resultados,
        }
        print("JSON_CONFERENCIA_START:" + json.dumps(resposta_json, ensure_ascii=False))
        return
    if getattr(args, "diagnostico_json", False):
        diag = diagnosticar_elegibilidade_lote(
            origem_srv=getattr(args, "origem", None),
            sge_filtro=getattr(args, "sge", "") or "",
            ignorar_col_m=getattr(args, "ignorar_coluna_m", False),
        )
        print("JSON_DIAGNOSTICO_START:" + json.dumps(diag, ensure_ascii=False))
        return
    if getattr(args, "diagnostico", False):
        diag = diagnosticar_elegibilidade_lote(
            origem_srv=getattr(args, "origem", None),
            sge_filtro=getattr(args, "sge", "") or "",
            ignorar_col_m=getattr(args, "ignorar_coluna_m", False),
        )
        imprimir_diagnostico_elegibilidade(diag)
        return
    if getattr(args, "conferir_servidor", False):
        executar_conferencia_servidor_cli(args)
        return
    if args.listar or not args.vigia:
        executar_rodada(args, em_vigia=False)
        return

    intervalo = max(10, int(args.intervalo or 60))
    PARAR.unlink(missing_ok=True)
    print("=" * 70)
    print(f"VIGIA CONTÍNUO MAISBOOK INICIADO (intervalo: {intervalo}s)")
    print("Monitorando STATUS = DIAGRAMADO -> Drive + MaisBook + Coluna K (Aguardando Aprovação) + SGE (Enviado para Aprovação)")
    print("=" * 70)
    ciclo = 0
    try:
        while not pediu_parar():
            ciclo += 1
            print(f"\n[CICLO #{ciclo} — {datetime.now():%d/%m/%Y %H:%M:%S}] Verificando planilha e servidor...")
            try:
                executar_rodada(args, em_vigia=True)
            except Exception as erro:
                print(f"[AVISO VIGIA] Erro no ciclo #{ciclo}: {erro}")
            if pediu_parar():
                break
            print(f"[VIGIA] Aguardando {intervalo}s para a próxima verificação (use Parar para encerrar)...")
            for _ in range(intervalo):
                if pediu_parar():
                    break
                time.sleep(1)
    except KeyboardInterrupt:
        print("\n[VIGIA] Interrompido pelo usuário.")
    finally:
        PARAR.unlink(missing_ok=True)
        print("\n[VIGIA] Monitoramento encerrado.")


if __name__ == "__main__":
    try:
        main()
    except (FileNotFoundError, ValueError) as erro:
        print(f"\nERRO: {erro}")
        sys.exit(1)

