# -*- coding: utf-8 -*-
"""
Janela do Organizador de Álbuns — Painel Unificado (Dashboard + Monitor + Envio MaisBook + Baixa SGE).

Guias do programa:
  1. 📊 Monitor & Controle Geral (Dashboard):
     - Gráfico de Rosca interativo (distribuição percentual das 9 etapas da Coluna K)
     - Cartões KPI em tempo real (Total, Diagramados, Prontos c/ CAPA, Retidos s/ CAPA, Aguardando, Aprovados)
     - Controle completo de todos os alunos da aba GERAL ENVIOS
     - Reescrita em tempo real da planilha (Todas as colunas: Nome, Contrato, SGE, Tamanho, Curso, CPF, MaisBook, Status Col K, Lote, Capa Col M)
     - Disparo individual por aluno e disparo em lote tanto de Envio (MaisBook) quanto de Baixa (SGE)
     - Conferência física anti-erro de pasta e capa no servidor antes de enviar
  2. 📁 Buscar e copiar — planilha, pasta e plano; Simular só mostra, Copiar traz capa e lâminas.
  3. 🏷️ Renomear — troca um lote já existente pelo código SGE.
  4. 📤 Enviar aprovação (Avançado) — configurações detalhadas do robô do MaisBook.
"""

import csv
import ctypes
import json
import math
import os
import re
import string
import subprocess
import sys
import threading
import time
import unicodedata
import urllib.request
from datetime import datetime
from importlib.machinery import SourceFileLoader
from pathlib import Path

import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk

import openpyxl
from openpyxl.styles import Font
from openpyxl.utils import column_index_from_string, get_column_letter


PASTA = Path(__file__).resolve().parent
SCRIPT = next(PASTA.glob("PROCUARA*.PY"))
SCRIPT_ENVIO = PASTA / "ENVIO_APROVACAO_MAISBOOK.PY"
SCRIPT_BAIXA = PASTA.parent / "baixa_sge" / "BAIXA.PY"
CONFIG = PASTA / "organizador_config.json"
EDICOES_LOCAIS = PASTA / "edicoes_planilha_local.json"
RELATORIO = PASTA / "relatorio_organizador.csv"
RELATORIO_ENVIO = PASTA / "relatorio_envio_aprovacao.csv"
PARAR_ENVIO = PASTA / "envio_aprovacao_parar.flag"
PARAR_BAIXA = PASTA.parent / "baixa_sge" / "baixa_parar.flag"
PASTA_ERROS = PASTA / "erros"
CONSOLE_PIRATA = PASTA / "console_pirata.py"
LOG_LOTE = PASTA / "organizador_saida.log"
ALTURA_CMD = 660
RESULTADOS_ENVIO_OK = {"enviado", "simulado"}

LINK_PLANILHA_PADRAO = "https://docs.google.com/spreadsheets/d/1FeYQop9z6rVAFQuIs1EOiuQcE84t0iBgamYtMWUIAU0/edit?usp=sharing"
ABA_PADRAO = "GERAL ENVIOS"

RE_RESUMO = re.compile(r"RESUMO ORGANIZADOR — (.+)")
RE_RESUMO_RENOMEIO = re.compile(
    r"RESUMO RENOMEIO — (\w+): (\d+) \| JA OK: (\d+) \| NAO ENCONTRADOS: (\d+) \| ERROS: (\d+)"
)
RE_ABA = re.compile(r"^ABA:\s*(.+)\s*$")
RE_LINK = re.compile(r"/spreadsheets/d/[\w-]+")
ORIGEM_PADRAO = r"D:\ÁLBUNS COLAÇÃO 2025.2"
PLANILHA_BAIXADA = PASTA / "planilha_organizador.xlsx"
PLANILHA_ENVIO_XLSX = PASTA / "planilha_envio_aprovacao.xlsx"
TAMANHOS = ["20X20", "24X34", "30X20", "30X22"]
CATEGORIAS = ["CAPAS", "ESTOJO", "LAMINAS", "POSTERS"]
LETRAS = list(string.ascii_uppercase)

ETAPAS_COLUNA_K = [
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

CORES_ETAPAS_K = {
    "Sem álbum": "#94A3B8",             # Slate-400
    "AGUARDANDO FOTO": "#F59E0B",       # Amber-500
    "EM DIAGRAMAÇÃO": "#6366F1",        # Indigo-500
    "DIAGRAMADO": "#2563EB",            # Blue-600
    "Aguardando Aprovação": "#0284C7",  # Sky-600
    "Com Modificações": "#EA580C",      # Orange-600
    "Aprovado": "#059669",              # Emerald-600
    "IMPRESSÃO": "#7C3AED",             # Violet-600
    "Finalizado": "#0F766E",            # Teal-700
    "Outros / Vazio": "#CBD5E1",        # Slate-300
}

ETAPAS_BAIXA_SGE = [
    "(usar o da planilha)",
    "Em Diagramação",
    "Enviado para Aprovação",
    "Aprovado",
    "Em Impressão",
    "Pronto",
]

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

CODIGO_APPS_SCRIPT_WEBHOOK = """// Cole este código em Extensões > Apps Script na planilha do Google
// Depois clique em Implantar > Nova implantação > App da Web (Acesso: Qualquer pessoa)
function doPost(e) {
  try {
    var payload = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetName = payload.aba || payload.sheetName || "GERAL ENVIOS";
    var sheet = ss.getSheetByName(sheetName) || ss.getActiveSheet();
    var row = Number(payload.linha || payload.row || 0);
    var cells = payload.celulas || payload.cells || {};
    if (row > 0 && cells) {
      for (var colLetter in cells) {
        var rangeRef = colLetter + row;
        sheet.getRange(rangeRef).setValue(cells[colLetter]);
      }
    }
    return ContentService.createTextOutput(JSON.stringify({ ok: true, row: row }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}"""

# Tema Visual Moderno (Estilo Preview Monitor Contínuo — Sem Tema Hacker)
FUNDO = "#F8FAFC"          # Slate-50 (Fundo limpo da Preview)
CARTAO = "#FFFFFF"         # Branco puro (Superfície de cartões)
VERDE = "#0F172A"          # Slate-900 (Texto principal escuro e nítido)
VERDE_FORTE = "#2563EB"    # Blue-600 (Destaque / Ação Primária da Preview)
VERDE_FRACO = "#64748B"    # Slate-500 (Texto secundário / legendas)
BORDA = "#CBD5E1"          # Slate-300 (Bordas suaves de cartões)
BOTAO = "#FFFFFF"          # Branco puro (Superfície de botões)
BOTAO_ATIVO = "#EFF6FF"    # Blue-50 (Hover / Ativo suave)
SELECAO = "#DBEAFE"        # Blue-100 (Seleção de linhas e listas)
SUCESSO = "#059669"        # Emerald-600 (Pronto / Sucesso / Enviar)
AVISO = "#D97706"          # Amber-600 (Avisos / Retido sem Capa)
ERRO = "#DC2626"           # Red-600 (Erros / Parar)
ROXO = "#4F46E5"           # Indigo-600 (Baixa SGE / Diagramado)
FONTE = ("Segoe UI", 9)
FONTE_NEGRITO = ("Segoe UI", 9, "bold")
FONTE_TITULO = ("Segoe UI", 11, "bold")
FONTE_KPI = ("Segoe UI", 16, "bold")


def normalizar_txt(texto):
    texto = unicodedata.normalize("NFKD", str(texto or ""))
    texto = "".join(c for c in texto if not unicodedata.combining(c))
    return re.sub(r"\s+", " ", texto).strip().lower()


def capa_liberada_m(valor):
    t = normalizar_txt(valor)
    return bool(t) and t not in BLOQUEIOS_COLUNA_M


def classificar_etapa_k(valor):
    t = normalizar_txt(valor)
    if not t:
        return "Outros / Vazio"
    for etapa in ETAPAS_COLUNA_K:
        if normalizar_txt(etapa) == t:
            return etapa
    if "diagramado" in t and "em " not in t:
        return "DIAGRAMADO"
    if "diagramacao" in t:
        return "EM DIAGRAMAÇÃO"
    if "aguardando aprov" in t or "enviado" in t:
        return "Aguardando Aprovação"
    if "modificac" in t or "alterac" in t:
        return "Com Modificações"
    if "aprovado" in t:
        return "Aprovado"
    if "impressao" in t:
        return "IMPRESSÃO"
    if "finalizado" in t or "pronto" in t or "entregue" in t:
        return "Finalizado"
    if "foto" in t:
        return "AGUARDANDO FOTO"
    if "sem album" in t:
        return "Sem álbum"
    return "Outros / Vazio"


def limpar_sge(valor):
    texto = str(valor or "").strip()
    if texto.upper().startswith("633."):
        texto = texto[4:].strip()
    texto = re.sub(r"\.\d+$", "", texto)
    return texto.strip()


def aplicar_tema_terminal(root):
    """Tema Moderno da Preview (Slate Claro, Azul Royal e Esmeralda — Fim do Tema Hacker)."""

    root.configure(background=FUNDO)
    root.option_add("*TCombobox*Listbox.background", "#FFFFFF")
    root.option_add("*TCombobox*Listbox.foreground", VERDE)
    root.option_add("*TCombobox*Listbox.selectBackground", SELECAO)
    root.option_add("*TCombobox*Listbox.selectForeground", VERDE_FORTE)
    root.option_add("*TCombobox*Listbox.font", "SegoeUI 9")

    estilo = ttk.Style(root)
    try:
        estilo.theme_use("clam")
    except tk.TclError:
        pass

    estilo.configure(
        ".",
        background=FUNDO,
        foreground=VERDE,
        font=FONTE,
        troughcolor="#E2E8F0",
        bordercolor=BORDA,
        lightcolor="#FFFFFF",
        darkcolor=BORDA,
    )
    estilo.configure("TFrame", background=FUNDO)
    estilo.configure("Card.TFrame", background=CARTAO)
    estilo.configure("TLabel", background=FUNDO, foreground=VERDE, font=FONTE)
    estilo.configure("Card.TLabel", background=CARTAO, foreground=VERDE, font=FONTE)
    estilo.configure("Titulo.TLabel", background=FUNDO, foreground=VERDE, font=FONTE_TITULO)
    estilo.configure("Fraco.TLabel", background=FUNDO, foreground=VERDE_FRACO, font=FONTE)
    estilo.configure(
        "TLabelframe",
        background=CARTAO,
        foreground=VERDE,
        bordercolor=BORDA,
        relief="solid",
    )
    estilo.configure("TLabelframe.Label", background=CARTAO, foreground=VERDE_FORTE, font=FONTE_NEGRITO)
    estilo.configure(
        "TButton",
        background=BOTAO,
        foreground=VERDE,
        bordercolor=BORDA,
        focusthickness=1,
        focuscolor=VERDE_FORTE,
        padding=(8, 4),
        font=FONTE_NEGRITO,
    )
    estilo.map(
        "TButton",
        background=[("pressed", SELECAO), ("active", BOTAO_ATIVO), ("disabled", FUNDO)],
        foreground=[("disabled", VERDE_FRACO), ("pressed", VERDE_FORTE), ("active", VERDE_FORTE)],
        bordercolor=[("active", VERDE_FORTE), ("focus", VERDE_FORTE)],
    )
    estilo.configure("Primario.TButton", foreground="#FFFFFF", background=VERDE_FORTE, bordercolor="#1D4ED8")
    estilo.map(
        "Primario.TButton",
        background=[("pressed", "#1E40AF"), ("active", "#1D4ED8"), ("disabled", "#94A3B8")],
        foreground=[("disabled", "#F1F5F9"), ("pressed", "#FFFFFF"), ("active", "#FFFFFF")],
    )
    estilo.configure("Sucesso.TButton", foreground="#FFFFFF", background=SUCESSO, bordercolor="#047857")
    estilo.map(
        "Sucesso.TButton",
        background=[("pressed", "#065F46"), ("active", "#047857"), ("disabled", "#94A3B8")],
        foreground=[("disabled", "#F1F5F9"), ("pressed", "#FFFFFF"), ("active", "#FFFFFF")],
    )
    estilo.configure("Roxo.TButton", foreground="#FFFFFF", background=ROXO, bordercolor="#4338CA")
    estilo.map(
        "Roxo.TButton",
        background=[("pressed", "#3730A3"), ("active", "#4338CA"), ("disabled", "#94A3B8")],
        foreground=[("disabled", "#F1F5F9"), ("pressed", "#FFFFFF"), ("active", "#FFFFFF")],
    )
    estilo.configure("Aviso.TButton", foreground="#92400E", background="#FEF3C7", bordercolor="#F59E0B")
    estilo.map(
        "Aviso.TButton",
        background=[("pressed", "#FDE68A"), ("active", "#FDE68A"), ("disabled", FUNDO)],
        foreground=[("disabled", VERDE_FRACO), ("active", "#78350F")],
    )
    estilo.configure("Perigo.TButton", foreground=ERRO, bordercolor="#FCA5A5", background="#FEF2F2")
    estilo.map(
        "Perigo.TButton",
        background=[("pressed", "#FEE2E2"), ("active", "#FEE2E2"), ("disabled", FUNDO)],
        foreground=[("disabled", "#FCA5A5"), ("active", "#B91C1C")],
        bordercolor=[("active", ERRO), ("focus", ERRO)],
    )
    estilo.configure(
        "TEntry",
        fieldbackground="#FFFFFF",
        foreground=VERDE,
        insertcolor=VERDE,
        bordercolor=BORDA,
        lightcolor="#FFFFFF",
        darkcolor=BORDA,
    )
    estilo.map(
        "TEntry",
        fieldbackground=[("disabled", FUNDO)],
        foreground=[("disabled", VERDE_FRACO)],
    )
    estilo.configure(
        "TCombobox",
        fieldbackground="#FFFFFF",
        background=BOTAO,
        foreground=VERDE,
        arrowcolor=VERDE,
        bordercolor=BORDA,
        insertcolor=VERDE,
        arrowsize=12,
    )
    estilo.map(
        "TCombobox",
        fieldbackground=[("readonly", "#FFFFFF"), ("disabled", FUNDO)],
        foreground=[("readonly", VERDE), ("disabled", VERDE_FRACO)],
        selectbackground=[("readonly", SELECAO)],
        selectforeground=[("readonly", VERDE_FORTE)],
        background=[("active", BOTAO_ATIVO)],
        arrowcolor=[("disabled", VERDE_FRACO)],
    )
    estilo.configure("TCheckbutton", background=CARTAO, foreground=VERDE, font=FONTE)
    estilo.map(
        "TCheckbutton",
        background=[("active", CARTAO)],
        foreground=[("active", VERDE_FORTE), ("disabled", VERDE_FRACO)],
    )
    estilo.configure("TNotebook", background=FUNDO, bordercolor=BORDA, tabmargins=(2, 4, 2, 0))
    estilo.configure(
        "TNotebook.Tab",
        background="#E2E8F0",
        foreground=VERDE_FRACO,
        padding=(14, 6),
        font=FONTE_NEGRITO,
        bordercolor=BORDA,
    )
    estilo.map(
        "TNotebook.Tab",
        background=[("selected", CARTAO), ("active", BOTAO_ATIVO)],
        foreground=[("selected", VERDE_FORTE), ("active", VERDE_FORTE)],
        bordercolor=[("selected", VERDE_FORTE)],
    )
    estilo.configure(
        "Treeview",
        background="#FFFFFF",
        fieldbackground="#FFFFFF",
        foreground=VERDE,
        bordercolor=BORDA,
        font=FONTE,
        rowheight=25,
    )
    estilo.configure(
        "Treeview.Heading",
        background="#F1F5F9",
        foreground="#334155",
        font=FONTE_NEGRITO,
        bordercolor=BORDA,
        relief="flat",
    )
    estilo.map(
        "Treeview",
        background=[("selected", SELECAO)],
        foreground=[("selected", "#1E3A8A")],
    )
    estilo.map("Treeview.Heading", background=[("active", BOTAO_ATIVO)])
    estilo.configure("TSeparator", background=BORDA)
    for nome in ("Vertical.TScrollbar", "Horizontal.TScrollbar"):
        estilo.configure(
            nome,
            background=BOTAO,
            troughcolor="#F1F5F9",
            arrowcolor=VERDE,
            bordercolor=BORDA,
            relief="flat",
        )
        estilo.map(nome, background=[("active", BOTAO_ATIVO)])


def barra_escura(janela):
    try:
        janela.update_idletasks()
        hwnd = ctypes.windll.user32.GetParent(janela.winfo_id())
        valor = ctypes.c_int(0)
        ctypes.windll.dwmapi.DwmSetWindowAttribute(
            hwnd, 20, ctypes.byref(valor), ctypes.sizeof(valor)
        )
    except Exception:
        return


def pintar_console(texto):
    texto.configure(
        background="#FFFFFF",
        foreground=VERDE,
        insertbackground=VERDE,
        selectbackground=SELECAO,
        selectforeground=VERDE_FORTE,
        relief="flat",
        highlightthickness=1,
        highlightbackground=BORDA,
        highlightcolor=VERDE_FORTE,
        bd=0,
    )
    moldura = getattr(texto, "frame", None)
    if moldura is not None:
        moldura.configure(background=FUNDO)
    barra = getattr(texto, "vbar", None)
    if barra is not None and moldura is not None:
        barra.pack_forget()
        nova = ttk.Scrollbar(moldura, command=texto.yview)
        nova.pack(side="right", fill="y")
        texto.configure(yscrollcommand=nova.set)
        texto.vbar = nova


def caminho_planilha_erros(operacao):
    PASTA_ERROS.mkdir(exist_ok=True)
    return PASTA_ERROS / f"erros_{operacao}_{datetime.now():%Y-%m-%d_%H-%M-%S}.xlsx"


def planilha_erros_de_csv(relatorio, destino, eh_erro):
    """Copia do relatório CSV só as linhas com erro. Devolve quantas linhas foram gravadas."""
    with relatorio.open(encoding="utf-8-sig", newline="") as arquivo:
        leitor = csv.DictReader(arquivo, delimiter=";")
        campos = leitor.fieldnames or []
        linhas = [linha for linha in leitor if eh_erro(linha)]
    if not linhas:
        return 0
    livro = openpyxl.Workbook()
    aba = livro.active
    aba.title = "Erros"
    aba.append(campos)
    for celula in aba[1]:
        celula.font = Font(bold=True)
    for linha in linhas:
        aba.append([linha.get(campo, "") for campo in campos])
    for indice, campo in enumerate(campos, start=1):
        maior = max([len(str(campo))] + [len(str(linha.get(campo) or "")) for linha in linhas])
        aba.column_dimensions[get_column_letter(indice)].width = min(max(maior + 2, 10), 80)
    aba.freeze_panes = "A2"
    livro.save(destino)
    return len(linhas)


def script_renomeio():
    for pasta in PASTA.parent.iterdir():
        if not pasta.is_dir():
            continue
        achados = list(pasta.glob("PASTA RENOMEIO.py"))
        if achados:
            return achados[0]
    return None


def aplicar_overlay_edicoes_no_xlsx(caminho_xlsx):
    """Garante que qualquer célula reescrita em tempo real apareça imediatamente no .xlsx local."""
    if not EDICOES_LOCAIS.is_file() or not Path(caminho_xlsx).is_file():
        return
    try:
        dados = json.loads(EDICOES_LOCAIS.read_text(encoding="utf-8"))
        registros = dados.get("linhas", {})
        if not registros:
            return
        agora = time.time()
        wb = openpyxl.load_workbook(caminho_xlsx)
        alterou = False
        for _, info in list(registros.items()):
            if agora - float(info.get("ts", 0)) > 7200:
                continue
            aba = info.get("aba") or ABA_PADRAO
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


def baixar_planilha_google(link, destino):
    achado = re.search(r"/spreadsheets/d/([\w-]+)", link or "")
    if not achado:
        raise ValueError("Link da planilha inválido")
    gid = re.search(r"[#&?]gid=(\d+)", link)
    url = "https://docs.google.com/spreadsheets/d/" + achado.group(1) + "/export?format=xlsx"
    if gid:
        url += "&gid=" + gid.group(1)
    destino = Path(destino)
    temporario = destino.with_suffix(".baixando")
    try:
        urllib.request.urlretrieve(url, temporario)
        if temporario.read_bytes()[:2] != b"PK":
            raise ValueError(
                "o Google não devolveu a planilha "
                "(o link está compartilhado como 'qualquer pessoa com o link'?)"
            )
        temporario.replace(destino)
        aplicar_overlay_edicoes_no_xlsx(destino)
    except Exception:
        temporario.unlink(missing_ok=True)
        if destino.is_file():
            aplicar_overlay_edicoes_no_xlsx(destino)
            return destino
        raise
    return destino


def python_console():
    """python.exe ao lado do pythonw.exe atual (a saída precisa ser capturada)."""
    exe = Path(sys.executable)
    if exe.name.lower() == "pythonw.exe":
        candidato = exe.with_name("python.exe")
        if candidato.is_file():
            return str(candidato)
    return str(exe)


def carregar_config():
    try:
        return json.loads(CONFIG.read_text(encoding="utf-8"))
    except Exception:
        return {}


class JanelaOrganizador:

    def __init__(self, root):
        self.root = root
        self.processo = None
        self.resumo = None
        self.modo = None
        self.inicio_operacao = 0.0
        self.arquivo_erros = None
        self.pediu_parar_envio = False
        self.avisos = []
        self._timer_abas = None
        self._geracao_abas = 0
        self._timer_autosync = None
        self._carregando_dashboard = False

        # Dados do Dashboard (Monitor & Controle Geral)
        self.alunos_dashboard = []
        self.alunos_por_linha = {}
        self.status_servidor_por_sge = {}
        self.filtro_etapa_donut = None
        self.fatias_donut = []

        root.title("Organizador de Álbuns — Monitor Contínuo & Controle Geral (MaisBook + SGE)")
        root.geometry("1380x900")
        root.minsize(1100, 740)
        aplicar_tema_terminal(root)

        cfg = carregar_config()
        self.link = tk.StringVar(value=cfg.get("link") or LINK_PLANILHA_PADRAO)
        self.planilha = tk.StringVar(value=cfg.get("planilha", ""))
        self.aba = tk.StringVar(value=cfg.get("aba") or ABA_PADRAO)
        self.origem = tk.StringVar(value=cfg.get("origem", ORIGEM_PADRAO))
        self.destino = tk.StringVar(value=cfg.get("destino", ""))
        self.col_plano = tk.StringVar(value=cfg.get("col_plano", "F"))
        self.col_nome = tk.StringVar(value=cfg.get("col_nome", "C"))
        self.col_codigo = tk.StringVar(value=cfg.get("col_codigo", "E"))
        self.plano = tk.StringVar(value=cfg.get("plano", ""))
        self.tamanhos = {
            t: tk.BooleanVar(value=t in cfg.get("tamanhos", TAMANHOS)) for t in TAMANHOS
        }
        self.categorias = {
            c: tk.BooleanVar(value=c in cfg.get("categorias", CATEGORIAS)) for c in CATEGORIAS
        }
        self.pastas_existentes = tk.StringVar(value="")
        self.envio_sge = tk.StringVar(value="")
        self.envio_limite = tk.StringVar(value="")
        self.envio_intervalo = tk.StringVar(value=str(cfg.get("envio_intervalo", 60)))
        self.envio_simular = tk.BooleanVar(value=False)
        self.envio_exigir_coluna_m = tk.BooleanVar(value=cfg.get("envio_exigir_coluna_m", True))
        self.envio_atualizar_planilha = tk.BooleanVar(value=cfg.get("envio_atualizar_planilha", True))
        self.envio_baixa_sge = tk.BooleanVar(value=cfg.get("envio_baixa_sge", True))
        self.webhook_apps_script = tk.StringVar(value=cfg.get("webhook_apps_script", ""))
        self.gravar_via_chrome = tk.BooleanVar(value=cfg.get("gravar_via_chrome", False))
        self.autosync_ativo = tk.BooleanVar(value=cfg.get("autosync_ativo", False))

        # Filtros e controles do Dashboard
        self.dash_busca = tk.StringVar(value="")
        self.dash_filtro_status = tk.StringVar(value="Todos os Status (Col. K)")
        self.dash_filtro_capa = tk.StringVar(value="Todos (Col. M)")
        self.dash_etapa_sge = tk.StringVar(value="(usar o da planilha)")
        self.dash_novo_status_k = tk.StringVar(value="DIAGRAMADO")
        self.dash_resumo_selecionado = tk.StringVar(
            value="Selecione um formando na tabela para editar colunas, alternar CAPA (Col. M), enviar MaisBook ou dar baixa no SGE."
        )

        self.status = tk.StringVar(value="Carregando painel de controle e planilha...")
        self.combos_aba = []

        self._montar()
        self.link.trace_add("write", self._agendar_abas)
        self.planilha.trace_add("write", self._agendar_abas)
        self.destino.trace_add("write", lambda *_: self._atualizar_pastas())
        self.aba.trace_add("write", lambda *_: self._atualizar_previa())
        self.col_nome.trace_add("write", lambda *_: self._atualizar_previa())
        self.col_codigo.trace_add("write", lambda *_: self._atualizar_previa())
        self.dash_busca.trace_add("write", lambda *_: self._filtrar_tabela_dashboard())

        self._atualizar_pastas()
        if self.link.get().strip() or self.planilha.get().strip():
            self.root.after(200, self._carregar_abas_automatico)
        self.root.after(350, lambda: self._sincronizar_dashboard(forcar_download=True))
        barra_escura(root)

    def _montar(self):
        principal = ttk.Frame(self.root, padding=8)
        principal.pack(fill="both", expand=True)
        principal.columnconfigure(0, weight=1)
        principal.rowconfigure(0, weight=3)
        principal.rowconfigure(2, weight=1)

        caderno = ttk.Notebook(principal)
        caderno.grid(row=0, column=0, sticky="nsew")
        self.caderno = caderno

        # 1ª Guia Principal: Monitor & Controle Geral (Dashboard)
        self._montar_dashboard(caderno)
        # Guias Originais preservadas:
        self._montar_busca(caderno)
        self._montar_renomeio(caderno)
        self._montar_envio(caderno)

        barra_status = ttk.Frame(principal)
        barra_status.grid(row=1, column=0, sticky="we", pady=(6, 0))
        barra_status.columnconfigure(0, weight=1)
        ttk.Label(barra_status, textvariable=self.status, font=FONTE_NEGRITO).grid(row=0, column=0, sticky="w")
        ttk.Button(barra_status, text="Limpar Console", command=self._limpar).grid(row=0, column=1, sticky="e")

        self.texto = scrolledtext.ScrolledText(principal, height=8, state="disabled", font=("Consolas", 9))
        pintar_console(self.texto)
        self.texto.tag_configure("aviso", foreground=AVISO)
        self.texto.tag_configure("erro", foreground=ERRO)
        self.texto.tag_configure("ok", foreground=SUCESSO)
        self.texto.grid(row=2, column=0, sticky="nsew", pady=(4, 0))

    # ==========================================================================
    # GUIA 1: MONITOR & CONTROLE GERAL (DASHBOARD UNIFICADO)
    # ==========================================================================
    def _montar_dashboard(self, caderno):
        aba = ttk.Frame(caderno, padding=8)
        caderno.add(aba, text="📊 Monitor & Controle Geral (Dashboard)")
        aba.columnconfigure(0, weight=1)
        aba.rowconfigure(3, weight=1)

        # --- 1. Cartões KPI Superiores ---
        grade_kpi = ttk.Frame(aba)
        grade_kpi.grid(row=0, column=0, sticky="we", pady=(0, 6))
        self.kpi_vars = {}
        definicoes_kpi = [
            ("total", "TOTAL NA PLANILHA", "Formandos na aba", "#2563EB"),
            ("diagramados", "DIAGRAMADO (COL. K)", "Total na etapa", "#4F46E5"),
            ("prontos", "PRONTOS P/ ENVIAR", "DIAGRAMADO + CAPA (Col. M)", "#059669"),
            ("retidos", "RETIDOS SEM CAPA", "Aguardando OK na Col. M", "#D97706"),
            ("aguardando", "AGUARDANDO APROV.", "Já enviados no MaisBook", "#0284C7"),
            ("concluidos", "APROVADO / IMPRESSÃO", "Aprovados e Finalizados", "#0D9488"),
        ]
        for idx, (chave, titulo, sub, cor) in enumerate(definicoes_kpi):
            grade_kpi.columnconfigure(idx, weight=1)
            card = tk.Frame(
                grade_kpi,
                bg=CARTAO,
                highlightthickness=1,
                highlightbackground=BORDA,
                bd=0,
            )
            card.grid(row=0, column=idx, sticky="nsew", padx=(0 if idx == 0 else 6, 0))
            faixa = tk.Frame(card, bg=cor, height=4)
            faixa.pack(fill="x", side="top")
            corpo = tk.Frame(card, bg=CARTAO, padx=10, pady=6)
            corpo.pack(fill="both", expand=True)
            tk.Label(corpo, text=titulo, bg=CARTAO, fg=VERDE_FRACO, font=("Segoe UI", 8, "bold")).pack(anchor="w")
            var_num = tk.StringVar(value="0")
            self.kpi_vars[chave] = var_num
            tk.Label(corpo, textvariable=var_num, bg=CARTAO, fg=cor, font=FONTE_KPI).pack(anchor="w")
            tk.Label(corpo, text=sub, bg=CARTAO, fg=VERDE_FRACO, font=("Segoe UI", 8)).pack(anchor="w")

        # --- 2. Painel Meio: Gráfico de Rosca (Coluna K) + Central de Disparo em Lote ---
        meio = ttk.Frame(aba)
        meio.grid(row=1, column=0, sticky="we", pady=(0, 6))
        meio.columnconfigure(0, weight=5)
        meio.columnconfigure(1, weight=7)

        # 2A. Gráfico de Rosca (Donut Chart) Interativo da Coluna K
        quadro_donut = ttk.LabelFrame(
            meio,
            text="Distribuição Percentual — Etapas da Coluna K (Clique numa fatia para filtrar)",
            padding=8,
        )
        quadro_donut.grid(row=0, column=0, sticky="nsew", padx=(0, 6))
        quadro_donut.columnconfigure(1, weight=1)

        self.canvas_donut = tk.Canvas(
            quadro_donut,
            width=195,
            height=185,
            bg=CARTAO,
            highlightthickness=0,
            cursor="hand2",
        )
        self.canvas_donut.grid(row=0, column=0, rowspan=2, sticky="ns", padx=(0, 8))
        self.canvas_donut.bind("<Button-1>", self._clique_canvas_donut)

        self.quadro_legenda_donut = tk.Frame(quadro_donut, bg=CARTAO)
        self.quadro_legenda_donut.grid(row=0, column=1, sticky="nsew")

        rodape_donut = tk.Frame(quadro_donut, bg=CARTAO)
        rodape_donut.grid(row=1, column=1, sticky="we", pady=(4, 0))
        self.lbl_filtro_donut = tk.Label(
            rodape_donut,
            text="Exibindo: Todas as etapas",
            bg=CARTAO,
            fg=VERDE_FORTE,
            font=("Segoe UI", 8, "bold"),
        )
        self.lbl_filtro_donut.pack(side="left")
        ttk.Button(
            rodape_donut,
            text="Mostrar Todos",
            command=self._limpar_filtro_donut,
        ).pack(side="right")

        # 2B. Central de Comando Unificada (Envio MaisBook + Baixa SGE + Planilha em Tempo Real)
        quadro_cmd = ttk.LabelFrame(
            meio,
            text="Central de Comando em Lote — Envio MaisBook, Baixa SGE & Reescrita em Tempo Real",
            padding=8,
        )
        quadro_cmd.grid(row=0, column=1, sticky="nsew")
        quadro_cmd.columnconfigure(0, weight=1)

        # Linha 1: Sincronização e Planilha Google em Tempo Real
        bloco_sync = tk.Frame(quadro_cmd, bg=CARTAO)
        bloco_sync.grid(row=0, column=0, sticky="we", pady=(0, 6))
        ttk.Button(
            bloco_sync,
            text="🔄 Sincronizar Planilha Agora",
            style="Primario.TButton",
            command=lambda: self._sincronizar_dashboard(forcar_download=True),
        ).pack(side="left")
        ttk.Checkbutton(
            bloco_sync,
            text="Auto-Sync (30s)",
            variable=self.autosync_ativo,
            command=self._alternar_autosync,
        ).pack(side="left", padx=(8, 0))
        ttk.Checkbutton(
            bloco_sync,
            text="Gravar também via Chrome no Google Sheets",
            variable=self.gravar_via_chrome,
            command=self._salvar_config,
        ).pack(side="left", padx=(8, 0))
        ttk.Button(
            bloco_sync,
            text="⚡ Configurar Webhook Apps Script",
            command=self._abrir_modal_webhook,
        ).pack(side="right")

        ttk.Separator(quadro_cmd, orient="horizontal").grid(row=1, column=0, sticky="we", pady=4)

        # Linha 2: Disparo em Lote — Envio de Aprovação (MaisBook)
        bloco_mb = tk.Frame(quadro_cmd, bg=CARTAO)
        bloco_mb.grid(row=2, column=0, sticky="we", pady=4)
        tk.Label(
            bloco_mb,
            text="ENVIO MAISBOOK (LOTE):",
            bg=CARTAO,
            fg=SUCESSO,
            font=FONTE_NEGRITO,
            width=22,
            anchor="w",
        ).pack(side="left")
        self.btn_dash_enviar_lote = ttk.Button(
            bloco_mb,
            text="🚀 Enviar Prontos (DIAGRAMADO + CAPA)",
            style="Sucesso.TButton",
            command=self._dash_enviar_prontos_lote,
        )
        self.btn_dash_enviar_lote.pack(side="left", padx=(0, 6))
        self.btn_dash_vigia_mb = ttk.Button(
            bloco_mb,
            text="♾ Vigia Contínuo MaisBook",
            command=self._envio_vigia,
        )
        self.btn_dash_vigia_mb.pack(side="left", padx=(0, 6))
        self.btn_dash_conferir_srv = ttk.Button(
            bloco_mb,
            text="🔍 Conferir Pastas/Capas no Servidor",
            command=self._conferir_todos_no_servidor,
        )
        self.btn_dash_conferir_srv.pack(side="left")

        ttk.Separator(quadro_cmd, orient="horizontal").grid(row=3, column=0, sticky="we", pady=4)

        # Linha 3: Disparo em Lote — Baixa no SGE (baixa_sge/BAIXA.PY)
        bloco_sge = tk.Frame(quadro_cmd, bg=CARTAO)
        bloco_sge.grid(row=4, column=0, sticky="we", pady=(4, 0))
        tk.Label(
            bloco_sge,
            text="BAIXA NO SGE (LOTE):",
            bg=CARTAO,
            fg=ROXO,
            font=FONTE_NEGRITO,
            width=22,
            anchor="w",
        ).pack(side="left")
        ttk.Combobox(
            bloco_sge,
            textvariable=self.dash_etapa_sge,
            values=ETAPAS_BAIXA_SGE,
            state="readonly",
            width=21,
        ).pack(side="left", padx=(0, 6))
        self.btn_dash_baixa_lote = ttk.Button(
            bloco_sge,
            text="⚡ Executar Baixa SGE (Lote)",
            style="Roxo.TButton",
            command=self._dash_executar_baixa_sge_lote,
        )
        self.btn_dash_baixa_lote.pack(side="left", padx=(0, 6))
        self.btn_dash_vigia_sge = ttk.Button(
            bloco_sge,
            text="♾ Vigia SGE",
            command=self._dash_vigia_sge,
        )
        self.btn_dash_vigia_sge.pack(side="left", padx=(0, 6))
        self.btn_dash_parar_tudo = ttk.Button(
            bloco_sge,
            text="■ Parar Robô",
            style="Perigo.TButton",
            command=self._parar_qualquer_robo,
            state="disabled",
        )
        self.btn_dash_parar_tudo.pack(side="right")

        # --- 3. Barra de Busca e Filtros da Tabela de Alunos ---
        barra_filtro = ttk.Frame(aba)
        barra_filtro.grid(row=2, column=0, sticky="we", pady=(0, 4))
        ttk.Label(barra_filtro, text="🔎 Buscar aluno:", font=FONTE_NEGRITO).pack(side="left")
        ttk.Entry(barra_filtro, textvariable=self.dash_busca, width=32).pack(side="left", padx=(4, 12))

        ttk.Label(barra_filtro, text="Etapa (Col. K):").pack(side="left")
        opcoes_status_filtro = ["Todos os Status (Col. K)"] + ETAPAS_COLUNA_K + ["Outros / Vazio"]
        combo_st = ttk.Combobox(
            barra_filtro,
            textvariable=self.dash_filtro_status,
            values=opcoes_status_filtro,
            state="readonly",
            width=22,
        )
        combo_st.pack(side="left", padx=(4, 12))
        combo_st.bind("<<ComboboxSelected>>", lambda _e: self._ao_mudar_combo_status())

        ttk.Label(barra_filtro, text="Filtro CAPA (Col. M):").pack(side="left")
        opcoes_capa_filtro = [
            "Todos (Col. M)",
            "✅ Prontos p/ Enviar (DIAGRAMADO + CAPA)",
            "⏳ Diagramados Retidos (Sem CAPA na Col. M)",
            "📗 Com CAPA liberada (Col. M preenchida)",
            "📕 Sem CAPA (Col. M vazia)",
        ]
        combo_cp = ttk.Combobox(
            barra_filtro,
            textvariable=self.dash_filtro_capa,
            values=opcoes_capa_filtro,
            state="readonly",
            width=36,
        )
        combo_cp.pack(side="left", padx=(4, 8))
        combo_cp.bind("<<ComboboxSelected>>", lambda _e: self._filtrar_tabela_dashboard())

        self.lbl_contagem_tabela = ttk.Label(barra_filtro, text="0 alunos listados", style="Fraco.TLabel")
        self.lbl_contagem_tabela.pack(side="right")

        # --- 4. Tabela de Controle de Todos os Alunos ---
        quadro_tabela = ttk.Frame(aba)
        quadro_tabela.grid(row=3, column=0, sticky="nsew")
        quadro_tabela.columnconfigure(0, weight=1)
        quadro_tabela.rowconfigure(0, weight=1)

        colunas_tab = (
            "linha",
            "sge",
            "nome",
            "contrato",
            "tamanho",
            "curso",
            "cpf",
            "maisbook",
            "status_k",
            "lote",
            "capa_m",
            "servidor",
        )
        self.tabela_alunos = ttk.Treeview(
            quadro_tabela,
            columns=colunas_tab,
            show="headings",
            selectmode="browse",
            height=13,
        )
        definicoes_col = [
            ("linha", "Linha", 52, "center", False),
            ("sge", "SGE (E)", 78, "center", False),
            ("nome", "Nome do Formando (Col. C)", 250, "w", True),
            ("contrato", "Contrato (D)", 85, "center", False),
            ("tamanho", "Tamanho (Col. F)", 140, "w", False),
            ("curso", "Curso (Col. G)", 135, "w", False),
            ("cpf", "CPF (Col. H)", 110, "center", False),
            ("maisbook", "MaisBook (I)", 90, "center", False),
            ("status_k", "Status (Col. K)", 150, "center", False),
            ("lote", "Lote (L)", 80, "center", False),
            ("capa_m", "CAPA (Col. M)", 115, "center", False),
            ("servidor", "Conferência Servidor", 175, "w", False),
        ]
        for id_col, titulo, larg, anc, esticar in definicoes_col:
            self.tabela_alunos.heading(
                id_col,
                text=titulo,
                command=lambda c=id_col: self._ordenar_tabela_dashboard(c),
            )
            self.tabela_alunos.column(id_col, width=larg, anchor=anc, stretch=esticar)

        # Tags visuais por status
        self.tabela_alunos.tag_configure("pronto_envio", background="#ECFDF5", foreground="#065F46")
        self.tabela_alunos.tag_configure("retido_capa", background="#FFFBEB", foreground="#92400E")
        self.tabela_alunos.tag_configure("aguardando_aprov", background="#F0F9FF", foreground="#075985")
        self.tabela_alunos.tag_configure("aprovado_fim", background="#F0FDF4", foreground="#166534")
        self.tabela_alunos.tag_configure("modificacoes", background="#FEF2F2", foreground="#991B1B")
        self.tabela_alunos.tag_configure("erro_servidor", background="#FEE2E2", foreground="#991B1B")
        self.tabela_alunos.tag_configure("padrao", background="#FFFFFF", foreground=VERDE)

        scroll_v = ttk.Scrollbar(quadro_tabela, orient="vertical", command=self.tabela_alunos.yview)
        scroll_h = ttk.Scrollbar(quadro_tabela, orient="horizontal", command=self.tabela_alunos.xview)
        self.tabela_alunos.configure(yscrollcommand=scroll_v.set, xscrollcommand=scroll_h.set)

        self.tabela_alunos.grid(row=0, column=0, sticky="nsew")
        scroll_v.grid(row=0, column=1, sticky="ns")
        scroll_h.grid(row=1, column=0, sticky="we")

        self.tabela_alunos.bind("<<TreeviewSelect>>", self._ao_selecionar_aluno_dashboard)
        self.tabela_alunos.bind("<Double-1>", lambda _e: self._abrir_modal_editar_aluno())

        # --- 5. Barra Inferior de Controle Individual do Aluno & Reescrita em Tempo Real ---
        painel_ind = ttk.LabelFrame(
            aba,
            text="Controle Individual do Aluno Selecionado — Reescrita em Tempo Real na Planilha + Envio MaisBook + Baixa SGE",
            padding=8,
        )
        painel_ind.grid(row=4, column=0, sticky="we", pady=(6, 0))
        painel_ind.columnconfigure(0, weight=1)

        tk.Label(
            painel_ind,
            textvariable=self.dash_resumo_selecionado,
            bg=CARTAO,
            fg=VERDE_FORTE,
            font=FONTE_NEGRITO,
            anchor="w",
        ).grid(row=0, column=0, sticky="we", pady=(0, 6))

        botoes_ind = tk.Frame(painel_ind, bg=CARTAO)
        botoes_ind.grid(row=1, column=0, sticky="we")

        ttk.Button(
            botoes_ind,
            text="✅ Alternar CAPA (Col. M: CAPA ↔ Vazio)",
            style="Aviso.TButton",
            command=self._dash_alternar_capa_selecionado,
        ).pack(side="left", padx=(0, 8))

        tk.Label(botoes_ind, text="Status (Col. K):", bg=CARTAO, fg=VERDE, font=FONTE).pack(side="left")
        ttk.Combobox(
            botoes_ind,
            textvariable=self.dash_novo_status_k,
            values=ETAPAS_COLUNA_K,
            state="readonly",
            width=19,
        ).pack(side="left", padx=(4, 4))
        ttk.Button(
            botoes_ind,
            text="📌 Gravar Status na Planilha",
            command=self._dash_aplicar_status_k_selecionado,
        ).pack(side="left", padx=(0, 10))

        ttk.Button(
            botoes_ind,
            text="✏️ Editar Todas as Colunas do Aluno",
            style="Primario.TButton",
            command=self._abrir_modal_editar_aluno,
        ).pack(side="left", padx=(0, 10))

        ttk.Separator(botoes_ind, orient="vertical").pack(side="left", fill="y", padx=6)

        self.btn_ind_enviar_mb = ttk.Button(
            botoes_ind,
            text="📤 Enviar MaisBook (Só Este Aluno)",
            style="Sucesso.TButton",
            command=self._dash_enviar_maisbook_individual,
        )
        self.btn_ind_enviar_mb.pack(side="left", padx=(4, 6))

        self.btn_ind_baixa_sge = ttk.Button(
            botoes_ind,
            text="📋 Dar Baixa SGE (Só Este Aluno)",
            style="Roxo.TButton",
            command=self._dash_baixa_sge_individual,
        )
        self.btn_ind_baixa_sge.pack(side="left", padx=(0, 6))

        ttk.Button(
            botoes_ind,
            text="📁 Conferir Pasta/Capa",
            command=self._dash_conferir_servidor_individual,
        ).pack(side="left")

    # ==========================================================================
    # LÓGICA DO DASHBOARD: LEITURA, GRÁFICO DE ROSCA, EDIÇÃO EM TEMPO REAL E AÇÕES
    # ==========================================================================
    def _alternar_autosync(self):
        self._salvar_config()
        if self._timer_autosync:
            self.root.after_cancel(self._timer_autosync)
            self._timer_autosync = None
        if self.autosync_ativo.get():
            self._agendar_proximo_autosync()

    def _agendar_proximo_autosync(self):
        if not self.autosync_ativo.get():
            return
        self._timer_autosync = self.root.after(30000, self._executar_autosync)

    def _executar_autosync(self):
        self._timer_autosync = None
        if not self.processo and not self._carregando_dashboard:
            self._sincronizar_dashboard(forcar_download=True, silencioso=True)
        self._agendar_proximo_autosync()

    def _sincronizar_dashboard(self, forcar_download=True, silencioso=False):
        if self._carregando_dashboard:
            return
        self._carregando_dashboard = True
        if not silencioso:
            self.status.set("Sincronizando alunos da planilha (GERAL ENVIOS)...")
        threading.Thread(
            target=self._thread_ler_planilha_dashboard,
            args=(forcar_download, silencioso),
            daemon=True,
        ).start()

    def _thread_ler_planilha_dashboard(self, forcar_download, silencioso):
        erro = ""
        alunos = []
        nome_aba_lida = self.aba.get().strip() or ABA_PADRAO
        try:
            link = self.link.get().strip() or LINK_PLANILHA_PADRAO
            arq_local = self.planilha.get().strip()
            if arq_local and Path(arq_local).is_file() and not link:
                caminho = Path(arq_local)
                aplicar_overlay_edicoes_no_xlsx(caminho)
            else:
                if forcar_download or not PLANILHA_BAIXADA.is_file():
                    caminho = baixar_planilha_google(link, PLANILHA_BAIXADA)
                else:
                    caminho = PLANILHA_BAIXADA
                    aplicar_overlay_edicoes_no_xlsx(caminho)

            wb = openpyxl.load_workbook(caminho, read_only=True, data_only=True)
            if nome_aba_lida not in wb.sheetnames:
                nome_aba_lida = "GERAL ENVIOS" if "GERAL ENVIOS" in wb.sheetnames else wb.sheetnames[0]
            ws = wb[nome_aba_lida]

            linhas_brutas = list(ws.iter_rows(values_only=True))
            wb.close()

            # Localiza a linha de cabeçalho (procura NOME, SGE, TAMANHO, STATUS)
            linha_cab = 1
            col_map = {
                "nome": 3,      # C
                "contrato": 4,  # D
                "sge": 5,       # E
                "tamanho": 6,   # F
                "curso": 7,     # G
                "cpf": 8,       # H
                "maisbook": 9,  # I
                "status": 11,   # K
                "lote": 12,     # L
                "capa": 13,     # M
            }
            for idx_l, valores in enumerate(linhas_brutas[:35], start=1):
                textos = [normalizar_txt(v) for v in valores]
                if "sge" in textos and "status" in textos:
                    linha_cab = idx_l
                    for i_c, t_c in enumerate(textos, start=1):
                        if t_c in ("nome", "nome completo", "cliente", "aluno", "formando"):
                            col_map["nome"] = i_c
                        elif t_c in ("cont", "contrato", "ctt"):
                            col_map["contrato"] = i_c
                        elif t_c == "sge":
                            col_map["sge"] = i_c
                        elif t_c == "tamanho":
                            col_map["tamanho"] = i_c
                        elif t_c == "curso":
                            col_map["curso"] = i_c
                        elif t_c == "cpf":
                            col_map["cpf"] = i_c
                        elif t_c == "maisbook":
                            col_map["maisbook"] = i_c
                        elif t_c == "status":
                            col_map["status"] = i_c
                        elif t_c == "lote":
                            col_map["lote"] = i_c
                        elif t_c == "capa" and i_c > col_map.get("status", 10):
                            col_map["capa"] = i_c
                    break

            def pegar_cel(row_vals, idx_1based):
                if idx_1based < 1 or idx_1based > len(row_vals):
                    return ""
                v = row_vals[idx_1based - 1]
                if v is None:
                    return ""
                if isinstance(v, float) and v.is_integer():
                    return str(int(v))
                return str(v).strip()

            for num_linha in range(linha_cab + 1, len(linhas_brutas) + 1):
                vals = linhas_brutas[num_linha - 1]
                nome = pegar_cel(vals, col_map["nome"])
                sge_bruto = pegar_cel(vals, col_map["sge"])
                sge = limpar_sge(sge_bruto)
                if not nome and not sge:
                    continue
                status_k = pegar_cel(vals, col_map["status"])
                capa_m = pegar_cel(vals, col_map["capa"])
                etapa_k = classificar_etapa_k(status_k)
                liberado_m = capa_liberada_m(capa_m)
                alunos.append({
                    "linha": num_linha,
                    "aba": nome_aba_lida,
                    "sge": sge or sge_bruto,
                    "nome": nome,
                    "contrato": pegar_cel(vals, col_map["contrato"]),
                    "tamanho": pegar_cel(vals, col_map["tamanho"]),
                    "curso": pegar_cel(vals, col_map["curso"]),
                    "cpf": pegar_cel(vals, col_map["cpf"]),
                    "maisbook": pegar_cel(vals, col_map["maisbook"]),
                    "status_k": status_k,
                    "etapa_k": etapa_k,
                    "lote": pegar_cel(vals, col_map["lote"]),
                    "capa_m": capa_m,
                    "liberado_m": liberado_m,
                })
        except Exception as e:
            erro = str(e)

        self.root.after(0, self._aplicar_dados_dashboard, alunos, nome_aba_lida, erro, silencioso)

    def _aplicar_dados_dashboard(self, alunos, nome_aba, erro, silencioso):
        self._carregando_dashboard = False
        if erro and not alunos:
            if not silencioso:
                self.status.set(f"Erro ao sincronizar planilha: {erro}")
            return

        self.alunos_dashboard = alunos
        self.alunos_por_linha = {a["linha"]: a for a in alunos}

        # Atualiza KPIs
        total = len(alunos)
        diagramados = [a for a in alunos if a["etapa_k"] == "DIAGRAMADO"]
        prontos = [a for a in diagramados if a["liberado_m"]]
        retidos = [a for a in diagramados if not a["liberado_m"]]
        aguardando = [a for a in alunos if a["etapa_k"] == "Aguardando Aprovação"]
        concluidos = [a for a in alunos if a["etapa_k"] in ("Aprovado", "IMPRESSÃO", "Finalizado")]

        self.kpi_vars["total"].set(str(total))
        self.kpi_vars["diagramados"].set(str(len(diagramados)))
        self.kpi_vars["prontos"].set(str(len(prontos)))
        self.kpi_vars["retidos"].set(str(len(retidos)))
        self.kpi_vars["aguardando"].set(str(len(aguardando)))
        self.kpi_vars["concluidos"].set(str(len(concluidos)))

        self._desenhar_grafico_rosca()
        self._filtrar_tabela_dashboard()

        if not silencioso:
            hora = datetime.now().strftime("%H:%M:%S")
            self.status.set(
                f"Sincronizado às {hora} ({nome_aba}) — {total} alunos | "
                f"{len(diagramados)} DIAGRAMADO ({len(prontos)} prontos c/ CAPA, {len(retidos)} aguardando CAPA na Col. M)"
            )

    def _desenhar_grafico_rosca(self):
        """Desenha o Gráfico de Rosca (Donut Chart) da distribuição percentual da Coluna K."""
        c = self.canvas_donut
        c.delete("all")
        for w in self.quadro_legenda_donut.winfo_children():
            w.destroy()

        alunos = self.alunos_dashboard
        total = len(alunos)
        contagem = {etapa: 0 for etapa in ETAPAS_COLUNA_K}
        contagem["Outros / Vazio"] = 0
        for a in alunos:
            et = a["etapa_k"]
            contagem[et] = contagem.get(et, 0) + 1

        cx, cy = 96, 92
        r_ext = 78
        r_int = 46
        caixa = (cx - r_ext, cy - r_ext, cx + r_ext, cy + r_ext)

        self.fatias_donut = []
        if total == 0:
            c.create_oval(caixa, fill="#E2E8F0", outline=CARTAO, width=2)
        else:
            angulo_inicio = 90.0
            etapas_com_dados = [
                (et, contagem[et])
                for et in (ETAPAS_COLUNA_K + ["Outros / Vazio"])
                if contagem.get(et, 0) > 0
            ]
            for etapa, qtd in etapas_com_dados:
                ext = -(qtd / total) * 360.0
                if abs(ext) >= 359.9:
                    ext = -359.9
                cor = CORES_ETAPAS_K.get(etapa, "#94A3B8")
                selecionada = self.filtro_etapa_donut == etapa
                largura_borda = 3 if selecionada else 1.5
                cor_borda = "#0F172A" if selecionada else CARTAO
                c.create_arc(
                    caixa,
                    start=angulo_inicio,
                    extent=ext,
                    fill=cor,
                    outline=cor_borda,
                    width=largura_borda,
                    style="pieslice",
                )
                self.fatias_donut.append((angulo_inicio, angulo_inicio + ext, etapa, qtd))
                angulo_inicio += ext

        # Furo central do Donut (branco puro)
        c.create_oval(
            cx - r_int,
            cy - r_int,
            cx + r_int,
            cy + r_int,
            fill=CARTAO,
            outline="#E2E8F0",
            width=1,
        )

        if self.filtro_etapa_donut and total > 0:
            qtd_f = contagem.get(self.filtro_etapa_donut, 0)
            pct_f = (qtd_f / total) * 100.0
            c.create_text(
                cx,
                cy - 8,
                text=f"{pct_f:.1f}%",
                fill=CORES_ETAPAS_K.get(self.filtro_etapa_donut, VERDE),
                font=("Segoe UI", 13, "bold"),
            )
            c.create_text(
                cx,
                cy + 12,
                text=f"{qtd_f} de {total}",
                fill=VERDE_FRACO,
                font=("Segoe UI", 8, "bold"),
            )
        else:
            c.create_text(cx, cy - 8, text=str(total), fill=VERDE, font=("Segoe UI", 15, "bold"))
            c.create_text(cx, cy + 12, text="Formandos", fill=VERDE_FRACO, font=("Segoe UI", 8))

        # Legenda interativa ao lado do Gráfico de Rosca
        for idx, etapa in enumerate(ETAPAS_COLUNA_K):
            qtd = contagem.get(etapa, 0)
            pct = (qtd / total * 100.0) if total > 0 else 0.0
            cor = CORES_ETAPAS_K.get(etapa, "#94A3B8")
            ativa = self.filtro_etapa_donut == etapa
            fundo_item = SELECAO if ativa else CARTAO

            linha_leg = tk.Frame(self.quadro_legenda_donut, bg=fundo_item, cursor="hand2", padx=4, pady=1)
            linha_leg.grid(row=idx // 2, column=idx % 2, sticky="we", padx=2, pady=1)
            self.quadro_legenda_donut.columnconfigure(idx % 2, weight=1)

            bolinha = tk.Canvas(linha_leg, width=10, height=10, bg=fundo_item, highlightthickness=0)
            bolinha.pack(side="left", padx=(0, 4))
            bolinha.create_oval(1, 1, 9, 9, fill=cor, outline=cor)

            lbl_nome = tk.Label(
                linha_leg,
                text=f"{etapa}:",
                bg=fundo_item,
                fg=VERDE_FORTE if ativa else VERDE,
                font=("Segoe UI", 8, "bold" if ativa else "normal"),
                anchor="w",
            )
            lbl_nome.pack(side="left")

            lbl_val = tk.Label(
                linha_leg,
                text=f"{qtd} ({pct:.1f}%)",
                bg=fundo_item,
                fg=VERDE_FRACO,
                font=("Segoe UI", 8, "bold"),
                anchor="e",
            )
            lbl_val.pack(side="right", padx=(4, 0))

            for widget in (linha_leg, bolinha, lbl_nome, lbl_val):
                widget.bind("<Button-1>", lambda _e, et=etapa: self._selecionar_etapa_donut(et))

    def _clique_canvas_donut(self, evento):
        cx, cy = 96, 92
        dx = evento.x - cx
        dy = cy - evento.y  # Y invertido na tela
        dist = math.hypot(dx, dy)
        if dist < 44 or dist > 82:
            self._limpar_filtro_donut()
            return
        ang = (math.degrees(math.atan2(dy, dx)) + 360.0) % 360.0
        for ini, fim, etapa, _ in self.fatias_donut:
            # ini decresce até fim (sentido horário)
            a_ini = ini % 360.0
            ext = ini - fim
            diff = (a_ini - ang + 360.0) % 360.0
            if diff <= ext:
                self._selecionar_etapa_donut(etapa)
                return

    def _selecionar_etapa_donut(self, etapa):
        if self.filtro_etapa_donut == etapa:
            self._limpar_filtro_donut()
            return
        self.filtro_etapa_donut = etapa
        self.dash_filtro_status.set(etapa)
        self.lbl_filtro_donut.configure(text=f"Filtro ativo: {etapa}")
        self._desenhar_grafico_rosca()
        self._filtrar_tabela_dashboard()

    def _limpar_filtro_donut(self):
        self.filtro_etapa_donut = None
        self.dash_filtro_status.set("Todos os Status (Col. K)")
        self.lbl_filtro_donut.configure(text="Exibindo: Todas as etapas")
        self._desenhar_grafico_rosca()
        self._filtrar_tabela_dashboard()

    def _ao_mudar_combo_status(self):
        val = self.dash_filtro_status.get()
        if val == "Todos os Status (Col. K)":
            self.filtro_etapa_donut = None
            self.lbl_filtro_donut.configure(text="Exibindo: Todas as etapas")
        else:
            self.filtro_etapa_donut = val
            self.lbl_filtro_donut.configure(text=f"Filtro ativo: {val}")
        self._desenhar_grafico_rosca()
        self._filtrar_tabela_dashboard()

    def _filtrar_tabela_dashboard(self):
        if not hasattr(self, "tabela_alunos"):
            return
        busca = normalizar_txt(self.dash_busca.get())
        filtro_st = self.dash_filtro_status.get()
        filtro_cp = self.dash_filtro_capa.get()

        selecionados_antes = self.tabela_alunos.selection()
        iid_antes = selecionados_antes[0] if selecionados_antes else None

        self.tabela_alunos.delete(*self.tabela_alunos.get_children())
        mostrados = 0

        for a in self.alunos_dashboard:
            if filtro_st != "Todos os Status (Col. K)" and a["etapa_k"] != filtro_st:
                continue
            if filtro_cp == "✅ Prontos p/ Enviar (DIAGRAMADO + CAPA)":
                if not (a["etapa_k"] == "DIAGRAMADO" and a["liberado_m"]):
                    continue
            elif filtro_cp == "⏳ Diagramados Retidos (Sem CAPA na Col. M)":
                if not (a["etapa_k"] == "DIAGRAMADO" and not a["liberado_m"]):
                    continue
            elif filtro_cp == "📗 Com CAPA liberada (Col. M preenchida)":
                if not a["liberado_m"]:
                    continue
            elif filtro_cp == "📕 Sem CAPA (Col. M vazia)":
                if a["liberado_m"]:
                    continue

            if busca:
                alvo = normalizar_txt(
                    f"{a['sge']} {a['nome']} {a['contrato']} {a['tamanho']} "
                    f"{a['curso']} {a['cpf']} {a['maisbook']} {a['status_k']} {a['lote']} {a['capa_m']}"
                )
                if busca not in alvo:
                    continue

            srv_txt = self.status_servidor_por_sge.get(a["sge"], "—")
            capa_exib = f"✅ {a['capa_m']}" if a["liberado_m"] else (a["capa_m"] or "(vazia)")

            if "⚠" in srv_txt or "Falta" in srv_txt:
                tag = "erro_servidor"
            elif a["etapa_k"] == "DIAGRAMADO" and a["liberado_m"]:
                tag = "pronto_envio"
            elif a["etapa_k"] == "DIAGRAMADO" and not a["liberado_m"]:
                tag = "retido_capa"
            elif a["etapa_k"] == "Aguardando Aprovação":
                tag = "aguardando_aprov"
            elif a["etapa_k"] in ("Aprovado", "IMPRESSÃO", "Finalizado"):
                tag = "aprovado_fim"
            elif a["etapa_k"] == "Com Modificações":
                tag = "modificacoes"
            else:
                tag = "padrao"

            iid = str(a["linha"])
            self.tabela_alunos.insert(
                "",
                "end",
                iid=iid,
                values=(
                    a["linha"],
                    a["sge"],
                    a["nome"],
                    a["contrato"],
                    a["tamanho"],
                    a["curso"],
                    a["cpf"],
                    a["maisbook"],
                    a["status_k"] or "(vazio)",
                    a["lote"],
                    capa_exib,
                    srv_txt,
                ),
                tags=(tag,),
            )
            mostrados += 1

        self.lbl_contagem_tabela.configure(
            text=f"{mostrados} de {len(self.alunos_dashboard)} alunos exibidos"
        )
        if iid_antes and self.tabela_alunos.exists(iid_antes):
            self.tabela_alunos.selection_set(iid_antes)
            self._ao_selecionar_aluno_dashboard()

    def _ordenar_tabela_dashboard(self, col):
        reverso = getattr(self, "_ordem_reversa_col", None) == col
        self._ordem_reversa_col = None if reverso else col

        def chave_ord(a):
            if col == "linha":
                return int(a.get("linha") or 0)
            if col == "sge":
                s = re.sub(r"\D", "", str(a.get("sge") or ""))
                return int(s) if s else 0
            return normalizar_txt(a.get(col, ""))

        self.alunos_dashboard.sort(key=chave_ord, reverse=reverso)
        self._filtrar_tabela_dashboard()

    def _aluno_selecionado_dashboard(self):
        sel = self.tabela_alunos.selection()
        if not sel:
            return None
        linha = int(sel[0])
        return self.alunos_por_linha.get(linha)

    def _ao_selecionar_aluno_dashboard(self, *_):
        aluno = self._aluno_selecionado_dashboard()
        if not aluno:
            self.dash_resumo_selecionado.set(
                "Selecione um formando na tabela para editar colunas, alternar CAPA (Col. M), enviar MaisBook ou dar baixa no SGE."
            )
            return
        capa_txt = aluno["capa_m"] or "(vazia)"
        lib_txt = "LIBERADO P/ ENVIAR" if aluno["liberado_m"] else "RETIDO SEM CAPA"
        self.dash_resumo_selecionado.set(
            f"Selecionado: Linha {aluno['linha']} | SGE {aluno['sge']} — {aluno['nome']} | "
            f"Tamanho: {aluno['tamanho'] or '-'} | Status (Col. K): {aluno['status_k'] or '-'} | "
            f"Coluna M (CAPA): {capa_txt} [{lib_txt}]"
        )
        if aluno["etapa_k"] in ETAPAS_COLUNA_K:
            self.dash_novo_status_k.set(aluno["etapa_k"])
        self.envio_sge.set(aluno["sge"])

    # --------------------------------------------------------------------------
    # REESCRITA EM TEMPO REAL DA PLANILHA (TODAS AS COLUNAS DO FORMANDO)
    # --------------------------------------------------------------------------
    def _executar_reescrita_planilha(self, aluno, celulas_dict, descricao_acao):
        """Grava as alterações em tempo real (Overlay 0ms + .xlsx local + Webhook Apps Script + Chrome opcional)."""
        aba_alvo = aluno.get("aba") or self.aba.get().strip() or ABA_PADRAO
        linha_num = int(aluno["linha"])
        sge_aluno = aluno.get("sge", "")

        # Atualiza em memória imediatamente para resposta em tempo real na tela
        mapa_campo = {
            "C": "nome",
            "D": "contrato",
            "E": "sge",
            "F": "tamanho",
            "G": "curso",
            "H": "cpf",
            "I": "maisbook",
            "K": "status_k",
            "L": "lote",
            "M": "capa_m",
        }
        for letra, val in celulas_dict.items():
            campo = mapa_campo.get(letra.upper())
            if campo:
                aluno[campo] = val
                if campo == "status_k":
                    aluno["etapa_k"] = classificar_etapa_k(val)
                elif campo == "capa_m":
                    aluno["liberado_m"] = capa_liberada_m(val)

        # Recalcula KPIs, Gráfico de Rosca e Tabela na hora (0ms)
        self._aplicar_dados_dashboard(self.alunos_dashboard, aba_alvo, "", silencioso=True)
        self.status.set(f"Gravando na planilha em tempo real: {descricao_acao}...")
        self._escrever(f"\n>>> [TEMPO REAL] {descricao_acao} (Linha {linha_num}, SGE {sge_aluno})\n")

        usar_nav = bool(self.gravar_via_chrome.get())
        webhook = self.webhook_apps_script.get().strip()

        def _worker():
            try:
                mod_envio = SourceFileLoader("envio_mb_writer", str(SCRIPT_ENVIO)).load_module()
                res = mod_envio.gravar_edicoes_planilha(
                    aba=aba_alvo,
                    linha_num=linha_num,
                    celulas_dict=celulas_dict,
                    sge=sge_aluno,
                    webhook_url=webhook,
                    usar_navegador=usar_nav,
                )
                msg = res.get("mensagem", "Planilha atualizada.")
                self.root.after(0, lambda: self.status.set(f"✔ Planilha atualizada: {descricao_acao}"))
                self.root.after(0, lambda: self._escrever(f"    ✔ {msg}\n"))
            except Exception as e:
                self.root.after(0, lambda: self.status.set(f"Erro ao gravar na planilha: {e}"))
                self.root.after(0, lambda: self._escrever(f"    ⚠ Erro ao gravar na planilha: {e}\n"))

        threading.Thread(target=_worker, daemon=True).start()

    def _dash_alternar_capa_selecionado(self):
        aluno = self._aluno_selecionado_dashboard()
        if not aluno:
            messagebox.showinfo("Alternar CAPA (Coluna M)", "Selecione um formando na tabela primeiro.")
            return
        novo_valor = "" if aluno["liberado_m"] else "CAPA"
        acao_txt = (
            f"MARCAR 'CAPA' na Coluna M para liberar o envio de:\n\n{aluno['nome']} (SGE {aluno['sge']}, Linha {aluno['linha']})"
            if novo_valor
            else f"REMOVER o OK da Coluna M (deixar vazia/retida) para:\n\n{aluno['nome']} (SGE {aluno['sge']}, Linha {aluno['linha']})"
        )
        if not messagebox.askyesno("Confirmar Alteração na Coluna M (CAPA)", acao_txt + "\n\nConfirmar gravação em tempo real?"):
            return
        self._executar_reescrita_planilha(
            aluno,
            {"M": novo_valor},
            f"Coluna M (CAPA) -> {novo_valor or '(vazia)'} para {aluno['nome']}",
        )

    def _dash_aplicar_status_k_selecionado(self):
        aluno = self._aluno_selecionado_dashboard()
        if not aluno:
            messagebox.showinfo("Mudar Status (Coluna K)", "Selecione um formando na tabela primeiro.")
            return
        novo_status = self.dash_novo_status_k.get().strip()
        if not novo_status:
            return
        if not messagebox.askyesno(
            "Confirmar Alteração na Coluna K (STATUS)",
            f"Alterar a Coluna K (STATUS) de:\n\n"
            f"Formando: {aluno['nome']} (SGE {aluno['sge']}, Linha {aluno['linha']})\n"
            f"De: '{aluno['status_k'] or '(vazio)'}'\n"
            f"Para: '{novo_status}'\n\n"
            f"Confirmar reescrita em tempo real na planilha?",
        ):
            return
        self._executar_reescrita_planilha(
            aluno,
            {"K": novo_status},
            f"Coluna K (STATUS) -> '{novo_status}' para {aluno['nome']}",
        )

    def _abrir_modal_editar_aluno(self):
        aluno = self._aluno_selecionado_dashboard()
        if not aluno:
            messagebox.showinfo("Editar Formando", "Selecione um formando na tabela primeiro.")
            return

        janela = tk.Toplevel(self.root)
        janela.title(f"Editar Formando (Linha {aluno['linha']}) — Reescrita em Tempo Real na Planilha")
        janela.geometry("620x520")
        janela.configure(bg=CARTAO)
        janela.transient(self.root)
        janela.grab_set()

        topo = tk.Frame(janela, bg=VERDE_FORTE, padx=14, pady=10)
        topo.pack(fill="x")
        tk.Label(
            topo,
            text=f"✏️ Editar Todas as Colunas — Linha {aluno['linha']} ({aluno['aba']})",
            bg=VERDE_FORTE,
            fg="#FFFFFF",
            font=FONTE_TITULO,
        ).pack(anchor="w")
        tk.Label(
            topo,
            text="As alterações serão gravadas em tempo real na planilha e refletidas imediatamente no Monitor.",
            bg=VERDE_FORTE,
            fg="#DBEAFE",
            font=("Segoe UI", 8),
        ).pack(anchor="w")

        form = ttk.Frame(janela, padding=14, style="Card.TFrame")
        form.pack(fill="both", expand=True)
        form.columnconfigure(1, weight=1)

        campos_def = [
            ("C", "Nome do Formando (Col. C):", "nome", None),
            ("D", "Contrato (Col. D):", "contrato", None),
            ("E", "Código SGE (Col. E):", "sge", None),
            ("F", "Tamanho / Plano (Col. F):", "tamanho", None),
            ("G", "Curso (Col. G):", "curso", None),
            ("H", "CPF (Col. H):", "cpf", None),
            ("I", "Código MaisBook (Col. I):", "maisbook", None),
            ("K", "Status Etapa (Col. K):", "status_k", ETAPAS_COLUNA_K),
            ("L", "Lote (Col. L):", "lote", None),
            ("M", "CAPA Liberação (Col. M):", "capa_m", ["CAPA", "DIAGRAMADO", "OK", "CAPA ENTREGUE", ""]),
        ]
        vars_modal = {}
        for idx, (letra, rotulo, chave, opcoes) in enumerate(campos_def):
            ttk.Label(form, text=rotulo, style="Card.TLabel", font=FONTE_NEGRITO).grid(
                row=idx, column=0, sticky="w", pady=4, padx=(0, 10)
            )
            v = tk.StringVar(value=aluno.get(chave, ""))
            vars_modal[letra] = (chave, v)
            if opcoes is not None:
                cb = ttk.Combobox(form, textvariable=v, values=opcoes)
                cb.grid(row=idx, column=1, sticky="we", pady=4)
            else:
                ent = ttk.Entry(form, textvariable=v)
                ent.grid(row=idx, column=1, sticky="we", pady=4)

        rodape = tk.Frame(janela, bg=CARTAO, padx=14, pady=10)
        rodape.pack(fill="x")

        def _salvar():
            alteracoes = {}
            resumo_mudancas = []
            for letra, (chave, var_tk) in vars_modal.items():
                novo_v = var_tk.get().strip()
                antigo_v = str(aluno.get(chave, "") or "").strip()
                if novo_v != antigo_v:
                    alteracoes[letra] = novo_v
                    resumo_mudancas.append(f"• Coluna {letra} ({chave}): '{antigo_v}' -> '{novo_v}'")

            if not alteracoes:
                janela.destroy()
                return

            if not messagebox.askyesno(
                "Confirmar Reescrita na Planilha",
                f"Gravar as seguintes alterações na Linha {aluno['linha']} ({aluno['nome']})?\n\n"
                + "\n".join(resumo_mudancas),
                parent=janela,
            ):
                return

            janela.destroy()
            self._executar_reescrita_planilha(
                aluno,
                alteracoes,
                f"Edição completa ({', '.join(alteracoes.keys())}) de {vars_modal['C'][1].get()}",
            )

        ttk.Button(rodape, text="Cancelar", command=janela.destroy).pack(side="right", padx=(6, 0))
        ttk.Button(
            rodape,
            text="💾 Salvar e Reescrever na Planilha Agora",
            style="Sucesso.TButton",
            command=_salvar,
        ).pack(side="right")

    def _abrir_modal_webhook(self):
        janela = tk.Toplevel(self.root)
        janela.title("Configurar Gravação em Tempo Real no Google Sheets (Apps Script + Chrome)")
        janela.geometry("680x460")
        janela.configure(bg=CARTAO)
        janela.transient(self.root)
        janela.grab_set()

        corpo = ttk.Frame(janela, padding=14, style="Card.TFrame")
        corpo.pack(fill="both", expand=True)
        corpo.columnconfigure(0, weight=1)

        ttk.Label(
            corpo,
            text="⚡ Sincronização Instantânea com a Planilha Google (GERAL ENVIOS)",
            style="Card.TLabel",
            font=FONTE_TITULO,
        ).grid(row=0, column=0, sticky="w", pady=(0, 6))

        ttk.Label(
            corpo,
            text="1. Cole abaixo a URL do App da Web do Google Apps Script (para gravação instantânea em <0.5s sem abrir navegador)\n"
            "   OU marque 'Gravar também via Chrome no Google Sheets' no painel para o robô digitar direto na planilha.",
            style="Card.TLabel",
            wraplength=640,
        ).grid(row=1, column=0, sticky="w", pady=(0, 6))

        ttk.Label(corpo, text="URL do Webhook Google Apps Script:", style="Card.TLabel", font=FONTE_NEGRITO).grid(
            row=2, column=0, sticky="w"
        )
        ttk.Entry(corpo, textvariable=self.webhook_apps_script).grid(row=3, column=0, sticky="we", pady=(4, 10))

        ttk.Label(
            corpo,
            text="Código pronto do Google Apps Script (Extensões > Apps Script na planilha):",
            style="Card.TLabel",
            font=FONTE_NEGRITO,
        ).grid(row=4, column=0, sticky="w")

        caixa_cod = scrolledtext.ScrolledText(corpo, height=11, font=("Consolas", 9))
        caixa_cod.grid(row=5, column=0, sticky="nsew", pady=(4, 10))
        caixa_cod.insert("1.0", CODIGO_APPS_SCRIPT_WEBHOOK)
        caixa_cod.configure(state="disabled")

        botoes = tk.Frame(corpo, bg=CARTAO)
        botoes.grid(row=6, column=0, sticky="we")

        def _copiar():
            self.root.clipboard_clear()
            self.root.clipboard_append(CODIGO_APPS_SCRIPT_WEBHOOK)
            messagebox.showinfo(
                "Código Copiado",
                "Código copiado para a área de transferência!\n\n"
                "Na sua planilha Google:\n"
                "1. Clique em Extensões > Apps Script\n"
                "2. Cole o código e clique em Implantar > Nova implantação > App da Web\n"
                "3. Em 'Quem pode acessar', escolha 'Qualquer pessoa' e cole a URL gerada aqui.",
                parent=janela,
            )

        def _salvar_wh():
            self._salvar_config()
            janela.destroy()
            self.status.set("Configuração de gravação em tempo real salva.")

        ttk.Button(botoes, text="📋 Copiar Código Apps Script", command=_copiar).pack(side="left")
        ttk.Button(botoes, text="Salvar Configuração", style="Primario.TButton", command=_salvar_wh).pack(side="right")

    # --------------------------------------------------------------------------
    # DISPARO INDIVIDUAL E EM LOTE: ENVIO MAISBOOK E BAIXA SGE
    # --------------------------------------------------------------------------
    def _conferir_todos_no_servidor(self):
        """Conferência física em segundo plano no servidor (Fase 1) para todos os alunos DIAGRAMADO."""
        pasta_orig = self.origem.get().strip() or ORIGEM_PADRAO
        if not Path(pasta_orig).is_dir():
            messagebox.showwarning(
                "Conferência no Servidor",
                f"A pasta do servidor não está acessível neste computador agora:\n\n{pasta_orig}\n\n"
                "Verifique o caminho na guia 'Buscar e copiar' antes de conferir fisicamente.",
            )
            return

        self.status.set("Conferindo pastas de lâminas JPG e capas no servidor para os alunos DIAGRAMADO...")
        self._escrever("\n>>> [CONFERÊNCIA SERVIDOR] Verificando capa e lâminas JPG de quem está DIAGRAMADO...\n")

        def _worker():
            try:
                mod_envio = SourceFileLoader("envio_mb_check", str(SCRIPT_ENVIO)).load_module()
                albuns, indice = mod_envio.preparar_indice()
                diagramados = [a for a in self.alunos_dashboard if a["etapa_k"] == "DIAGRAMADO"]
                ok_count = 0
                falta_count = 0
                for a in diagramados:
                    tamanhos = mod_envio.tamanhos_do_plano(a["tamanho"])
                    pessoa = {
                        "nome": a["nome"],
                        "sge": a["sge"],
                        "contrato": mod_envio.so_contrato(a["contrato"]),
                        "tamanhos": tamanhos,
                        "planos": [a["tamanho"]] if a["tamanho"] else [],
                    }
                    achados, msg, resultado = mod_envio.localizar_albuns(pessoa, albuns, indice)
                    if resultado:
                        self.status_servidor_por_sge[a["sge"]] = f"⚠ {msg}"
                        falta_count += 1
                    else:
                        resumo_jpg = ", ".join(f"{it['tamanho']}:{it['laminas']}jpg" for it in (achados or []))
                        self.status_servidor_por_sge[a["sge"]] = f"✔ OK ({resumo_jpg} + capa)"
                        ok_count += 1

                self.root.after(0, self._filtrar_tabela_dashboard)
                self.root.after(
                    0,
                    lambda: self.status.set(
                        f"Conferência no servidor concluída: {ok_count} completo(s) | {falta_count} com arquivo faltando."
                    ),
                )
                self.root.after(
                    0,
                    lambda: self._escrever(
                        f"    ✔ Conferência concluída: {ok_count} completo(s) | {falta_count} pendente(s) no servidor.\n"
                    ),
                )
            except Exception as e:
                self.root.after(0, lambda: self.status.set(f"Aviso na conferência do servidor: {e}"))
                self.root.after(0, lambda: self._escrever(f"    ⚠ Aviso servidor: {e}\n"))

        threading.Thread(target=_worker, daemon=True).start()

    def _dash_conferir_servidor_individual(self):
        aluno = self._aluno_selecionado_dashboard()
        if not aluno:
            messagebox.showinfo("Conferir no Servidor", "Selecione um formando na tabela primeiro.")
            return
        pasta_orig = self.origem.get().strip() or ORIGEM_PADRAO
        if not Path(pasta_orig).is_dir():
            messagebox.showwarning(
                "Servidor não acessível",
                f"A pasta do servidor não foi encontrada neste computador:\n\n{pasta_orig}",
            )
            return

        self.status.set(f"Procurando álbum e capa de {aluno['nome']} (SGE {aluno['sge']}) no servidor...")

        def _worker():
            try:
                mod_envio = SourceFileLoader("envio_mb_single", str(SCRIPT_ENVIO)).load_module()
                albuns, indice = mod_envio.preparar_indice()
                tamanhos = mod_envio.tamanhos_do_plano(aluno["tamanho"])
                pessoa = {
                    "nome": aluno["nome"],
                    "sge": aluno["sge"],
                    "contrato": mod_envio.so_contrato(aluno["contrato"]),
                    "tamanhos": tamanhos,
                    "planos": [aluno["tamanho"]] if aluno["tamanho"] else [],
                }
                achados, msg, resultado = mod_envio.localizar_albuns(pessoa, albuns, indice)
                if resultado:
                    self.status_servidor_por_sge[aluno["sge"]] = f"⚠ {msg}"
                    texto_msg = (
                        f"⚠ FALTAM ARQUIVOS NO SERVIDOR PARA:\n{aluno['nome']} (SGE {aluno['sge']})\n\n"
                        f"Detalhe: {msg}\n\n"
                        "O envio para o MaisBook ficará bloqueado por segurança até que a capa e todas as lâminas estejam na pasta."
                    )
                    self.root.after(0, lambda: messagebox.showwarning("Conferência no Servidor", texto_msg))
                else:
                    resumo_jpg = ", ".join(f"{it['tamanho']}:{it['laminas']}/{it['esperadas']} JPG" for it in (achados or []))
                    self.status_servidor_por_sge[aluno["sge"]] = f"✔ OK ({resumo_jpg} + capa)"
                    texto_msg = (
                        f"✔ ÁLBUM E CAPA COMPLETOS NO SERVIDOR!\n\n"
                        f"Formando: {aluno['nome']} (SGE {aluno['sge']})\n"
                        f"Arquivos conferidos: {resumo_jpg} + Capa encontrada."
                    )
                    self.root.after(0, lambda: messagebox.showinfo("Conferência no Servidor", texto_msg))
                self.root.after(0, self._filtrar_tabela_dashboard)
            except Exception as e:
                self.root.after(0, lambda: messagebox.showerror("Erro ao conferir no servidor", str(e)))

        threading.Thread(target=_worker, daemon=True).start()

    def _dash_enviar_prontos_lote(self):
        if self.processo:
            messagebox.showinfo("Robô em execução", "Já existe uma operação em andamento. Aguarde ou clique em Parar.")
            return
        prontos = [a for a in self.alunos_dashboard if a["etapa_k"] == "DIAGRAMADO" and a["liberado_m"]]
        retidos = [a for a in self.alunos_dashboard if a["etapa_k"] == "DIAGRAMADO" and not a["liberado_m"]]
        if not messagebox.askyesno(
            "Disparo em Lote — Envio de Aprovação (MaisBook)",
            f"Prontos com STATUS = DIAGRAMADO e OK na Coluna M (CAPA): {len(prontos)} formando(s)\n"
            f"Retidos aguardando OK na Coluna M: {len(retidos)} formando(s)\n\n"
            "O robô fará:\n"
            "1. Conferência anti-erro no servidor (capa + todas as lâminas JPG)\n"
            "2. Upload para ALBUM APROVAÇAO no Drive e envio do e-mail no MaisBook\n"
            "3. Avanço da Coluna K para 'Aguardando Aprovação' e baixa no SGE\n\n"
            "Iniciar envio em lote agora?",
        ):
            return
        self.envio_sge.set("")
        argumentos = self._envio_argumentos()
        if argumentos is None:
            return
        if self.envio_simular.get():
            argumentos.append("--simular")
        self._envio_executar(argumentos, "Enviando lote de álbuns prontos (DIAGRAMADO + CAPA) no MaisBook...")

    def _dash_enviar_maisbook_individual(self):
        if self.processo:
            messagebox.showinfo("Robô em execução", "Já existe uma operação em andamento.")
            return
        aluno = self._aluno_selecionado_dashboard()
        if not aluno:
            messagebox.showinfo("Enviar MaisBook Individual", "Selecione um formando na tabela primeiro.")
            return
        aviso_extra = ""
        if not aluno["liberado_m"]:
            aviso_extra = (
                f"\n⚠ ATENÇÃO: A Coluna M (CAPA) deste aluno está '{aluno['capa_m'] or '(vazia)'}'.\n"
                "Deseja liberar e enviar este aluno mesmo assim (o robô ainda conferirá se a capa e as lâminas existem no servidor)?"
            )
        if not messagebox.askyesno(
            "Enviar Álbum Individual — MaisBook",
            f"Formando: {aluno['nome']}\n"
            f"SGE: {aluno['sge']} | Tamanho: {aluno['tamanho']}\n"
            f"Status (Col. K): {aluno['status_k']} | Coluna M (CAPA): {aluno['capa_m'] or '(vazia)'}\n"
            f"{aviso_extra}\n\n"
            "O robô vai conferir a pasta e a capa no servidor antes de subir no Drive e enviar no MaisBook. Continuar?",
        ):
            return
        args = ["--sge", str(aluno["sge"]), "--ignorar-coluna-m"]
        if not self.envio_atualizar_planilha.get():
            args.append("--sem-planilha")
        if not self.envio_baixa_sge.get():
            args.append("--sem-sge")
        if self.envio_simular.get():
            args.append("--simular")
        self._envio_executar(args, f"Enviando álbum de {aluno['nome']} (SGE {aluno['sge']}) no MaisBook...")

    def _dash_executar_baixa_sge_lote(self):
        if self.processo:
            messagebox.showinfo("Robô em execução", "Já existe uma operação em andamento.")
            return
        if not SCRIPT_BAIXA.is_file():
            messagebox.showerror("Baixa no SGE", f"Script não encontrado:\n{SCRIPT_BAIXA}")
            return
        etapa = self.dash_etapa_sge.get().strip()
        link = self.link.get().strip() or LINK_PLANILHA_PADRAO
        aba = self.aba.get().strip() or ABA_PADRAO

        if not messagebox.askyesno(
            "Disparo em Lote — Baixa no SGE",
            f"Executar o robô de Baixa no SGE (BAIXA.PY) em lote?\n\n"
            f"Planilha: {aba}\n"
            f"Modo de Status: {etapa}\n"
            f"Regra anti-regressão ativa: itens já adiantados no SGE não voltam para trás.\n\n"
            "Continuar?",
        ):
            return

        args = ["--link", link, "--aba", aba, "--col-sge", "E", "--col-nome", "C", "--col-status", "K", "--col-item", "F"]
        if etapa and etapa != "(usar o da planilha)":
            args += ["--status", etapa]
        if self.envio_simular.get():
            args.append("--simular")
        self._baixa_sge_executar(args, f"Executando Baixa no SGE em lote ({etapa})...")

    def _dash_vigia_sge(self):
        if self.processo:
            messagebox.showinfo("Robô em execução", "Já existe uma operação em andamento.")
            return
        link = self.link.get().strip() or LINK_PLANILHA_PADRAO
        aba = self.aba.get().strip() or ABA_PADRAO
        etapa = self.dash_etapa_sge.get().strip()
        args = [
            "--vigia",
            "--link", link,
            "--aba", aba,
            "--col-sge", "E",
            "--col-nome", "C",
            "--col-status", "K",
            "--col-item", "F",
        ]
        if etapa and etapa != "(usar o da planilha)":
            args += ["--status", etapa]
        if self.envio_simular.get():
            args.append("--simular")
        self._baixa_sge_executar(args, "Vigia Contínuo do SGE ativo — monitorando planilha...")

    def _dash_baixa_sge_individual(self):
        if self.processo:
            messagebox.showinfo("Robô em execução", "Já existe uma operação em andamento.")
            return
        aluno = self._aluno_selecionado_dashboard()
        if not aluno:
            messagebox.showinfo("Baixa no SGE Individual", "Selecione um formando na tabela primeiro.")
            return
        etapa = self.dash_etapa_sge.get().strip()
        link = self.link.get().strip() or LINK_PLANILHA_PADRAO
        aba = self.aba.get().strip() or ABA_PADRAO

        if not messagebox.askyesno(
            "Baixa Individual no SGE",
            f"Dar baixa no SGE apenas para:\n\n"
            f"Formando: {aluno['nome']}\n"
            f"SGE: {aluno['sge']} | Item/Tamanho: {aluno['tamanho']}\n"
            f"Status na Coluna K: {aluno['status_k']}\n"
            f"Status a aplicar no SGE: {etapa}\n\n"
            "Continuar?",
        ):
            return

        args = [
            "--link", link,
            "--aba", aba,
            "--col-sge", "E",
            "--col-nome", "C",
            "--col-status", "K",
            "--col-item", "F",
            "--sge", str(aluno["sge"]),
        ]
        if etapa and etapa != "(usar o da planilha)":
            args += ["--status", etapa]
        if self.envio_simular.get():
            args.append("--simular")
        self._baixa_sge_executar(args, f"Dando baixa no SGE para {aluno['nome']} (SGE {aluno['sge']})...")

    def _baixa_sge_executar(self, argumentos, status_msg):
        if not SCRIPT_BAIXA.is_file():
            messagebox.showerror("Baixa SGE", f"Script não encontrado:\n{SCRIPT_BAIXA}")
            return
        PARAR_BAIXA.unlink(missing_ok=True)
        self.pediu_parar_envio = False
        self.modo = "baixa_sge"
        self.inicio_operacao = time.time()
        self.resumo = None
        self._limpar()
        self._escrever("> " + SCRIPT_BAIXA.name + " " + " ".join(argumentos) + "\n")
        self._habilitar(False)
        self.status.set(status_msg)
        ambiente = dict(os.environ)
        ambiente["PYTHONIOENCODING"] = "utf-8"
        ambiente["PYTHONUNBUFFERED"] = "1"
        try:
            self.processo = subprocess.Popen(
                [python_console(), "-u", str(SCRIPT_BAIXA)] + argumentos,
                cwd=str(SCRIPT_BAIXA.parent),
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                stdin=subprocess.DEVNULL,
                env=ambiente,
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
            )
        except Exception as e:
            self.processo = None
            self.modo = None
            self._habilitar(True)
            self.status.set("Não foi possível iniciar o robô do SGE.")
            self._escrever(str(e) + "\n")
            return
        threading.Thread(target=self._ler_saida, daemon=True).start()

    def _parar_qualquer_robo(self):
        if not self.processo:
            return
        if self.modo == "baixa_sge":
            PARAR_BAIXA.write_text("parar", encoding="utf-8")
        else:
            PARAR_ENVIO.write_text("parar", encoding="utf-8")
        if not self.pediu_parar_envio:
            self.pediu_parar_envio = True
            self.btn_dash_parar_tudo.configure(text="■ Forçar Parada")
            self.botao_envio_parar.configure(text="Forçar parada")
            self.status.set("Parando ao terminar o formando atual...")
            self._escrever("\n>>> Pedido enviado para parar ao terminar o formando atual.\n")
            return
        if messagebox.askyesno("Forçar Parada", "Encerrar imediatamente o processo do robô?"):
            subprocess.run(
                ["taskkill", "/PID", str(self.processo.pid), "/T", "/F"],
                capture_output=True,
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
            )

    # ==========================================================================
    # GUIAS ORIGINAIS (BUSCAR E COPIAR / RENOMEAR / ENVIAR APROVAÇÃO)
    # ==========================================================================
    def _combo_coluna(self, pai, variavel):
        return ttk.Combobox(pai, textvariable=variavel, values=LETRAS, state="readonly", width=4)

    def _montar_busca(self, caderno):
        aba = ttk.Frame(caderno, padding=8)
        caderno.add(aba, text="📁 Buscar e copiar")
        aba.columnconfigure(0, weight=1)

        procura = ttk.LabelFrame(aba, text="Onde procurar", padding=6)
        procura.grid(row=0, column=0, sticky="we")
        procura.columnconfigure(1, weight=1)

        ttk.Label(procura, text="Pasta", style="Card.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Entry(procura, textvariable=self.origem).grid(row=0, column=1, sticky="we", padx=4)
        ttk.Button(procura, text="Escolher pasta...", command=self._escolher_origem).grid(row=0, column=2)
        ttk.Label(
            procura,
            text="Capas e álbuns são procurados só dentro desta pasta.",
            style="Fraco.TLabel",
        ).grid(row=1, column=1, sticky="w", padx=4, pady=(4, 0))

        dados = ttk.LabelFrame(aba, text="Planilha e destino", padding=6)
        dados.grid(row=1, column=0, sticky="we", pady=(8, 0))
        dados.columnconfigure(1, weight=1)

        ttk.Label(dados, text="Link do Google", style="Card.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Entry(dados, textvariable=self.link).grid(row=0, column=1, columnspan=2, sticky="we", padx=4)

        ttk.Label(dados, text="Ou arquivo", style="Card.TLabel").grid(row=1, column=0, sticky="w", pady=(6, 0))
        ttk.Entry(dados, textvariable=self.planilha).grid(row=1, column=1, sticky="we", padx=4, pady=(6, 0))
        ttk.Button(dados, text="Escolher...", command=self._escolher_planilha).grid(row=1, column=2, pady=(6, 0))

        ttk.Label(dados, text="Aba", style="Card.TLabel").grid(row=2, column=0, sticky="w", pady=(6, 0))
        self.combo_aba = ttk.Combobox(dados, textvariable=self.aba, width=28)
        self.combo_aba.grid(row=2, column=1, sticky="w", padx=4, pady=(6, 0))
        self.combos_aba.append(self.combo_aba)
        self.botao_abas = ttk.Button(dados, text="Atualizar abas", command=self._carregar_abas)
        self.botao_abas.grid(row=2, column=2, pady=(6, 0))

        colunas = ttk.Frame(dados, style="Card.TFrame")
        colunas.grid(row=3, column=1, columnspan=2, sticky="w", pady=(6, 0))
        ttk.Label(dados, text="Colunas", style="Card.TLabel").grid(row=3, column=0, sticky="w", pady=(6, 0))
        ttk.Label(colunas, text="Nome", style="Card.TLabel").pack(side="left")
        self._combo_coluna(colunas, self.col_nome).pack(side="left", padx=(4, 12))
        ttk.Label(colunas, text="Código SGE", style="Card.TLabel").pack(side="left")
        self._combo_coluna(colunas, self.col_codigo).pack(side="left", padx=(4, 12))
        ttk.Label(colunas, text="Plano", style="Card.TLabel").pack(side="left")
        self.combo_col_plano = ttk.Combobox(
            colunas,
            textvariable=self.col_plano,
            values=LETRAS,
            width=4,
            state="readonly",
        )
        self.combo_col_plano.pack(side="left", padx=(4, 0))

        ttk.Label(dados, text="Destino", style="Card.TLabel").grid(row=4, column=0, sticky="w", pady=(6, 0))
        ttk.Entry(dados, textvariable=self.destino).grid(row=4, column=1, sticky="we", padx=4, pady=(6, 0))
        destino_botoes = ttk.Frame(dados, style="Card.TFrame")
        destino_botoes.grid(row=4, column=2, pady=(6, 0))
        ttk.Button(destino_botoes, text="Escolher...", command=self._escolher_destino).pack(side="left")
        ttk.Button(destino_botoes, text="Abrir", command=self._abrir_destino).pack(side="left", padx=(4, 0))

        ttk.Label(dados, text="Filtrar plano", style="Card.TLabel").grid(row=5, column=0, sticky="w", pady=(6, 0))
        ttk.Entry(dados, textvariable=self.plano, width=32).grid(row=5, column=1, sticky="w", padx=4, pady=(6, 0))
        ttk.Label(
            dados,
            text="Opcional. Ex.: Essencial, Prime, Moscou, Berlim, Diamond, 20X20",
            style="Fraco.TLabel",
        ).grid(row=6, column=1, sticky="w", padx=4)

        acoes = ttk.Frame(aba)
        acoes.grid(row=2, column=0, sticky="we", pady=(8, 0))
        self.botao_simular = ttk.Button(acoes, text="Simular", command=lambda: self._rodar(True))
        self.botao_simular.pack(side="left")
        self.botao_copiar = ttk.Button(acoes, text="Copiar", style="Primario.TButton", command=lambda: self._rodar(False))
        self.botao_copiar.pack(side="left", padx=(8, 0))
        self.botao_relatorio = ttk.Button(acoes, text="Abrir relatório", command=self._abrir_relatorio)
        self.botao_relatorio.pack(side="left", padx=(8, 0))

        ttk.Label(
            aba,
            text="A capa e o álbum são procurados pelo nome da pessoa. "
            "Copiar grava o lote com esse nome e, em seguida, troca pelo código SGE. "
            "A guia Renomear faz só essa troca num lote que já existe.",
            wraplength=920,
        ).grid(row=3, column=0, sticky="w", pady=(8, 0))

    def _montar_renomeio(self, caderno):
        aba = ttk.Frame(caderno, padding=8)
        caderno.add(aba, text="🏷️ Renomear")
        aba.columnconfigure(0, weight=1)

        lote = ttk.LabelFrame(aba, text="Lote", padding=6)
        lote.grid(row=0, column=0, sticky="we")
        lote.columnconfigure(1, weight=1)
        ttk.Label(lote, text="Pasta do lote", style="Card.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Entry(lote, textvariable=self.destino).grid(row=0, column=1, sticky="we", padx=4)
        ttk.Button(lote, text="Escolher...", command=self._escolher_destino).grid(row=0, column=2)
        ttk.Button(lote, text="Abrir", command=self._abrir_destino).grid(row=0, column=3, padx=(4, 0))
        ttk.Label(lote, textvariable=self.pastas_existentes, style="Fraco.TLabel").grid(
            row=1, column=1, columnspan=3, sticky="w", pady=(6, 0)
        )

        colunas = ttk.LabelFrame(aba, text="Planilha — nome e código SGE", padding=6)
        colunas.grid(row=1, column=0, sticky="we", pady=(8, 0))
        colunas.columnconfigure(6, weight=1)
        ttk.Label(colunas, text="Aba", style="Card.TLabel").grid(row=0, column=0, sticky="w")
        combo_aba = ttk.Combobox(colunas, textvariable=self.aba, state="readonly", width=22)
        combo_aba.grid(row=0, column=1, sticky="w", padx=(4, 12))
        self.combos_aba.append(combo_aba)
        ttk.Label(colunas, text="Coluna do nome", style="Card.TLabel").grid(row=0, column=2, sticky="w")
        self._combo_coluna(colunas, self.col_nome).grid(row=0, column=3, sticky="w", padx=(4, 12))
        ttk.Label(colunas, text="Coluna do código SGE", style="Card.TLabel").grid(row=0, column=4, sticky="w")
        self._combo_coluna(colunas, self.col_codigo).grid(row=0, column=5, sticky="w", padx=(4, 12))

        self.previa = ttk.Treeview(colunas, columns=("linha", "nome", "codigo"), show="headings", height=4)
        for col, titulo, largura in (("linha", "Linha", 60), ("nome", "Nome", 420), ("codigo", "Código SGE", 160)):
            self.previa.heading(col, text=titulo)
            self.previa.column(col, width=largura, anchor="w", stretch=(col == "nome"))
        self.previa.grid(row=1, column=0, columnspan=7, sticky="we", pady=(6, 0))

        escopo = ttk.LabelFrame(aba, text="O que processar", padding=6)
        escopo.grid(row=2, column=0, sticky="we", pady=(8, 0))
        ttk.Label(escopo, text="Tamanhos:", style="Card.TLabel").pack(side="left")
        for tamanho in TAMANHOS:
            ttk.Checkbutton(escopo, text=tamanho, variable=self.tamanhos[tamanho]).pack(side="left", padx=4)
        ttk.Separator(escopo, orient="vertical").pack(side="left", fill="y", padx=12)
        ttk.Label(escopo, text="Categorias:", style="Card.TLabel").pack(side="left")
        for categoria in CATEGORIAS:
            ttk.Checkbutton(escopo, text=categoria, variable=self.categorias[categoria]).pack(side="left", padx=4)

        acoes = ttk.Frame(aba)
        acoes.grid(row=3, column=0, sticky="we", pady=(8, 0))
        self.botao_simular_renomeio = ttk.Button(acoes, text="Simular", command=lambda: self._renomear(True))
        self.botao_simular_renomeio.pack(side="left")
        self.botao_renomear = ttk.Button(acoes, text="Renomear", style="Primario.TButton", command=lambda: self._renomear(False))
        self.botao_renomear.pack(side="left", padx=6)
        ttk.Label(
            aba,
            text="Usa a mesma planilha da outra guia. Simular não muda nenhum arquivo. "
            "Renomear troca capas, estojos e lâminas pelo código SGE, só nos tamanhos e categorias marcados. "
            "30X22 (e 30X22 PÓS-VENDA): capa em FRENTE e VERSO vira código_001 e código_002; sem estojo; "
            "qualquer quantidade de lâminas. POSTERS: um por aluno vira código; mais de um, código_001, código_002.",
            wraplength=920,
        ).grid(row=4, column=0, sticky="w", pady=(8, 0))

    def _montar_envio(self, caderno):
        aba = ttk.Frame(caderno, padding=8)
        caderno.add(aba, text="📤 Enviar aprovação (Avançado)")
        aba.columnconfigure(0, weight=1)

        quem = ttk.LabelFrame(aba, text="Quem entra", padding=6)
        quem.grid(row=0, column=0, sticky="we")
        quem.columnconfigure(0, weight=1)
        ttk.Label(
            quem,
            style="Card.TLabel",
            text="1º Filtro: STATUS = DIAGRAMADO (Coluna K) + OK na Coluna M (CAPA = DIAGRAMADO / OK / CAPA ENTREGUE).\n"
            "2º Conferência Anti-Erro no Servidor (Fase 1): Mesmo com OK na Coluna M, confere antes no servidor se a capa "
            "e todas as lâminas JPG estão completas. Se você tiver dado OK por engano e faltar arquivo, o envio é bloqueado.\n"
            "3º Envio (Fase 2): Sobe para ALBUM APROVAÇAO no Drive, cola o link no MaisBook, envia o e-mail, "
            "avança a Coluna K para 'Aguardando Aprovação' e dá baixa no SGE como 'Enviado para Aprovação'.",
            wraplength=920,
        ).grid(row=0, column=0, sticky="w")

        filtros = ttk.LabelFrame(aba, text="Filtro e automação", padding=6)
        filtros.grid(row=1, column=0, sticky="we", pady=(8, 0))
        ttk.Label(filtros, text="SGE", style="Card.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Entry(filtros, textvariable=self.envio_sge, width=14).grid(row=0, column=1, sticky="w", padx=(4, 16))
        ttk.Label(filtros, text="Limite", style="Card.TLabel").grid(row=0, column=2, sticky="w")
        ttk.Entry(filtros, textvariable=self.envio_limite, width=6).grid(row=0, column=3, sticky="w", padx=(4, 16))
        ttk.Label(filtros, text="Intervalo vigia (s)", style="Card.TLabel").grid(row=0, column=4, sticky="w")
        ttk.Entry(filtros, textvariable=self.envio_intervalo, width=6).grid(row=0, column=5, sticky="w", padx=(4, 16))
        ttk.Checkbutton(
            filtros,
            text="Simular (não sobe, não cola e não envia)",
            variable=self.envio_simular,
        ).grid(row=0, column=6, sticky="w")

        opcoes_auto = ttk.Frame(filtros, style="Card.TFrame")
        opcoes_auto.grid(row=1, column=0, columnspan=7, sticky="w", pady=(6, 0))
        ttk.Checkbutton(
            opcoes_auto,
            text="Exigir OK na Coluna M (CAPA) para enviar",
            variable=self.envio_exigir_coluna_m,
        ).pack(side="left", padx=(0, 16))
        ttk.Checkbutton(
            opcoes_auto,
            text="Atualizar Coluna K ('Aguardando Aprovação')",
            variable=self.envio_atualizar_planilha,
        ).pack(side="left", padx=(0, 16))
        ttk.Checkbutton(
            opcoes_auto,
            text="Baixa no SGE ('Enviado para Aprovação')",
            variable=self.envio_baixa_sge,
        ).pack(side="left")

        ttk.Label(
            filtros,
            text="Mesmo com o OK na Coluna M (CAPA), o robô confere fisicamente a capa e todas as lâminas no servidor antes de enviar.",
            style="Fraco.TLabel",
        ).grid(row=2, column=0, columnspan=7, sticky="w", pady=(6, 0))

        acoes = ttk.Frame(aba)
        acoes.grid(row=2, column=0, sticky="we", pady=(8, 0))
        self.botao_envio_iniciar = ttk.Button(acoes, text="Iniciar", style="Sucesso.TButton", command=self._envio_iniciar)
        self.botao_envio_iniciar.pack(side="left")
        self.botao_envio_vigia = ttk.Button(
            acoes, text="Iniciar Vigia Contínuo", style="Primario.TButton", command=self._envio_vigia
        )
        self.botao_envio_vigia.pack(side="left", padx=(8, 0))
        self.botao_envio_continuar = ttk.Button(
            acoes, text="Continuar de onde parou", command=self._envio_continuar
        )
        self.botao_envio_continuar.pack(side="left", padx=(8, 0))
        self.botao_envio_fila = ttk.Button(acoes, text="Ver fila (DIAGRAMADO)", command=self._envio_listar)
        self.botao_envio_fila.pack(side="left", padx=(8, 0))
        self.botao_envio_parar = ttk.Button(
            acoes, text="Parar", command=self._parar_qualquer_robo, state="disabled", style="Perigo.TButton"
        )
        self.botao_envio_parar.pack(side="left", padx=(8, 0))
        self.botao_envio_relatorio = ttk.Button(acoes, text="Abrir relatório", command=self._abrir_relatorio_envio)
        self.botao_envio_relatorio.pack(side="left", padx=(8, 0))

    def _escolher_planilha(self):
        caminho = filedialog.askopenfilename(
            title="Planilha",
            filetypes=[("Excel", "*.xlsx"), ("Todos", "*.*")],
        )
        if caminho:
            self.planilha.set(caminho)

    def _escolher_origem(self):
        caminho = filedialog.askdirectory(title="Pasta onde procurar capas e álbuns")
        if caminho:
            self.origem.set(caminho)

    def _escolher_destino(self):
        caminho = filedialog.askdirectory(title="Pasta de destino")
        if not caminho:
            return
        self.destino.set(os.path.normpath(caminho))
        credenciais = Path(caminho) / "CREDENCIAIS.xlsx"
        if credenciais.is_file() and not self.planilha.get().strip() and not self.link.get().strip():
            self.planilha.set(str(credenciais))
        self._atualizar_pastas()

    def _abrir_destino(self):
        pasta = self.destino.get().strip()
        if pasta and Path(pasta).is_dir():
            os.startfile(pasta)
        else:
            messagebox.showinfo("Organizador", "A pasta de destino ainda não existe.")

    def _atualizar_pastas(self):
        lote = Path(self.destino.get().strip()) if self.destino.get().strip() else None
        if not lote or not lote.is_dir():
            self.pastas_existentes.set("A pasta do lote ainda não existe.")
            return
        partes = []
        for tamanho in TAMANHOS:
            if (lote / tamanho).is_dir():
                cats = [c for c in CATEGORIAS if (lote / tamanho / c).is_dir()]
                partes.append(f"{tamanho}: {', '.join(cats) if cats else 'sem subpastas'}")
            else:
                partes.append(f"{tamanho}: não existe")
        self.pastas_existentes.set("   |   ".join(partes))

    def _caminho_planilha_local(self):
        arquivo = self.planilha.get().strip()
        if arquivo and Path(arquivo).is_file():
            return Path(arquivo)
        if self.link.get().strip() and PLANILHA_BAIXADA.is_file():
            return PLANILHA_BAIXADA
        return None

    def _atualizar_previa(self):
        if not hasattr(self, "previa"):
            return
        self.previa.delete(*self.previa.get_children())
        caminho = self._caminho_planilha_local()
        if not caminho or not self.aba.get().strip():
            return
        try:
            wb = openpyxl.load_workbook(caminho, read_only=True, data_only=True)
            if self.aba.get() not in wb.sheetnames:
                wb.close()
                return
            ws = wb[self.aba.get()]
            i_nome = LETRAS.index(self.col_nome.get())
            i_cod = LETRAS.index(self.col_codigo.get())
            mostradas = 0
            total = 0
            for n, linha in enumerate(ws.iter_rows(values_only=True), start=1):
                nome = linha[i_nome] if len(linha) > i_nome else None
                cod = linha[i_cod] if len(linha) > i_cod else None
                if nome in (None, "") and cod in (None, ""):
                    continue
                total += 1
                if mostradas < 4:
                    self.previa.insert("", "end", values=(n, nome or "", cod or ""))
                    mostradas += 1
            wb.close()
        except Exception as e:
            self.status.set(f"Erro ao ler a prévia: {e}")

    def _salvar_config(self):
        dados = {
            "link": self.link.get().strip(),
            "planilha": self.planilha.get().strip(),
            "aba": self.aba.get().strip(),
            "col_plano": self.col_plano.get().strip(),
            "col_nome": self.col_nome.get().strip(),
            "col_codigo": self.col_codigo.get().strip(),
            "origem": self.origem.get().strip(),
            "destino": self.destino.get().strip(),
            "plano": self.plano.get().strip(),
            "tamanhos": [t for t, v in self.tamanhos.items() if v.get()],
            "categorias": [c for c, v in self.categorias.items() if v.get()],
            "envio_intervalo": int(self.envio_intervalo.get().strip() or 60),
            "envio_exigir_coluna_m": bool(self.envio_exigir_coluna_m.get()),
            "envio_atualizar_planilha": bool(self.envio_atualizar_planilha.get()),
            "envio_baixa_sge": bool(self.envio_baixa_sge.get()),
            "webhook_apps_script": self.webhook_apps_script.get().strip(),
            "gravar_via_chrome": bool(self.gravar_via_chrome.get()),
            "autosync_ativo": bool(self.autosync_ativo.get()),
        }
        CONFIG.write_text(json.dumps(dados, indent=2, ensure_ascii=False), encoding="utf-8")

    def _argumentos_planilha(self):
        link = self.link.get().strip()
        arquivo = self.planilha.get().strip()
        if link:
            return ["--link", link]
        if arquivo:
            return ["--planilha", arquivo]
        return None

    def _agendar_abas(self, *_):
        if self._timer_abas:
            self.root.after_cancel(self._timer_abas)
        self._timer_abas = self.root.after(700, lambda: self._carregar_abas_automatico(avisar=True))

    def _carregar_abas_automatico(self, avisar=False):
        self._timer_abas = None
        link = self.link.get().strip()
        arquivo = self.planilha.get().strip()
        if link and not RE_LINK.search(link):
            return
        if not link and not arquivo:
            return
        self._carregar_abas(avisar=avisar)

    def _carregar_abas(self, avisar=True):
        if self.processo:
            return
        args = self._argumentos_planilha()
        if not args:
            if avisar:
                messagebox.showwarning("Organizador", "Informe o link do Google ou escolha um arquivo.")
            return
        if args[0] == "--planilha" and not Path(args[1]).is_file():
            if avisar:
                messagebox.showerror("Organizador", "Arquivo não encontrado:\n\n" + args[1])
            return
        if args[0] == "--link" and not RE_LINK.search(args[1]):
            return

        self._geracao_abas += 1
        geracao = self._geracao_abas
        self.botao_abas.configure(state="disabled")
        threading.Thread(target=self._ler_abas, args=(args, geracao, avisar), daemon=True).start()

    def _ler_abas(self, args, geracao, avisar):
        abas = []
        erro = ""
        if args[0] == "--planilha":
            try:
                livro = openpyxl.load_workbook(args[1], read_only=True, data_only=True)
                abas = list(livro.sheetnames)
                livro.close()
            except Exception as e:
                erro = str(e)
        else:
            try:
                resultado = subprocess.run(
                    [python_console(), "-u", str(SCRIPT), "--sem-pausa", "--listar-abas"] + args,
                    cwd=str(PASTA),
                    capture_output=True,
                    env={**os.environ, "PYTHONIOENCODING": "utf-8", "PYTHONUNBUFFERED": "1"},
                    creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
                )
                saida = resultado.stdout.decode("utf-8", errors="replace")
                abas = [m.group(1) for linha in saida.splitlines() if (m := RE_ABA.match(linha))]
                if resultado.returncode != 0 and not abas:
                    erro = saida.strip() or resultado.stderr.decode("utf-8", errors="replace")
            except Exception as e:
                erro = str(e)
        self.root.after(0, self._preencher_abas, abas, erro, geracao, avisar)

    def _preencher_abas(self, abas, erro, geracao, avisar):
        if geracao != self._geracao_abas:
            return
        if not self.processo:
            self.botao_abas.configure(state="normal")
        if erro and not abas:
            if avisar:
                messagebox.showerror("Organizador", erro)
            return
        for combo in self.combos_aba:
            combo["values"] = abas
            if abas:
                combo.configure(state="readonly")
            else:
                combo.configure(state="normal")
        if abas and self.aba.get() not in abas:
            self.aba.set("GERAL ENVIOS" if "GERAL ENVIOS" in abas else abas[0])
        self._atualizar_previa()

    def _rodar(self, simular):
        if self.processo:
            return
        args_planilha = self._argumentos_planilha()
        origem = self.origem.get().strip()
        destino = self.destino.get().strip()
        if not args_planilha:
            messagebox.showwarning("Organizador", "Informe o link do Google ou escolha um arquivo.")
            return
        if not origem:
            messagebox.showwarning("Organizador", "Escolha a pasta onde procurar as capas e os álbuns.")
            return
        if not destino:
            messagebox.showwarning("Organizador", "Escolha a pasta de destino. Se ela não existir, será criada no modo Copiar.")
            return
        if not SCRIPT.is_file():
            messagebox.showerror("Organizador", "Script não encontrado:\n\n" + str(SCRIPT))
            return
        if not CONSOLE_PIRATA.is_file():
            messagebox.showerror("Organizador", "Script não encontrado:\n\n" + str(CONSOLE_PIRATA))
            return
        if not simular and not messagebox.askyesno(
            "Copiar",
            "Os arquivos serão COPIADOS para:\n\n"
            + destino
            + "\n\nSe essa pasta não existir, ela será criada com a organização\n"
            "tamanho / CAPAS, LAMINAS/{nome da pessoa} e ESTOJO (24X34 e 30X20).\n"
            "Depois da cópia, capa e lâminas são renomeadas com o código SGE.\n\n"
            "Os originais do servidor são mantidos. Continuar?",
        ):
            return

        self._salvar_config()
        args = args_planilha + [
            "--origem", origem,
            "--destino", destino,
            "--sem-pausa",
            "--simular" if simular else "--real",
        ]
        if self.aba.get().strip():
            args += ["--aba", self.aba.get().strip()]
        if self.col_plano.get().strip():
            args += ["--col-plano", self.col_plano.get().strip()]
        if self.col_nome.get().strip():
            args += ["--col-nome", self.col_nome.get().strip()]
        if self.col_codigo.get().strip():
            args += ["--col-codigo", self.col_codigo.get().strip()]
        if self.plano.get().strip():
            args += ["--plano", self.plano.get().strip()]

        self.resumo = None
        self.modo = "busca"
        self.inicio_operacao = time.time()
        self._limpar()
        self._escrever("> " + SCRIPT.name + " " + " ".join(f'"{a}"' if " " in a else a for a in args) + "\n")
        self._habilitar(False)
        self.status.set("Simulando..." if simular else "Copiando...")

        try:
            LOG_LOTE.unlink(missing_ok=True)
        except OSError:
            pass
        x, y, largura, altura = self._posicao_cmd()
        try:
            self.processo = subprocess.Popen(
                [
                    python_console(), "-u", str(CONSOLE_PIRATA),
                    "--log", str(LOG_LOTE),
                    "--x", str(x), "--y", str(y),
                    "--largura", str(largura), "--altura", str(altura),
                    "--titulo", "Separando o lote - " + ("Simular" if simular else "Copiar"),
                    "--", python_console(), "-u", str(SCRIPT),
                ] + args,
                cwd=str(PASTA),
                creationflags=getattr(subprocess, "CREATE_NEW_CONSOLE", 0),
            )
        except Exception as e:
            self.processo = None
            self._habilitar(True)
            self.status.set("Não foi possível iniciar o script.")
            self._escrever(str(e) + "\n")
            return

        threading.Thread(target=self._acompanhar_log, args=(self.processo,), daemon=True).start()

    def _posicao_cmd(self):
        self.root.update_idletasks()
        x = self.root.winfo_x()
        y = self.root.winfo_rooty() + self.root.winfo_height() + 8
        largura = self.root.winfo_width() + 16
        altura = ALTURA_CMD
        tela = self.root.winfo_screenheight()
        if y + altura > tela - 40:
            y = max(tela - 40 - altura, 0)
        return x, y, largura, altura

    def _acompanhar_log(self, processo):
        codigo = None
        arquivo = None
        sobra = ""
        try:
            while codigo is None:
                if arquivo is None:
                    if LOG_LOTE.is_file():
                        arquivo = open(LOG_LOTE, "r", encoding="utf-8", errors="replace")
                    elif processo.poll() is not None:
                        break
                    else:
                        time.sleep(0.1)
                        continue
                pedaco = arquivo.read()
                if not pedaco:
                    if processo.poll() is not None:
                        break
                    time.sleep(0.1)
                    continue
                sobra += pedaco
                linhas = sobra.split("\n")
                sobra = linhas.pop()
                texto = []
                for linha in linhas:
                    if linha.startswith("__FIM__"):
                        partes = linha.split()
                        codigo = int(partes[1]) if len(partes) > 1 and partes[1].lstrip("-").isdigit() else 1
                        break
                    texto.append(linha + "\n")
                if texto:
                    self.root.after(0, self._escrever, "".join(texto))
        finally:
            if arquivo:
                arquivo.close()
        if codigo is None:
            codigo = processo.poll()
            if codigo is None or codigo == 0:
                codigo = 1
            self.root.after(0, self._escrever, "\n⚠ O CMD foi fechado antes do fim do processo.\n")
        self.root.after(0, self._finalizar, codigo)

    def _ler_saida(self):
        processo = self.processo
        for bruto in iter(processo.stdout.readline, b""):
            self.root.after(0, self._escrever, bruto.decode("utf-8", errors="replace"))
        codigo = processo.wait()
        self.root.after(0, self._finalizar, codigo)

    def _gerar_planilha_erros(self, modo):
        try:
            if modo == "renomeio":
                arquivo = self.arquivo_erros
                if arquivo and arquivo.is_file():
                    livro = openpyxl.load_workbook(arquivo, read_only=True)
                    linhas = max(livro.active.max_row - 1, 0)
                    livro.close()
                    return arquivo, linhas
                return None, 0
            if modo == "busca":
                relatorio, operacao = RELATORIO, "organizador"
                eh_erro = lambda linha: (linha.get("STATUS") or "").strip().upper() != "OK"
            elif modo == "envio":
                relatorio, operacao = RELATORIO_ENVIO, "envio"
                eh_erro = lambda linha: (linha.get("RESULTADO") or "").strip().lower() not in RESULTADOS_ENVIO_OK
            else:
                return None, 0
            if not relatorio.is_file() or relatorio.stat().st_mtime < self.inicio_operacao:
                return None, 0
            destino = caminho_planilha_erros(operacao)
            linhas = planilha_erros_de_csv(relatorio, destino, eh_erro)
            return (destino, linhas) if linhas else (None, 0)
        except Exception as e:
            self._escrever(f"\n⚠ Não consegui gerar a planilha de erros: {e}\n")
            return None, 0

    def _finalizar(self, codigo):
        avisos = list(self.avisos)
        self.avisos = []
        modo_anterior = self.modo
        self.processo = None
        planilha_erros, linhas_erros = self._gerar_planilha_erros(modo_anterior)
        self.arquivo_erros = None
        self.modo = None
        PARAR_ENVIO.unlink(missing_ok=True)
        PARAR_BAIXA.unlink(missing_ok=True)
        self._habilitar(True)
        self._atualizar_pastas()
        self._atualizar_previa()
        if modo_anterior in ("envio", "baixa_sge"):
            self._sincronizar_dashboard(forcar_download=True, silencioso=True)
        if avisos:
            self.status.set(f"{len(avisos)} aviso(s): itens não encontrados ou pessoas duplicadas.")
        elif codigo != 0:
            self.status.set(f"Terminou com erro (código {codigo}). Veja a saída.")
        elif self.resumo:
            self.status.set("Concluído: " + self.resumo)
        else:
            self.status.set("Concluído.")
        if planilha_erros:
            self._escrever(f"\n>>> Planilha de erros ({linhas_erros} linha(s)): {planilha_erros}\n")
            self.status.set(self.status.get() + f"  |  Planilha de erros: {linhas_erros} linha(s)")
        if avisos:
            texto = "\n".join(
                item.replace("NÃO ENCONTREI: ", "Não encontrei: ").replace("DUPLICADO: ", "Duplicado: ")
                for item in avisos[:25]
            )
            if len(avisos) > 25:
                texto += f"\n\n... e mais {len(avisos) - 25}."
            if planilha_erros:
                texto += f"\n\nPlanilha de erros gerada ({linhas_erros} linha(s)). Abrir agora?"
                if messagebox.askyesno("Itens não encontrados", texto, icon="warning"):
                    os.startfile(planilha_erros)
                return
            messagebox.showwarning("Itens não encontrados", texto)
        if planilha_erros and messagebox.askyesno(
            "Planilha de erros",
            f"A operação terminou com {linhas_erros} erro(s).\n\n{planilha_erros}\n\nAbrir agora?",
            icon="warning",
        ):
            os.startfile(planilha_erros)

    def _abrir_relatorio(self):
        if not RELATORIO.is_file():
            messagebox.showinfo("Organizador", "Ainda não há relatório. Rode Simular ou Copiar primeiro.")
            return
        os.startfile(RELATORIO)

    def _abrir_relatorio_envio(self):
        if not RELATORIO_ENVIO.is_file():
            messagebox.showinfo("Enviar aprovação", "Ainda não há relatório. Rode o envio primeiro.")
            return
        os.startfile(RELATORIO_ENVIO)

    def _habilitar(self, habilitar):
        estado = "normal" if habilitar else "disabled"
        for btn in (
            self.botao_simular,
            self.botao_copiar,
            self.botao_abas,
            self.botao_simular_renomeio,
            self.botao_renomear,
            self.botao_envio_iniciar,
            self.botao_envio_vigia,
            self.botao_envio_continuar,
            self.botao_envio_fila,
            self.botao_envio_relatorio,
            self.btn_dash_enviar_lote,
            self.btn_dash_vigia_mb,
            self.btn_dash_baixa_lote,
            self.btn_dash_vigia_sge,
            self.btn_ind_enviar_mb,
            self.btn_ind_baixa_sge,
        ):
            btn.configure(state=estado)

        if habilitar or self.modo not in ("envio", "baixa_sge"):
            self.pediu_parar_envio = False
            self.botao_envio_parar.configure(state="disabled", text="Parar")
            self.btn_dash_parar_tudo.configure(state="disabled", text="■ Parar Robô")
        else:
            self.botao_envio_parar.configure(state="normal", text="Parar")
            self.btn_dash_parar_tudo.configure(state="normal", text="■ Parar Robô")

    def _limpar(self):
        self.avisos = []
        self.texto.configure(state="normal")
        self.texto.delete("1.0", "end")
        self.texto.configure(state="disabled")

    def _escrever(self, texto):
        achado = RE_RESUMO.search(texto)
        if achado:
            self.resumo = achado.group(1)
        renomeio = RE_RESUMO_RENOMEIO.search(texto)
        if renomeio:
            self.resumo = (
                f"{renomeio.group(1).capitalize()} {renomeio.group(2)} | Já ok {renomeio.group(3)} | "
                f"Não encontrados {renomeio.group(4)} | Erros {renomeio.group(5)}"
            )
        aviso = None
        for linha in texto.splitlines():
            if linha.strip().startswith(("NÃO ENCONTREI:", "DUPLICADO:")):
                aviso = linha.strip()
                self.avisos.append(aviso)
        if aviso:
            self.status.set(aviso)
        self.texto.configure(state="normal")
        if aviso:
            self.texto.insert("end", texto, "aviso")
        else:
            self.texto.insert("end", texto)
        self.texto.see("end")
        self.texto.configure(state="disabled")

    def _renomear(self, simular):
        if self.processo:
            return
        lote = self.destino.get().strip()
        tamanhos = [t for t, v in self.tamanhos.items() if v.get()]
        categorias = [c for c, v in self.categorias.items() if v.get()]
        script = script_renomeio()
        link = self.link.get().strip()
        arquivo = self.planilha.get().strip()

        if not lote or not Path(lote).is_dir():
            messagebox.showwarning("Renomear", "Escolha uma pasta de lote que já exista.")
            return
        if not tamanhos or not categorias:
            messagebox.showwarning("Renomear", "Marque pelo menos um tamanho e uma categoria.")
            return
        if script is None or not script.is_file():
            messagebox.showerror("Renomear", "Não encontrei o PASTA RENOMEIO.py.")
            return
        if link and not RE_LINK.search(link):
            messagebox.showwarning("Renomear", "O link da planilha não parece um Google Sheets.")
            return
        if not link and (not arquivo or not Path(arquivo).is_file()):
            messagebox.showwarning("Renomear", "Informe o link do Google ou escolha a planilha.")
            return
        if not simular and not messagebox.askyesno(
            "Renomear",
            "Os arquivos e pastas do lote serão RENOMEADOS de verdade:\n\n"
            + lote
            + "\n\nRecomendado simular antes. Continuar?",
        ):
            return

        self._salvar_config()
        self.resumo = None
        self.modo = "renomeio"
        self._limpar()
        self._habilitar(False)
        self.status.set("Baixando a planilha..." if link else ("Simulando renomeio..." if simular else "Renomeando..."))
        threading.Thread(
            target=self._preparar_renomeio,
            args=(simular, script, lote, tamanhos, categorias),
            daemon=True,
        ).start()

    def _preparar_renomeio(self, simular, script, lote, tamanhos, categorias):
        try:
            link = self.link.get().strip()
            if link:
                planilha = baixar_planilha_google(link, PLANILHA_BAIXADA)
            else:
                planilha = Path(self.planilha.get().strip())
        except Exception as e:
            self.root.after(0, self._falha_inicio, str(e))
            return
        self.root.after(0, self._comecar_renomeio, simular, script, lote, planilha, tamanhos, categorias)

    def _falha_inicio(self, mensagem):
        self.processo = None
        self._habilitar(True)
        self.status.set("Não foi possível iniciar.")
        self._escrever(mensagem + "\n")

    def _comecar_renomeio(self, simular, script, lote, planilha, tamanhos, categorias):
        args = [
            "--lote", lote,
            "--planilha", str(planilha),
            "--col-nome", self.col_nome.get(),
            "--col-codigo", self.col_codigo.get(),
            "--tamanhos", ",".join(tamanhos),
            "--categorias", ",".join(categorias),
            "--sem-pausa",
            "--simular" if simular else "--real",
        ]
        if self.aba.get().strip():
            args += ["--aba", self.aba.get().strip()]
        self.inicio_operacao = time.time()
        self.arquivo_erros = caminho_planilha_erros("renomeio_simulado" if simular else "renomeio")
        args += ["--erros", str(self.arquivo_erros)]

        self._escrever("> " + script.name + " " + " ".join(f'"{a}"' if " " in a else a for a in args) + "\n")
        self.status.set("Simulando renomeio..." if simular else "Renomeando...")
        ambiente = dict(os.environ)
        ambiente["PYTHONIOENCODING"] = "utf-8"
        ambiente["PYTHONUNBUFFERED"] = "1"
        try:
            self.processo = subprocess.Popen(
                [python_console(), "-u", str(script)] + args,
                cwd=str(script.parent),
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                stdin=subprocess.DEVNULL,
                env=ambiente,
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
            )
        except Exception as e:
            self._falha_inicio(str(e))
            return
        threading.Thread(target=self._ler_saida, daemon=True).start()

    def _envio_argumentos(self):
        argumentos = []
        sge = self.envio_sge.get().strip()
        if sge:
            if not re.fullmatch(r"\d+", sge.split(".")[-1]):
                messagebox.showerror("Enviar aprovação", "Digite só o número do SGE.")
                return None
            argumentos += ["--sge", sge]
        limite = self.envio_limite.get().strip()
        if limite:
            if not limite.isdigit() or int(limite) < 1:
                messagebox.showerror("Enviar aprovação", "O limite é a quantidade de pessoas.")
                return None
            argumentos += ["--limite", limite]
        if not self.envio_exigir_coluna_m.get():
            argumentos += ["--ignorar-coluna-m"]
        if not self.envio_atualizar_planilha.get():
            argumentos += ["--sem-planilha"]
        if not self.envio_baixa_sge.get():
            argumentos += ["--sem-sge"]
        return argumentos

    def _envio_iniciar(self):
        if self.processo:
            return
        argumentos = self._envio_argumentos()
        if argumentos is None:
            return
        simular = self.envio_simular.get()
        if simular:
            argumentos = argumentos + ["--simular"]
        elif not messagebox.askyesno(
            "Enviar de verdade",
            "Isso sobe as fotos no Drive, grava o link em Álbum pronto, envia o e-mail, "
            "atualiza a Coluna K para 'Aguardando Aprovação' e dá baixa no SGE.\n\nContinuar?",
        ):
            return
        self._envio_executar(
            argumentos,
            "Simulando o envio..." if simular else "Enviando álbum para aprovação...",
        )

    def _envio_vigia(self):
        if self.processo:
            return
        argumentos = self._envio_argumentos()
        if argumentos is None:
            return
        intervalo = self.envio_intervalo.get().strip() or "60"
        if not intervalo.isdigit() or int(intervalo) < 10:
            messagebox.showerror("Vigia Contínuo", "O intervalo do vigia deve ser de pelo menos 10 segundos.")
            return
        argumentos += ["--vigia", "--intervalo", intervalo]
        simular = self.envio_simular.get()
        if simular:
            argumentos += ["--simular"]
        self._envio_executar(
            argumentos,
            f"Vigia Contínuo ativo (a cada {intervalo}s) — monitorando DIAGRAMADO + CAPA...",
        )

    def _envio_continuar(self):
        if self.processo:
            return
        argumentos = self._envio_argumentos()
        if argumentos is None:
            return
        argumentos = argumentos + ["--continuar"]
        simular = self.envio_simular.get()
        if simular:
            argumentos = argumentos + ["--simular"]
        elif not messagebox.askyesno(
            "Continuar o envio",
            "Continua de quem ainda não foi enviado.\nSobe fotos, grava o link, manda o e-mail e atualiza Coluna K/SGE.\n\nContinuar?",
        ):
            return
        self._envio_executar(
            argumentos,
            "Simulando o envio..." if simular else "Continuando o envio...",
        )

    def _envio_listar(self):
        if self.processo:
            return
        self._envio_executar(["--listar"], "Listando quem está DIAGRAMADO...")

    def _envio_executar(self, argumentos, status):
        if not SCRIPT_ENVIO.is_file():
            messagebox.showerror("Enviar aprovação", "Script não encontrado:\n\n" + str(SCRIPT_ENVIO))
            return
        self._salvar_config()
        PARAR_ENVIO.unlink(missing_ok=True)
        self.pediu_parar_envio = False
        self.modo = "envio"
        self.inicio_operacao = time.time()
        self.resumo = None
        self._limpar()
        self._escrever("> " + SCRIPT_ENVIO.name + " " + " ".join(argumentos) + "\n")
        self._habilitar(False)
        self.status.set(status)
        ambiente = dict(os.environ)
        ambiente["PYTHONIOENCODING"] = "utf-8"
        ambiente["PYTHONUNBUFFERED"] = "1"
        try:
            self.processo = subprocess.Popen(
                [python_console(), "-u", str(SCRIPT_ENVIO)] + argumentos,
                cwd=str(PASTA),
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                stdin=subprocess.DEVNULL,
                env=ambiente,
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
            )
        except Exception as e:
            self.processo = None
            self.modo = None
            self._habilitar(True)
            self.status.set("Não foi possível iniciar o envio.")
            self._escrever(str(e) + "\n")
            return
        threading.Thread(target=self._ler_saida, daemon=True).start()


def main():
    root = tk.Tk()
    JanelaOrganizador(root)
    root.mainloop()


if __name__ == "__main__":
    main()
