# -*- coding: utf-8 -*-
"""
Janela do robô de baixa no SGE (BAIXA.PY).

- Planilha por arquivo .xlsx ou por link do Google, com aba e colunas
- Lê os itens e status da planilha e marca itens como prioritários
- Contas do SGE (as mesmas do robô do CPF), com troca automática de conta
- Status, data e observação da alteração; opção de reenviar no mesmo status
- Inicia, continua de onde parou e para o robô (a linha fica salva em baixa_progresso.json)
"""

import ctypes
import json
import os
import queue
import re
import subprocess
import sys
import threading
import tkinter as tk
import unicodedata
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox, scrolledtext, ttk

PASTA = Path(__file__).resolve().parent
ROBO = PASTA / "BAIXA.PY"
CONFIG = PASTA / "sge_config.json"
CRED_SGE = PASTA / "sge_credenciais.json"
PROGRESSO = PASTA / "baixa_progresso.json"
PARAR = PASTA / "baixa_parar.flag"
RELATORIO = PASTA / "baixa_sge.csv"

STATUS_DA_PLANILHA = "(usar o da planilha)"
SEM_COLUNA = "(nenhuma)"

# Tema Visual Moderno (Estilo Preview Monitor Contínuo — Sem Tema Hacker)
COR_FUNDO = "#F8FAFC"          # Slate-50 (Fundo limpo da Preview)
COR_CAIXA = "#FFFFFF"          # Branco puro (Campos e caixas de listagem)
COR_TEXTO = "#0F172A"          # Slate-900 (Texto principal escuro e nítido)
COR_FORTE = "#2563EB"          # Blue-600 (Destaque / Ação Primária da Preview)
COR_FRACO = "#64748B"          # Slate-500 (Texto secundário / legendas)
COR_BORDA = "#CBD5E1"          # Slate-300 (Bordas suaves)
COR_BOTAO = "#FFFFFF"          # Branco puro (Superfície de botões)
COR_BOTAO_ATIVO = "#EFF6FF"    # Blue-50 (Hover / Ativo)
COR_SELECAO = "#DBEAFE"        # Blue-100 (Seleção)
COR_VERDE = "#059669"          # Emerald-600 (Sucesso / Aprovado / Iniciar)
COR_VERMELHO = "#DC2626"       # Red-600 (Erro / Parar)
COR_CINZA = "#475569"          # Slate-600 (Ações secundárias)
COR_AVISO = "#D97706"          # Amber-600 (Avisos / Em Impressão)
FONTE = ("Segoe UI", 9)
FONTE_NEGRITO = ("Segoe UI", 9, "bold")


def aplicar_tema_terminal(root):
    """Tema Moderno da Preview (Slate Claro, Azul Royal e Esmeralda — Fim do Tema Hacker)."""

    root.configure(background=COR_FUNDO)
    root.option_add("*Font", "SegoeUI 9")
    root.option_add("*TCombobox*Listbox.background", COR_CAIXA)
    root.option_add("*TCombobox*Listbox.foreground", COR_TEXTO)
    root.option_add("*TCombobox*Listbox.selectBackground", COR_SELECAO)
    root.option_add("*TCombobox*Listbox.selectForeground", COR_FORTE)
    root.option_add("*TCombobox*Listbox.font", "SegoeUI 9")

    estilo = ttk.Style(root)
    try:
        estilo.theme_use("clam")
    except tk.TclError:
        pass

    estilo.configure(
        ".",
        background=COR_FUNDO,
        foreground=COR_TEXTO,
        font=FONTE,
        troughcolor="#E2E8F0",
        bordercolor=COR_BORDA,
        lightcolor=COR_CAIXA,
        darkcolor=COR_BORDA,
    )
    estilo.configure("TFrame", background=COR_FUNDO)
    estilo.configure("TLabel", background=COR_FUNDO, foreground=COR_TEXTO, font=FONTE)
    estilo.configure(
        "TCombobox",
        fieldbackground=COR_CAIXA,
        background=COR_BOTAO,
        foreground=COR_TEXTO,
        arrowcolor=COR_TEXTO,
        bordercolor=COR_BORDA,
        insertcolor=COR_TEXTO,
        arrowsize=12,
    )
    estilo.map(
        "TCombobox",
        fieldbackground=[("readonly", COR_CAIXA), ("disabled", COR_FUNDO)],
        foreground=[("readonly", COR_TEXTO), ("disabled", COR_FRACO)],
        selectbackground=[("readonly", COR_SELECAO)],
        selectforeground=[("readonly", COR_FORTE)],
        background=[("active", COR_BOTAO_ATIVO)],
        arrowcolor=[("disabled", COR_FRACO)],
    )
    estilo.configure(
        "Vertical.TScrollbar",
        background=COR_BOTAO,
        troughcolor="#E2E8F0",
        arrowcolor=COR_TEXTO,
        bordercolor=COR_BORDA,
    )
    estilo.map("Vertical.TScrollbar", background=[("active", COR_BOTAO_ATIVO)])


def barra_escura(janela):
    try:
        janela.update_idletasks()
        hwnd = ctypes.windll.user32.GetParent(janela.winfo_id())
        valor = ctypes.c_int(1)
        ctypes.windll.dwmapi.DwmSetWindowAttribute(
            hwnd, 20, ctypes.byref(valor), ctypes.sizeof(valor)
        )
    except Exception:
        return


def ler_json(arquivo):
    try:
        return json.loads(arquivo.read_text(encoding="utf-8"))
    except Exception:
        return {}


def gravar_json(arquivo, dados):
    arquivo.write_text(json.dumps(dados, indent=2, ensure_ascii=False), encoding="utf-8")


def normalizar(texto):
    texto = unicodedata.normalize("NFKD", str(texto or ""))
    texto = "".join(c for c in texto if not unicodedata.combining(c))
    return re.sub(r"\s+", " ", texto).strip().upper()


def python_console():
    """Com pythonw a saída se perde; o robô roda com o python.exe da mesma pasta."""

    exe = Path(sys.executable)
    candidato = exe.with_name("python.exe")
    return str(candidato if candidato.exists() else exe)


class Janela:

    def __init__(self, root):
        self.root = root
        self.processo = None
        self.fila = queue.Queue()
        self.pediu_parar = False
        self.contas_sge = []
        self.itens_lidos = []
        self.itens_dados = []

        root.title("Robô SGE — baixa de status")
        root.geometry("1120x900")
        root.minsize(980, 700)
        root.protocol("WM_DELETE_WINDOW", self._fechar)
        aplicar_tema_terminal(root)

        self._montar()
        self._carregar()
        self._atualizar_progresso()
        barra_escura(root)
        root.after(150, self._ler_saida)

    # ------------------------------------------------------------------ montagem
    def _secao(self, titulo):
        quadro = tk.LabelFrame(
            self.root, text=f" {titulo} ", bg=COR_FUNDO, fg=COR_TEXTO,
            font=("Consolas", 10, "bold"), padx=8, pady=4, bd=1, relief="solid",
            highlightthickness=1, highlightbackground=COR_BORDA, highlightcolor=COR_TEXTO,
        )
        quadro.pack(fill="x", padx=10, pady=(6, 0))
        return quadro

    def _rotulo(self, pai, texto, linha, coluna, **grid):
        tk.Label(pai, text=texto, bg=COR_FUNDO, fg=COR_TEXTO, font=FONTE).grid(
            row=linha, column=coluna, sticky="w", padx=(0, 6), pady=2, **grid)

    def _campo(self, pai, linha, coluna, largura=30, senha=False, colspan=1):
        campo = tk.Entry(
            pai, width=largura, bg=COR_CAIXA, fg=COR_TEXTO, insertbackground=COR_TEXTO,
            relief="flat", font=FONTE, show="•" if senha else "",
            highlightthickness=1, highlightbackground=COR_BORDA, highlightcolor=COR_TEXTO,
            selectbackground=COR_SELECAO, selectforeground=COR_FORTE,
        )
        campo.grid(row=linha, column=coluna, columnspan=colspan, sticky="we", pady=2, padx=(0, 12))
        return campo

    def _botao(self, pai, texto, comando, cor=COR_TEXTO):
        return tk.Button(
            pai, text=texto, command=comando, bg=COR_BOTAO, fg=cor,
            activebackground=COR_BOTAO_ATIVO, activeforeground=cor,
            disabledforeground=COR_FRACO, relief="flat", bd=0,
            highlightthickness=1, highlightbackground=cor, highlightcolor=cor,
            font=FONTE_NEGRITO, padx=10, pady=3, cursor="hand2",
        )

    def _check(self, pai, texto, variavel, comando=None):
        return tk.Checkbutton(
            pai, text=texto, variable=variavel, command=comando, bg=COR_FUNDO, fg=COR_TEXTO,
            selectcolor=COR_BOTAO, activebackground=COR_FUNDO, activeforeground=COR_FORTE,
            font=FONTE, highlightthickness=0,
        )

    def _lista(self, pai, altura, multipla=False):
        return tk.Listbox(
            pai, height=altura, bg=COR_CAIXA, fg=COR_TEXTO, relief="flat",
            selectbackground=COR_SELECAO, selectforeground=COR_FORTE,
            highlightthickness=1, highlightbackground=COR_BORDA, highlightcolor=COR_TEXTO,
            font=FONTE, exportselection=False,
            selectmode="extended" if multipla else "browse",
        )

    def _montar(self):
        self._montar_planilha()
        self._montar_itens()
        self._montar_contas()
        self._montar_execucao()

        self.status = tk.Label(self.root, text="Parado", bg=COR_FUNDO, fg=COR_TEXTO,
                               font=FONTE_NEGRITO, anchor="w")
        self.status.pack(fill="x", padx=12, pady=(6, 0))
        self.log = scrolledtext.ScrolledText(
            self.root, bg=COR_FUNDO, fg=COR_TEXTO, insertbackground=COR_TEXTO,
            font=("Consolas", 10), relief="flat", wrap="word", height=10,
            selectbackground=COR_SELECAO, selectforeground=COR_FORTE,
            highlightthickness=1, highlightbackground=COR_BORDA, highlightcolor=COR_TEXTO,
        )
        self.log.pack(fill="both", expand=True, padx=10, pady=(4, 10))
        self.log.frame.configure(background=COR_FUNDO)
        self.log.vbar.pack_forget()
        barra = ttk.Scrollbar(self.log.frame, command=self.log.yview)
        barra.pack(side="right", fill="y")
        self.log.configure(yscrollcommand=barra.set)
        self.log.vbar = barra
        self.log.tag_configure("ok", foreground=COR_FORTE)
        self.log.tag_configure("erro", foreground=COR_VERMELHO)
        self.log.tag_configure("aviso", foreground=COR_AVISO)
        self.log.tag_configure("aluno", foreground=COR_TEXTO)
        self._botoes_rodando(False)

    def _montar_planilha(self):
        sec = self._secao("Planilha")
        sec.columnconfigure(2, weight=1)
        self.origem = tk.StringVar(value="arquivo")
        tk.Radiobutton(
            sec, text="Arquivo:", variable=self.origem, value="arquivo", bg=COR_FUNDO, fg=COR_TEXTO,
            selectcolor=COR_BOTAO, activebackground=COR_FUNDO, activeforeground=COR_FORTE,
            font=FONTE, highlightthickness=0,
        ).grid(row=0, column=0, columnspan=2, sticky="w")
        self.arquivo = self._campo(sec, 0, 2, largura=60, colspan=6)
        self._botao(sec, "Procurar...", self._procurar, COR_CINZA).grid(row=0, column=8, sticky="w")
        tk.Radiobutton(
            sec, text="Link do Google:", variable=self.origem, value="link", bg=COR_FUNDO, fg=COR_TEXTO,
            selectcolor=COR_BOTAO, activebackground=COR_FUNDO, activeforeground=COR_FORTE,
            font=FONTE, highlightthickness=0,
        ).grid(row=1, column=0, columnspan=2, sticky="w")
        self.link = self._campo(sec, 1, 2, largura=60, colspan=6)
        self._botao(sec, "Identificar colunas", self._detectar, COR_VERDE).grid(row=1, column=8, sticky="w")
        self.link.bind("<FocusOut>", lambda _e: self._link_mudou())
        self.link.bind("<Return>", lambda _e: self._link_mudou())

        linha = tk.Frame(sec, bg=COR_FUNDO)
        linha.grid(row=2, column=0, columnspan=9, sticky="w", pady=(2, 0))
        self.colunas = {}
        for texto, chave, largura in (("Aba:", "aba", 22), ("1ª linha:", "linha_inicial", 4),
                                      ("SGE:", "col_sge", 16), ("Nome:", "col_nome", 16),
                                      ("Status:", "col_status", 16), ("Item:", "col_item", 16)):
            tk.Label(linha, text=texto, bg=COR_FUNDO, fg=COR_TEXTO, font=FONTE).pack(side="left", padx=(0, 4))
            if chave == "linha_inicial":
                campo = tk.Entry(
                    linha, width=largura, bg=COR_CAIXA, fg=COR_TEXTO, insertbackground=COR_TEXTO,
                    relief="flat", font=FONTE,
                    highlightthickness=1, highlightbackground=COR_BORDA, highlightcolor=COR_TEXTO,
                    selectbackground=COR_SELECAO, selectforeground=COR_FORTE,
                )
            else:
                campo = ttk.Combobox(linha, width=largura, font=FONTE)
            campo.pack(side="left", padx=(0, 10))
            self.colunas[chave] = campo
        self.colunas["aba"].bind("<<ComboboxSelected>>", lambda _e: self._aba_escolhida())
        botoes = tk.Frame(sec, bg=COR_FUNDO)
        botoes.grid(row=3, column=0, columnspan=9, sticky="w", pady=(2, 0))
        self._botao(botoes, "Salvar", self._salvar_planilha).pack(side="left", padx=(0, 3))
        self._botao(botoes, "Testar planilha", self._testar_planilha, COR_CINZA).pack(side="left", padx=3)
        self.aviso_colunas = tk.Label(botoes, text="", bg=COR_FUNDO, fg=COR_FRACO, font=FONTE)
        self.aviso_colunas.pack(side="left", padx=8)
        self.deteccao = {}
        self.link_detectado = ""

    def _montar_itens(self):
        sec = self._secao("Itens da planilha (aprender)")
        sec.columnconfigure(0, weight=3)
        sec.columnconfigure(1, weight=2)
        tk.Label(sec, text="Itens (★ = prioritário; selecione vários com Ctrl)", bg=COR_FUNDO, fg=COR_TEXTO,
                 font=FONTE).grid(row=0, column=0, sticky="w")
        tk.Label(sec, text="Status da planilha -> no SGE", bg=COR_FUNDO, fg=COR_TEXTO,
                 font=FONTE).grid(row=0, column=1, sticky="w", padx=(10, 0))
        self.lista_itens = self._lista(sec, 5, multipla=True)
        self.lista_itens.grid(row=1, column=0, sticky="we")
        self.lista_status = self._lista(sec, 5)
        self.lista_status.grid(row=1, column=1, sticky="we", padx=(10, 0))
        botoes = tk.Frame(sec, bg=COR_FUNDO)
        botoes.grid(row=2, column=0, columnspan=2, sticky="w", pady=(4, 0))
        self._botao(botoes, "Ler itens da planilha", self._ler_itens).pack(side="left", padx=(0, 3))
        self._botao(botoes, "★ Usar como prioritários", self._marcar_prioritarios, COR_VERDE).pack(side="left", padx=3)
        self._botao(botoes, "Tirar dos prioritários", self._tirar_prioritarios, COR_CINZA).pack(side="left", padx=3)
        tk.Label(
            sec,
            text="Prioritário = o robô escolhe esse item antes do plano (ex.: álbum vendido separado).",
            bg=COR_FUNDO, fg=COR_FRACO, font=FONTE, wraplength=1040, justify="left",
        ).grid(row=3, column=0, columnspan=2, sticky="w", pady=(4, 0))
        manual = tk.Frame(sec, bg=COR_FUNDO)
        manual.grid(row=4, column=0, columnspan=2, sticky="w", pady=(4, 0))
        tk.Label(manual, text="Item que não está na planilha:", bg=COR_FUNDO, fg=COR_TEXTO,
                 font=FONTE).pack(side="left", padx=(0, 4))
        self.item_manual = tk.Entry(
            manual, width=40, bg=COR_CAIXA, fg=COR_TEXTO, insertbackground=COR_TEXTO,
            relief="flat", font=FONTE,
            highlightthickness=1, highlightbackground=COR_BORDA, highlightcolor=COR_TEXTO,
            selectbackground=COR_SELECAO, selectforeground=COR_FORTE,
        )
        self.item_manual.pack(side="left", padx=(0, 6))
        self._botao(manual, "★ Prioritário desta planilha (link)", self._adicionar_item_planilha,
                    COR_VERDE).pack(side="left")

    def _montar_contas(self):
        sec = self._secao("Contas do SGE (mesmas do robô do CPF)")
        sec.columnconfigure(1, weight=1)
        self.lista_sge = self._lista(sec, 3)
        self.lista_sge.grid(row=0, column=0, columnspan=5, sticky="we", pady=2)
        self.lista_sge.bind("<<ListboxSelect>>", self._conta_selecionada)
        self._rotulo(sec, "Usuário:", 1, 0)
        self.sge_usuario = self._campo(sec, 1, 1, largura=30)
        self._rotulo(sec, "Senha:", 1, 2)
        self.sge_senha = self._campo(sec, 1, 3, largura=18, senha=True)
        self.mostrar = tk.BooleanVar(value=False)
        self._check(sec, "mostrar senhas", self.mostrar, self._mostrar_senhas).grid(row=1, column=4, sticky="w")
        botoes = tk.Frame(sec, bg=COR_FUNDO)
        botoes.grid(row=2, column=0, columnspan=5, sticky="w")
        self._botao(botoes, "Adicionar conta", self._adicionar_conta).pack(side="left", padx=(0, 3))
        self._botao(botoes, "Alterar selecionada", self._alterar_conta, COR_CINZA).pack(side="left", padx=3)
        self._botao(botoes, "Remover", self._remover_conta, COR_VERMELHO).pack(side="left", padx=3)
        self._botao(botoes, "▲ Subir", self._subir_conta, COR_CINZA).pack(side="left", padx=3)
        tk.Label(
            sec,
            text="Entra sempre com 'Sim'. Se a sessão cair 2 vezes, alguém está usando: troca de conta.",
            bg=COR_FUNDO, fg=COR_FRACO, font=FONTE, wraplength=1040, justify="left",
        ).grid(row=3, column=0, columnspan=5, sticky="w", pady=(4, 0))

    def _montar_execucao(self):
        sec = self._secao("Execução")
        sec.columnconfigure(5, weight=1)
        self.progresso = tk.Label(sec, text="", bg=COR_FUNDO, fg=COR_FORTE, font=FONTE_NEGRITO,
                                  anchor="w", justify="left")
        self.progresso.grid(row=0, column=0, columnspan=8, sticky="we", pady=(0, 4))

        self._rotulo(sec, "Status:", 1, 0)
        self.status_escolhido = ttk.Combobox(sec, width=28, font=FONTE)
        self.status_escolhido.grid(row=1, column=1, sticky="w", padx=(0, 12))
        self._rotulo(sec, "Data:", 1, 2)
        self.data = self._campo(sec, 1, 3, largura=12)
        self._rotulo(sec, "Observação:", 1, 4)
        self.obs = self._campo(sec, 1, 5, largura=40, colspan=3)

        self._rotulo(sec, "Começar na linha:", 2, 0)
        self.linha = self._campo(sec, 2, 1, largura=8)
        self._rotulo(sec, "Limite:", 2, 2)
        self.limite = self._campo(sec, 2, 3, largura=6)
        self.simular = tk.BooleanVar(value=True)
        self.refazer = tk.BooleanVar(value=False)
        self.mesmo_status = tk.BooleanVar(value=False)
        opcoes = tk.Frame(sec, bg=COR_FUNDO)
        opcoes.grid(row=2, column=4, columnspan=4, sticky="w")
        self._check(opcoes, "Simular (clica em Cancelar)", self.simular).pack(side="left", padx=(0, 10))
        self._check(opcoes, "Refazer os já feitos", self.refazer).pack(side="left", padx=(0, 10))
        self._check(sec, "Atualizar mesmo se já estiver no status (reenviar com nova data/observação)",
                    self.mesmo_status, self._avisar_reenvio).grid(row=3, column=0, columnspan=8, sticky="w")

        botoes = tk.Frame(sec, bg=COR_FUNDO)
        botoes.grid(row=4, column=0, columnspan=8, sticky="w", pady=(4, 0))
        self.b_iniciar = self._botao(botoes, "▶  Iniciar", self._iniciar, COR_VERDE)
        self.b_vigia = self._botao(botoes, "♾  Vigia Contínuo (Sem Fim)", self._iniciar_vigia, COR_FORTE)
        self.b_continuar = self._botao(botoes, "⏩  Continuar de onde parou", self._continuar, COR_VERDE)
        self.b_parar = self._botao(botoes, "■  Parar", self._parar, COR_VERMELHO)
        self.b_relatorio = self._botao(botoes, "Abrir relatório", self._abrir_relatorio, COR_CINZA)
        for botao in (self.b_iniciar, self.b_vigia, self.b_continuar, self.b_parar, self.b_relatorio):
            botao.pack(side="left", padx=3)

    # ------------------------------------------------------------------ dados
    def _carregar(self):
        cfg = ler_json(CONFIG)
        self.arquivo.insert(0, cfg.get("planilha", ""))
        self.link.insert(0, cfg.get("planilha_link") or cfg.get("ui_link", ""))
        self.origem.set("link" if cfg.get("planilha_link") else "arquivo")
        self.link_detectado = self.link.get().strip()
        for chave, campo in self.colunas.items():
            campo.insert(0, str(cfg.get(chave, "") or ""))
        self.obs.insert(0, cfg.get("observacao", ""))
        self.data.insert(0, datetime.now().strftime("%d/%m/%Y"))
        self._carregar_status(cfg)
        self.status_escolhido.set(cfg.get("ui_status") or STATUS_DA_PLANILHA)

        sge = ler_json(CRED_SGE)
        self.contas_sge = [{"usuario": c.get("usuario", ""), "senha": c.get("senha", "")}
                           for c in sge.get("contas", []) if c.get("usuario")]
        if not self.contas_sge and sge.get("usuario"):
            self.contas_sge = [{"usuario": sge["usuario"], "senha": sge.get("senha", "")}]
        self._mostrar_contas()
        self._preencher_itens()

    def _carregar_status(self, cfg):
        valores = [STATUS_DA_PLANILHA]
        for status in list(cfg.get("status_sge_conhecidos", [])) + list(cfg.get("mapa_status", {}).values()):
            if status and status not in valores:
                valores.append(status)
        self.status_escolhido.config(values=valores)

    def _procurar(self):
        caminho = filedialog.askopenfilename(title="Escolha a planilha", initialdir=str(PASTA),
                                             filetypes=[("Planilhas Excel", "*.xlsx *.xlsm"), ("Todos", "*.*")])
        if caminho:
            self.arquivo.delete(0, "end")
            self.arquivo.insert(0, caminho)
            self.origem.set("arquivo")
            self._detectar()

    # ------------------------------------------------------------------ identificar colunas (linha dos filtros)
    def _link_mudou(self):
        link = self.link.get().strip()
        if link and link != self.link_detectado and re.search(r"/spreadsheets/d/[\w-]+", link):
            self.origem.set("link")
            self._detectar()

    def _detectar(self):
        if self.processo:
            return
        if self.origem.get() == "link":
            link = self.link.get().strip()
            if not re.search(r"/spreadsheets/d/[\w-]+", link):
                messagebox.showerror("Link inválido", "Cole o link da planilha do Google (docs.google.com/spreadsheets/d/...).")
                return
            self.link_detectado = link
            argumentos = ["--detectar", "--link", link]
        else:
            arquivo = self.arquivo.get().strip()
            if not arquivo or not Path(arquivo).exists():
                messagebox.showerror("Planilha", "Escolha um arquivo .xlsx que exista (botão Procurar).")
                return
            argumentos = ["--detectar", "--planilha", arquivo]
        self.aviso_colunas.config(text="Identificando as colunas...")
        self._rodar(argumentos, simular=False, rodada=False)

    def _mostrar_deteccao(self, dados):
        self.deteccao = {a["aba"]: a for a in dados.get("abas", [])}
        self._preencher_itens()
        aba = self.colunas["aba"]
        aba.config(values=list(self.deteccao))
        atual = self.deteccao.get(aba.get().strip())
        sugerida = aba.get().strip() if atual and atual["pontos"] else dados.get("sugerida") or ""
        if not sugerida:
            self.aviso_colunas.config(text="Não achei uma linha de cabeçalho com SGE/código — escolha a aba e as colunas.")
            return
        aba.set(sugerida)
        self._aba_escolhida()

    def _aba_escolhida(self):
        info = self.deteccao.get(self.colunas["aba"].get())
        if not info:
            return
        opcoes = [f"{letra} — {texto}" for letra, texto in info["cabecalhos"]]
        rotulo = {letra: f"{letra} — {texto}" for letra, texto in info["cabecalhos"]}
        for chave in ("col_sge", "col_nome", "col_status", "col_item"):
            campo = self.colunas[chave]
            campo.config(values=opcoes + ([SEM_COLUNA] if chave in ("col_status", "col_item") else []))
            letra = info["sugestao"].get(chave[4:])
            campo.set(rotulo.get(letra, letra) if letra else (SEM_COLUNA if chave in ("col_status", "col_item") else ""))
        self.colunas["linha_inicial"].delete(0, "end")
        if info["linha"]:
            self.colunas["linha_inicial"].insert(0, str(info["linha"] + 1))
            filtro = f" (filtro em {info['filtro']})" if info.get("filtro") else ""
            faltam = [n for c, n in (("nome", "Nome"), ("status", "Status"), ("item", "Item")) if c not in info["sugestao"]]
            self.aviso_colunas.config(
                text=f"Cabeçalho na linha {info['linha']}{filtro}. Confira as colunas"
                     + (f" — não achei: {', '.join(faltam)}" if faltam else "") + ".")
        else:
            self.aviso_colunas.config(text="Nessa aba não achei cabeçalho com SGE/código — escolha as colunas.")

    def _letra(self, chave):
        valor = self.colunas[chave].get().strip()
        if not valor or valor == SEM_COLUNA:
            return ""
        return valor.split("—")[0].strip().upper()

    def _salvar_planilha(self):
        cfg = ler_json(CONFIG)
        antes = (cfg.get("planilha"), cfg.get("planilha_link"), cfg.get("aba"))
        link = self.link.get().strip()
        if self.origem.get() == "link":
            if not re.search(r"/spreadsheets/d/[\w-]+", link):
                messagebox.showerror("Link inválido", "Cole o link da planilha do Google (docs.google.com/spreadsheets/d/...).")
                return False
            cfg["planilha_link"] = link
        else:
            arquivo = self.arquivo.get().strip()
            if not arquivo or not Path(arquivo).exists():
                messagebox.showerror("Planilha", "Escolha um arquivo .xlsx que exista (botão Procurar).")
                return False
            cfg["planilha"] = arquivo
            cfg["planilha_link"] = ""
        cfg["ui_link"] = link
        for chave, campo in self.colunas.items():
            valor = campo.get().strip()
            if chave == "linha_inicial":
                if valor and not valor.isdigit():
                    messagebox.showerror("Planilha", "A 1ª linha precisa ser um número.")
                    return False
                cfg[chave] = int(valor or 2)
            elif chave == "aba":
                cfg[chave] = valor
            else:
                cfg[chave] = self._letra(chave)
                if cfg[chave] and not re.fullmatch(r"[A-Z]{1,3}|\d+", cfg[chave]):
                    messagebox.showerror("Planilha", f"Coluna inválida: '{valor}' (use a letra, ex.: E).")
                    return False
        if not cfg.get("col_sge") or (not cfg.get("col_status") and self._status_forcado() is None):
            messagebox.showerror("Planilha", "Preencha pelo menos a coluna do SGE (e a do Status, ou escolha um status).")
            return False
        gravar_json(CONFIG, cfg)
        if antes != (cfg.get("planilha"), cfg.get("planilha_link"), cfg.get("aba")):
            PROGRESSO.unlink(missing_ok=True)
            self._atualizar_progresso()
        return True

    def _status_forcado(self):
        status = self.status_escolhido.get().strip()
        return None if not status or status == STATUS_DA_PLANILHA else status

    # ------------------------------------------------------------------ itens
    def _ler_itens(self):
        if self._salvar_planilha():
            self._rodar(["--listar-itens"], simular=False, rodada=False)

    def _id_planilha(self):
        if self.origem.get() != "link":
            return ""
        achado = re.search(r"/spreadsheets/d/([\w-]+)", self.link.get().strip())
        return achado.group(1) if achado else ""

    def _preencher_itens(self):
        """Primeiro os prioritários fixos da planilha do link (não aparecem nas colunas), depois os itens lidos."""

        cfg = ler_json(CONFIG)
        prioritarios = {normalizar(p) for p in cfg.get("itens_prioritarios", [])}
        da_planilha = cfg.get("prioritarios_por_planilha", {}).get(self._id_planilha(), [])
        self.lista_itens.delete(0, "end")
        self.itens_lidos = []
        for nome in da_planilha:
            self.itens_lidos.append(("planilha", nome))
            self.lista_itens.insert("end", f"★ {nome}  [desta planilha]")
        for nome, quantos in self.itens_dados:
            self.itens_lidos.append(("geral", nome))
            marca = "★ " if normalizar(nome) in prioritarios else "   "
            self.lista_itens.insert("end", f"{marca}{nome}  ({quantos})")

    def _adicionar_item_planilha(self):
        nome = self.item_manual.get().strip()
        id_planilha = self._id_planilha()
        if not id_planilha:
            messagebox.showerror("Item desta planilha", "Escolha 'Link do Google' com o link da planilha primeiro.")
            return
        if not nome:
            messagebox.showerror("Item desta planilha", "Escreva o nome do item como aparece no SGE.")
            return
        cfg = ler_json(CONFIG)
        por_planilha = cfg.setdefault("prioritarios_por_planilha", {})
        lista = por_planilha.setdefault(id_planilha, [])
        if normalizar(nome) not in {normalizar(x) for x in lista}:
            lista.append(nome)
        gravar_json(CONFIG, cfg)
        self.item_manual.delete(0, "end")
        self._escrever(f"Prioritário desta planilha: {nome}\n", "ok")
        self._preencher_itens()

    def _mostrar_itens(self, dados):
        cfg = ler_json(CONFIG)
        self.itens_dados = dados.get("itens", [])
        self._preencher_itens()
        mapa = cfg.get("mapa_status", {})
        self.lista_status.delete(0, "end")
        for nome, quantos in dados.get("status", []):
            no_sge = mapa.get(normalizar(nome), nome)
            self.lista_status.insert("end", f"{nome} ({quantos})" + (f"  ->  {no_sge}" if no_sge != nome else ""))
        self._escrever(f"Itens lidos: {len(self.itens_dados)} itens diferentes em {dados.get('linhas', 0)} linhas.\n", "ok")

    def _itens_selecionados(self, tipo=None):
        escolhidos = [self.itens_lidos[i] for i in self.lista_itens.curselection() if i < len(self.itens_lidos)]
        if not escolhidos:
            messagebox.showinfo("Itens", "Clique em 'Ler itens da planilha' e selecione um ou mais itens.")
        if tipo:
            return [nome for t, nome in escolhidos if t == tipo]
        return escolhidos

    def _marcar_prioritarios(self):
        escolhidos = self._itens_selecionados("geral")
        if not escolhidos:
            return
        cfg = ler_json(CONFIG)
        for chave in ("itens_prioritarios", "itens_incluir"):
            lista = list(cfg.get(chave, []))
            for nome in escolhidos:
                if normalizar(nome) not in {normalizar(x) for x in lista}:
                    lista.append(normalizar(nome))
            cfg[chave] = lista
        gravar_json(CONFIG, cfg)
        self._escrever("Prioritários: " + ", ".join(escolhidos) + "\n", "ok")
        self._remarcar_itens()

    def _tirar_prioritarios(self):
        selecionados = self._itens_selecionados()
        escolhidos = {normalizar(n) for _, n in selecionados}
        if not escolhidos:
            return
        cfg = ler_json(CONFIG)
        for chave in ("itens_prioritarios", "itens_incluir"):
            cfg[chave] = [x for x in cfg.get(chave, []) if normalizar(x) not in escolhidos]
        id_planilha = self._id_planilha()
        da_planilha = {normalizar(n) for t, n in selecionados if t == "planilha"}
        if id_planilha and da_planilha:
            por_planilha = cfg.setdefault("prioritarios_por_planilha", {})
            por_planilha[id_planilha] = [x for x in por_planilha.get(id_planilha, [])
                                         if normalizar(x) not in da_planilha]
        gravar_json(CONFIG, cfg)
        self._escrever("Tirados dos prioritários: " + ", ".join(sorted(escolhidos)) + "\n", "ok")
        self._remarcar_itens()

    def _remarcar_itens(self):
        selecionados = self.lista_itens.curselection()
        self._preencher_itens()
        for i in selecionados:
            self.lista_itens.selection_set(i)

    # ------------------------------------------------------------------ contas do SGE
    def _mostrar_senhas(self):
        self.sge_senha.config(show="" if self.mostrar.get() else "•")
        self._mostrar_contas()

    def _mostrar_contas(self, selecionar=None):
        self.lista_sge.delete(0, "end")
        for i, conta in enumerate(self.contas_sge, 1):
            senha = conta["senha"] if self.mostrar.get() else "•" * len(conta["senha"])
            self.lista_sge.insert("end", f"{i}. {conta['usuario']}   ({senha})")
        if selecionar is not None and 0 <= selecionar < len(self.contas_sge):
            self.lista_sge.selection_set(selecionar)

    def _indice_conta(self):
        sel = self.lista_sge.curselection()
        return sel[0] if sel else None

    def _conta_selecionada(self, _evento=None):
        i = self._indice_conta()
        if i is None:
            return
        self.sge_usuario.delete(0, "end")
        self.sge_usuario.insert(0, self.contas_sge[i]["usuario"])
        self.sge_senha.delete(0, "end")
        self.sge_senha.insert(0, self.contas_sge[i]["senha"])

    def _conta_dos_campos(self):
        usuario, senha = self.sge_usuario.get().strip(), self.sge_senha.get()
        if not usuario or not senha:
            messagebox.showerror("Conta do SGE", "Preencha usuário e senha do SGE.")
            return None
        return {"usuario": usuario, "senha": senha}

    def _adicionar_conta(self):
        conta = self._conta_dos_campos()
        if conta is None:
            return
        if any(c["usuario"].lower() == conta["usuario"].lower() for c in self.contas_sge):
            messagebox.showerror("Conta do SGE", "Essa conta já está na lista (use 'Alterar selecionada').")
            return
        self.contas_sge.append(conta)
        self._mostrar_contas(len(self.contas_sge) - 1)
        self._salvar_contas()

    def _alterar_conta(self):
        i, conta = self._indice_conta(), self._conta_dos_campos()
        if i is None or conta is None:
            return
        self.contas_sge[i] = conta
        self._mostrar_contas(i)
        self._salvar_contas()

    def _remover_conta(self):
        i = self._indice_conta()
        if i is None:
            return
        if not messagebox.askyesno("Remover conta", f"Remover a conta {self.contas_sge[i]['usuario']}?"):
            return
        del self.contas_sge[i]
        self._mostrar_contas()
        self._salvar_contas()

    def _subir_conta(self):
        i = self._indice_conta()
        if not i:
            return
        self.contas_sge[i - 1], self.contas_sge[i] = self.contas_sge[i], self.contas_sge[i - 1]
        self._mostrar_contas(i - 1)
        self._salvar_contas()

    def _salvar_contas(self):
        dados = ler_json(CRED_SGE)
        dados["contas"] = self.contas_sge
        if self.contas_sge:
            dados["usuario"], dados["senha"] = self.contas_sge[0]["usuario"], self.contas_sge[0]["senha"]
        gravar_json(CRED_SGE, dados)
        self._escrever(f"Contas do SGE salvas ({len(self.contas_sge)}).\n", "ok")

    # ------------------------------------------------------------------ progresso
    def _atualizar_progresso(self):
        prog = ler_json(PROGRESSO)
        if prog and not prog.get("concluido") and prog.get("proxima_linha"):
            texto = (f"Parou depois da linha {prog['ultima_linha']} (SGE {prog.get('sge')} — {prog.get('nome')}) "
                     f"em {prog.get('data')}. Próxima: linha {prog['proxima_linha']}.")
            self.b_continuar.config(text=f"⏩  Continuar da linha {prog['proxima_linha']}")
        elif prog.get("concluido"):
            texto = f"Última rodada terminou a planilha toda em {prog.get('data')}."
            self.b_continuar.config(text="⏩  Continuar de onde parou")
        else:
            texto = "Nenhuma rodada interrompida."
            self.b_continuar.config(text="⏩  Continuar de onde parou")
        self.progresso.config(text=texto)

    # ------------------------------------------------------------------ ações
    def _avisar_reenvio(self):
        if self.mesmo_status.get():
            messagebox.showinfo(
                "Reenviar no mesmo status",
                "Itens que já estão no status escolhido serão alterados de novo (mesmo status),\n"
                "com a Data e a Observação preenchidas aqui.\n\n"
                "Entregue/Cancelado e itens em etapa adiante continuam ignorados.")

    def _opcoes_rodada(self):
        argumentos = ["--sem-pausa"]
        limite = self.limite.get().strip()
        if limite:
            if not limite.isdigit():
                messagebox.showerror("Limite inválido", "Digite só o número de alunos.")
                return None
            argumentos += ["--limite", limite]
        if self.refazer.get():
            argumentos.append("--refazer")

        data = self.data.get().strip()
        if data:
            try:
                datetime.strptime(data, "%d/%m/%Y")
            except ValueError:
                messagebox.showerror("Data inválida", "Use o formato dd/mm/aaaa (ex.: 02/10/2026).")
                return None
        obs = self.obs.get().strip()
        argumentos += ["--data", data, "--obs", obs]
        status = self._status_forcado()
        if status:
            argumentos += ["--status", status]
        if self.mesmo_status.get():
            if not obs and not messagebox.askyesno(
                    "Reenvio sem observação",
                    "Você marcou para reenviar no mesmo status, mas a Observação está vazia.\nContinuar assim mesmo?"):
                return None
            argumentos.append("--mesmo-status")

        cfg = ler_json(CONFIG)
        cfg["observacao"] = obs
        cfg["ui_status"] = status or ""
        gravar_json(CONFIG, cfg)
        return argumentos

    def _confirmar_real(self):
        if self.simular.get():
            return True
        status = self._status_forcado() or "o status da coluna da planilha"
        texto = (f"MODO REAL: o robô vai alterar o SGE de verdade.\n\nStatus: {status}\n"
                 f"Data: {self.data.get().strip() or 'hoje'}\nObservação: {self.obs.get().strip() or '(nenhuma)'}"
                 + ("\nReenvio no mesmo status: SIM" if self.mesmo_status.get() else "") + "\n\nContinuar?")
        return messagebox.askyesno("Confirmar", texto)

    def _iniciar(self):
        if not self._salvar_planilha():
            return
        argumentos = self._opcoes_rodada()
        if argumentos is None:
            return
        linha = self.linha.get().strip()
        if linha:
            if not linha.isdigit():
                messagebox.showerror("Linha inválida", "Digite só o número da linha.")
                return
            argumentos += ["--linha", linha]
        else:
            prog = ler_json(PROGRESSO)
            if prog.get("proxima_linha") and not prog.get("concluido") and not self.simular.get():
                resposta = messagebox.askyesnocancel(
                    "Rodada interrompida",
                    f"A última rodada parou antes da linha {prog['proxima_linha']}.\n\n"
                    "Sim = continuar dessa linha\nNão = começar do início")
                if resposta is None:
                    return
                if resposta:
                    argumentos.append("--continuar")
        if self._confirmar_real():
            self._rodar(argumentos)

    def _iniciar_vigia(self):
        if not self._salvar_planilha():
            return
        argumentos = self._opcoes_rodada()
        if argumentos is None:
            return
        argumentos.append("--vigia")
        if self._confirmar_real():
            self._rodar(argumentos)

    def _continuar(self):
        prog = ler_json(PROGRESSO)
        if not prog.get("proxima_linha") or prog.get("concluido"):
            messagebox.showinfo("Continuar", "Não há rodada interrompida — use Iniciar.")
            return
        if not self._salvar_planilha():
            return
        argumentos = self._opcoes_rodada()
        if argumentos is not None and self._confirmar_real():
            self._rodar(argumentos + ["--continuar"])

    def _rodar(self, argumentos, simular=None, rodada=True):
        if self.processo:
            return
        if rodada:
            self._salvar_contas()
        if simular is None:
            simular = self.simular.get()
        if simular:
            argumentos.append("--simular")
        PARAR.unlink(missing_ok=True)
        self.pediu_parar = False
        comando = [python_console(), "-u", str(ROBO), *argumentos]
        self._escrever("\n$ BAIXA.PY " + " ".join(f'"{a}"' if " " in a or not a else a for a in argumentos) + "\n",
                       "aviso")
        ambiente = dict(os.environ, PYTHONIOENCODING="utf-8", PYTHONUTF8="1")
        self.processo = subprocess.Popen(
            comando, cwd=str(PASTA), stdout=subprocess.PIPE, stderr=subprocess.STDOUT, stdin=subprocess.DEVNULL,
            text=True, encoding="utf-8", errors="replace", env=ambiente,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
        threading.Thread(target=self._ler_processo, args=(self.processo,), daemon=True).start()
        self.status.config(text="Rodando...")
        self._botoes_rodando(True)

    def _parar(self):
        if not self.processo:
            return
        if not self.pediu_parar:
            PARAR.write_text("parar", encoding="utf-8")
            self.pediu_parar = True
            self.status.config(text="Parando: termina o aluno atual e salva a linha...")
            self.b_parar.config(text="■  Forçar parada")
            self._escrever("\n>>> Pedido para parar: o robô termina o aluno atual e salva a linha.\n", "aviso")
            return
        if messagebox.askyesno("Forçar parada",
                               "Fechar o robô agora?\nO aluno que está no meio será refeito quando continuar."):
            self._matar()

    def _matar(self):
        if self.processo:
            subprocess.run(["taskkill", "/PID", str(self.processo.pid), "/T", "/F"],
                           capture_output=True, creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))

    def _abrir_relatorio(self):
        if RELATORIO.exists():
            os.startfile(str(RELATORIO))
        else:
            messagebox.showinfo("Relatório", "Ainda não há relatório.")

    def _testar_planilha(self):
        if self._salvar_planilha():
            self._rodar(["--listar"], simular=False, rodada=False)

    # ------------------------------------------------------------------ saída do robô
    def _ler_processo(self, processo):
        for linha in processo.stdout:
            self.fila.put(linha)
        processo.wait()
        self.fila.put(None)

    def _ler_saida(self):
        try:
            while True:
                linha = self.fila.get_nowait()
                if linha is None:
                    self._terminou()
                    continue
                if linha.startswith("DETECTAR_JSON:"):
                    try:
                        self._mostrar_deteccao(json.loads(linha[len("DETECTAR_JSON:"):]))
                    except ValueError:
                        pass
                    continue
                if linha.startswith("ITENS_JSON:"):
                    try:
                        self._mostrar_itens(json.loads(linha[len("ITENS_JSON:"):]))
                    except ValueError:
                        pass
                    continue
                if linha.startswith("NOTIFICACAO_JSON:"):
                    try:
                        evt = json.loads(linha[len("NOTIFICACAO_JSON:"):])
                        self.root.bell()
                        self._escrever(
                            f"🔔 [NOTIFICAÇÃO TEMPO REAL] SGE {evt.get('sge')} — {evt.get('nome')} "
                            f"(CPF: {evt.get('cpf') or 'N/I'}): Coluna K '{evt.get('statusColunaK')}' "
                            f"➔ SGE '{evt.get('acaoSge')}'\n",
                            "ok",
                        )
                    except ValueError:
                        pass
                    continue
                self._escrever(linha, self._tag(linha))
                achado = re.match(r"\s*\[(\d+)/(\d+)\]\s*(.*)", linha)
                if achado:
                    self.status.config(text=f"Aluno {achado.group(1)} de {achado.group(2)} — {achado.group(3).strip()}")
        except queue.Empty:
            pass
        self.root.after(150, self._ler_saida)

    def _tag(self, linha):
        texto = linha.upper()
        if re.match(r"\s*\[\d+/\d+\]", linha):
            return "aluno"
        if "ERRO" in texto and "RESUMO" not in texto or "NAO ENCONTRADO" in texto or "TRACEBACK" in texto:
            return "erro"
        if ("ALTERADO" in texto or "REENVIADO" in texto) and "RESUMO" not in texto:
            return "ok"
        if "JA NO STATUS" in texto or "SEM ITEM" in texto or "SIMULADO" in texto or "RESUMO" in texto \
                or "TROCANDO" in texto or "CAIU" in texto:
            return "aviso"
        return None

    def _terminou(self):
        codigo = self.processo.returncode if self.processo else None
        self.processo = None
        PARAR.unlink(missing_ok=True)
        self.pediu_parar = False
        self._botoes_rodando(False)
        self._atualizar_progresso()
        self._carregar_status(ler_json(CONFIG))
        self.status.config(text="Parado" + (f" (saiu com código {codigo})" if codigo else ""))
        self._escrever("\n--- robô parado ---\n", "aviso")

    def _escrever(self, texto, tag=None):
        self.log.insert("end", texto, tag)
        self.log.see("end")

    def _botoes_rodando(self, rodando):
        estado_livre = "disabled" if rodando else "normal"
        for botao in (self.b_iniciar, self.b_vigia, self.b_continuar):
            botao.config(state=estado_livre)
        self.b_parar.config(state="normal" if rodando else "disabled", text="■  Parar")

    def _fechar(self):
        if self.processo:
            if not messagebox.askyesno("Robô rodando",
                                       "O robô está rodando. Parar agora e fechar?\n"
                                       "A linha onde ele parou fica salva para continuar depois."):
                return
            self._matar()
        self.root.destroy()


def main():
    root = tk.Tk()
    Janela(root)
    root.mainloop()


if __name__ == "__main__":
    main()
